using Microsoft.Office.Interop.Outlook;  // Para o Outlook
using Primavera.Extensibility.CustomForm;
using StdBE100;
using System;
using System.Data;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class QuadroControlo : CustomForm
    {
        public QuadroControlo()
        {
            InitializeComponent();
            this.Load += new EventHandler(QuadroControlo_Load);
        }

        private void QuadroControlo_Load(object sender, EventArgs e)
        {
            DadosLista();
        }

        private void DadosLista()
        {
            try
            {
                string query = "SELECT id, Nome, CDU_EmailEnviado, CDU_DataEnvio FROM Geral_Entidade WHERE CDU_TrataSGS = 0";
                StdBELista dt = BSO.Consulta(query);

                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("ID", typeof(string));
                dataTable.Columns.Add("Nome", typeof(string));
                dataTable.Columns.Add("EmailEnviadoColumn", typeof(bool));
                dataTable.Columns.Add("DataEnvioColumn", typeof(DateTime));

                dt.Inicio();
                while (!dt.NoFim())
                {
                    string id = dt.Valor("id")?.ToString() ?? string.Empty;
                    string nome = dt.Valor("Nome")?.ToString() ?? string.Empty;
                    bool emailEnviado = bool.TryParse(dt.Valor("CDU_EmailEnviado")?.ToString(), out bool result) ? result : false;
                    DateTime dataEnvio = DateTime.TryParse(dt.Valor("CDU_DataEnvio")?.ToString(), out DateTime envio) ? envio : DateTime.MinValue;

                    dataTable.Rows.Add(id, nome, emailEnviado, dataEnvio);

                    dt.Seguinte();
                }

                dataGridView1.DataSource = dataTable;
                dataGridView1.Columns["ID"].Visible = false;
            }
            catch (System.Exception ex) // Usando explicitamente System.Exception
            {
                MessageBox.Show("Erro ao carregar dados: " + ex.Message);
            }
        }

        private void BT_Editar_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    string idSelecionado = dataGridView1.SelectedRows[0].Cells["ID"].Value?.ToString();
                    Menu menuForm = new Menu(BSO, PSO, idSelecionado);
                    menuForm.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Por favor, selecione uma linha para editar.");
                }
            }
            catch (System.Exception ex) // Usando explicitamente System.Exception
            {
                MessageBox.Show("Erro ao editar: " + ex.Message);
            }
        }

        private void Bt_Email_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    string idSelecionado = dataGridView1.SelectedRows[0].Cells["ID"].Value?.ToString();
                    string nome = dataGridView1.SelectedRows[0].Cells["Nome"].Value?.ToString();

                    // Consulta para buscar o e-mail da entidade
                    string query = $@"
                SELECT ec.Email 
              FROM Geral_Entidade ge
LEFT JOIN Geral_Entidade_Contactos ec ON CAST(ge.id AS uniqueidentifier) = ec.EntidadeID
                WHERE ge.id = '{idSelecionado}'";
               
                    // Consultando a base de dados para obter o e-mail
                    StdBELista dt = BSO.Consulta(query);
                    string email = null;

                    // Se houver resultados, pegar o e-mail
                    dt.Inicio();
                    if (!dt.NoFim())
                    {
                        email = dt.Valor("Email")?.ToString(); // Obtendo o e-mail da consulta
                    }

                    // Se não houver e-mail, exibir mensagem e retornar
                    if (string.IsNullOrEmpty(email))
                    {
                        MessageBox.Show("Não há e-mail registrado para esta entidade.");
                        return;
                    }

                    // Iniciando o Outlook
                    Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                    MailItem emailItem = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                    // Definindo o assunto e o corpo do e-mail
                    emailItem.Subject = "Assunto do E-mail";
                    emailItem.Body = $"Prezado(a) {nome},\n\nEste é um e-mail de teste.\n\nAtenciosamente,\nSua Empresa";

                    // Definindo o e-mail do destinatário
                    emailItem.To = email;

                    // Enviando o e-mail
                    emailItem.Send();

                    // Atualizando os campos na tabela após o envio do e-mail
                    string updateQuery = $@"
                UPDATE Geral_Entidade 
                SET CDU_EmailEnviado = 1, CDU_DataEnvio = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}'
                WHERE id = '{idSelecionado}'";
                    BSO.DSO.ExecuteSQL(updateQuery);

                    MessageBox.Show("E-mail enviado com sucesso!");
                }
                else
                {
                    MessageBox.Show("Por favor, selecione uma linha para enviar o e-mail.");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao enviar o e-mail: " + ex.Message);
            }
        }

    }
}
