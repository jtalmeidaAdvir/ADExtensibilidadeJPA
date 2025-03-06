using Primavera.Extensibility.CustomForm;
using StdBE100;
using System.Data;
using System.Windows.Forms;
using System;

namespace ADExtensibilidadeJPA
{
    public partial class QuadroControlo : CustomForm
    {
        public QuadroControlo()
        {
            InitializeComponent();

            // Adicionando o manipulador de evento no construtor do formulário
            this.Load += new EventHandler(QuadroControlo_Load);
        }

        // Método chamado quando o formulário é carregado
        private void QuadroControlo_Load(object sender, EventArgs e)
        {
            // Chamando o método DadosLista para carregar os dados no DataGridView
            DadosLista();
        }

        private void DadosLista()
        {
            try
            {
                string query = "SELECT id, Nome, CDU_EmailEnviado, CDU_DataEnvio FROM Geral_Entidade WHERE CDU_TrataSGS = 0";
                StdBELista dt = BSO.Consulta(query);

                // Criando um DataTable para armazenar os dados
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("ID", typeof(string));  // Adicionando coluna de ID
                dataTable.Columns.Add("Nome", typeof(string));
                dataTable.Columns.Add("EmailEnviadoColumn", typeof(bool)); // Ajuste o tipo conforme necessário
                dataTable.Columns.Add("DataEnvioColumn", typeof(DateTime)); // Ajuste o tipo conforme necessário

                dt.Inicio();
                while (!dt.NoFim())
                {
                    // Verificando e tratando valores nulos
                    string id = dt.Valor("id")?.ToString() ?? string.Empty;
                    string nome = dt.Valor("Nome")?.ToString() ?? string.Empty;
                    bool emailEnviado = false;

                    // Tentando converter o valor de CDU_EmailEnviado para booleano, se possível
                    if (bool.TryParse(dt.Valor("CDU_EmailEnviado")?.ToString(), out bool result))
                    {
                        emailEnviado = result;
                    }

                    DateTime dataEnvio;
                    // Tentando converter o valor de CDU_DataEnvio para DateTime, se possível
                    if (!DateTime.TryParse(dt.Valor("CDU_DataEnvio")?.ToString(), out dataEnvio))
                    {
                        dataEnvio = DateTime.MinValue; // Definir um valor padrão ou ajustar conforme necessário
                    }

                    // Adicionando os dados à tabela
                    dataTable.Rows.Add(id, nome, emailEnviado, dataEnvio);

                    dt.Seguinte();
                }

                // Associando o DataTable ao DataGridView
                dataGridView1.DataSource = dataTable;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao carregar dados: " + ex.Message);
            }
        }


        private void BT_Editar_Click(object sender, EventArgs e)
        {
            try
            {
                // Verificando se uma linha foi selecionada no DataGridView
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    // Obtendo o ID da linha selecionada
                    //string idSelecionado = dataGridView1.SelectedRows[0].Cells["ID"].Value?.ToString();

                    // Criando e exibindo o formulário "Menu" e passando o ID
                    Menu menuForm = new Menu(BSO,PSO);
                    menuForm.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Por favor, selecione uma linha para editar.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao editar: " + ex.Message);
            }
        }

    }
}
