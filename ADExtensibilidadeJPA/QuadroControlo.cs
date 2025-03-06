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
            ConfigurarInterface();
            DadosLista();
        }

        private TextBox txtFiltro;
        private Button btnFiltrar;
        private Button btnLimparFiltro;
        private Button btnFiltrarEnviados;
        private DataTable dataOriginal;

        private void ConfigurarInterface()
        {
            // Configuração do formulário principal com gradiente visual moderno
            this.BackColor = System.Drawing.Color.FromArgb(240, 242, 245);

            // Criar um painel de topo com gradiente
            System.Windows.Forms.Panel topPanel = new System.Windows.Forms.Panel
            {
                Height = 45,
                Dock = DockStyle.Top,
                BackColor = System.Drawing.Color.FromArgb(59, 89, 152)
            };
            this.Controls.Add(topPanel);

            // Adicionar título ao painel de topo
            Label lblTitulo = new Label
            {
                Text = "Gestão de Entidades",
                Font = new System.Drawing.Font("Calibri", 16F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                AutoSize = true,
                Location = new System.Drawing.Point(300, 9)
            };
            topPanel.Controls.Add(lblTitulo);

            // Mover os botões para o painel de topo e reestilizar
            topPanel.Controls.Add(BT_Editar);
            topPanel.Controls.Add(Bt_Email);
            BT_Editar.Location = new System.Drawing.Point(10, 9);
            Bt_Email.Location = new System.Drawing.Point(110, 9);

            // Adicionar controles de filtro
            System.Windows.Forms.Panel panelFiltro = new System.Windows.Forms.Panel
            {
                Height = 45,
                Dock = DockStyle.Top,
                BackColor = System.Drawing.Color.White,
                Location = new System.Drawing.Point(0, 45)
            };
            this.Controls.Add(panelFiltro);

            // Label para o filtro
            Label lblFiltro = new Label
            {
                Text = "Filtrar por Nome:",
                Font = new System.Drawing.Font("Calibri", 10F),
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                AutoSize = true,
                Location = new System.Drawing.Point(10, 14)
            };
            panelFiltro.Controls.Add(lblFiltro);

            // Textbox para o filtro
            txtFiltro = new TextBox
            {
                Location = new System.Drawing.Point(120, 12),
                Size = new System.Drawing.Size(300, 23),
                Font = new System.Drawing.Font("Calibri", 10F),
                BorderStyle = BorderStyle.FixedSingle
            };
            panelFiltro.Controls.Add(txtFiltro);

            // Botão Filtrar
            btnFiltrar = new Button
            {
                Text = "Filtrar",
                Location = new System.Drawing.Point(430, 11),
                Size = new System.Drawing.Size(80, 25),
                Font = new System.Drawing.Font("Calibri", 9.5F, System.Drawing.FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                BackColor = System.Drawing.Color.White
            };
            btnFiltrar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 89, 152);
            btnFiltrar.Click += BtnFiltrar_Click;
            panelFiltro.Controls.Add(btnFiltrar);

            // Botão Limpar Filtro
            btnLimparFiltro = new Button
            {
                Text = "Limpar",
                Location = new System.Drawing.Point(520, 11),
                Size = new System.Drawing.Size(80, 25),
                Font = new System.Drawing.Font("Calibri", 9.5F, System.Drawing.FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                BackColor = System.Drawing.Color.White
            };
            btnLimparFiltro.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 89, 152);
            btnLimparFiltro.Click += BtnLimparFiltro_Click;
            panelFiltro.Controls.Add(btnLimparFiltro);

            // Botão Filtrar Emails Enviados
            btnFiltrarEnviados = new Button
            {
                Text = "Ver Enviados",
                Location = new System.Drawing.Point(610, 11),
                Size = new System.Drawing.Size(100, 25),
                Font = new System.Drawing.Font("Calibri", 9.5F, System.Drawing.FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                BackColor = System.Drawing.Color.White
            };
            btnFiltrarEnviados.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 89, 152);
            btnFiltrarEnviados.Click += BtnFiltrarEnviados_Click;
            panelFiltro.Controls.Add(btnFiltrarEnviados);

            // Ajustar posição do DataGridView
            dataGridView1.Location = new System.Drawing.Point(10, 100);
            dataGridView1.Size = new System.Drawing.Size(780, 340);

            // Configuração avançada do DataGridView
            dataGridView1.BorderStyle = BorderStyle.None;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(59, 89, 152);
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold);
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersHeight = 40;
            dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9.5F);
            dataGridView1.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(192, 202, 221);
            dataGridView1.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;
            dataGridView1.RowTemplate.Height = 33;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 242, 245);
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView1.BackgroundColor = System.Drawing.Color.White;
            dataGridView1.GridColor = System.Drawing.Color.FromArgb(220, 220, 220);
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridView1.RowsDefaultCellStyle.Padding = new Padding(3);
            // Adicionar sombra (simulada com bordas)
            dataGridView1.BorderStyle = BorderStyle.Fixed3D;

            // Estilização dos botões
            EstilizarBotao(BT_Editar, "Editar");
            EstilizarBotao(Bt_Email, "Enviar Email");

            // Adicionar painel inferior com informações ou estatísticas
            System.Windows.Forms.Panel bottomPanel = new System.Windows.Forms.Panel
            {
                Height = 30,
                Dock = DockStyle.Bottom,
                BackColor = System.Drawing.Color.FromArgb(240, 242, 245)
            };
            this.Controls.Add(bottomPanel);

            // Adicionar Label para informação do total de registros
            Label lblInfo = new Label
            {
                Text = "Clique duas vezes em um registo para ver mais detalhes",
                Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic),
                ForeColor = System.Drawing.Color.DimGray,
                AutoSize = true,
                Location = new System.Drawing.Point(10, 8)
            };
            bottomPanel.Controls.Add(lblInfo);
        }

        private void EstilizarBotao(Button botao, string texto)
        {
            botao.FlatStyle = FlatStyle.Flat;
            botao.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(235, 235, 235);
            botao.FlatAppearance.BorderSize = 1;
            botao.BackColor = System.Drawing.Color.White;
            botao.ForeColor = System.Drawing.Color.FromArgb(59, 89, 152);
            botao.Font = new System.Drawing.Font("Calibri", 9.5F, System.Drawing.FontStyle.Bold);
            botao.Cursor = Cursors.Hand;
            botao.Text = texto;
            botao.Size = new System.Drawing.Size(90, 28);

            // Adicionar efeito de hover ao botão
            botao.MouseEnter += (s, e) => {
                botao.BackColor = System.Drawing.Color.FromArgb(59, 89, 152);
                botao.ForeColor = System.Drawing.Color.White;
            };
            botao.MouseLeave += (s, e) => {
                botao.BackColor = System.Drawing.Color.White;
                botao.ForeColor = System.Drawing.Color.FromArgb(59, 89, 152);
            };
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

                // Guardamos uma cópia dos dados originais para poder filtrar e restaurar
                dataOriginal = dataTable.Copy();
                dataGridView1.DataSource = dataTable;
                dataGridView1.Columns["ID"].Visible = false;

                // Configurando os cabeçalhos das colunas para melhor legibilidade
                dataGridView1.Columns["Nome"].HeaderText = "Nome da Entidade";
                dataGridView1.Columns["EmailEnviadoColumn"].HeaderText = "Email Enviado";
                dataGridView1.Columns["DataEnvioColumn"].HeaderText = "Data de Envio";

                // Ajustando o alinhamento e formato das colunas
                dataGridView1.Columns["EmailEnviadoColumn"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns["DataEnvioColumn"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";
                dataGridView1.Columns["DataEnvioColumn"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // Adicionando estilo condicional - destaque para emails enviados e tratamento de datas mínimas
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Verificar e formatar o campo de data
                    if (row.Cells["DataEnvioColumn"].Value != null)
                    {
                        // Verifica se o valor é uma data válida
                        if (row.Cells["DataEnvioColumn"].Value is DateTime dataEnvio)
                        {
                            if (dataEnvio == DateTime.MinValue || dataEnvio.Year == 1)
                            {
                                // Criar uma célula formatada para mostrar "Sem valor" sem alterar o tipo de dados
                                row.Cells["DataEnvioColumn"].Style.Format = null;
                                row.Cells["DataEnvioColumn"].Style.NullValue = "Sem valor";
                                row.Cells["DataEnvioColumn"].Style.ForeColor = System.Drawing.Color.Gray;
                                row.Cells["DataEnvioColumn"].Style.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Italic);
                                // Mantém o valor como DateTime para evitar problemas de formatação
                                row.Cells["DataEnvioColumn"].Value = DBNull.Value;
                            }
                        }
                    }

                    // Destacar emails enviados
                    if (row.Cells["EmailEnviadoColumn"].Value != null)
                    {
                        bool enviado = Convert.ToBoolean(row.Cells["EmailEnviadoColumn"].Value);
                        if (enviado)
                        {
                            row.Cells["EmailEnviadoColumn"].Style.ForeColor = System.Drawing.Color.Green;
                            row.Cells["EmailEnviadoColumn"].Style.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao carregar dados: " + ex.Message, "Erro",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    // Recarregar os dados para mostrar as alterações
                    DadosLista();
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

        private void BtnFiltrar_Click(object sender, EventArgs e)
        {
            try
            {
                string filtroNome = txtFiltro.Text.Trim().ToLower();

                if (string.IsNullOrEmpty(filtroNome))
                {
                    MessageBox.Show("Por favor, insira um texto para filtrar.");
                    return;
                }

                // Criar uma nova vista com o filtro
                DataTable dataFiltrada = dataOriginal.Clone();

                foreach (DataRow row in dataOriginal.Rows)
                {
                    string nome = row["Nome"].ToString().ToLower();
                    if (nome.Contains(filtroNome))
                    {
                        dataFiltrada.ImportRow(row);
                    }
                }

                dataGridView1.DataSource = dataFiltrada;

                if (dataFiltrada.Rows.Count == 0)
                {
                    MessageBox.Show("Nenhum resultado encontrado para o filtro aplicado.");
                }
                else
                {
                    // Adicionar uma mensagem no painel inferior para indicar que um filtro está ativo
                    // Você pode implementar isso adicionando um Label no ConfigurarInterface
                }
            }
            catch 
            {
                MessageBox.Show("Erro ao aplicar filtro: ");
            }
        }

        private void BtnLimparFiltro_Click(object sender, EventArgs e)
        {
            txtFiltro.Text = "";
            dataGridView1.DataSource = dataOriginal;
            // Limpar a mensagem de filtro ativo, se houver
        }

        private void BtnFiltrarEnviados_Click(object sender, EventArgs e)
        {
            try
            {
                // Criar uma nova vista com o filtro para emails enviados
                DataTable dataFiltrada = dataOriginal.Clone();

                foreach (DataRow row in dataOriginal.Rows)
                {
                    // Verificar se o email foi enviado
                    if (row["EmailEnviadoColumn"] != DBNull.Value && Convert.ToBoolean(row["EmailEnviadoColumn"]) == true)
                    {
                        dataFiltrada.ImportRow(row);
                    }
                }

                dataGridView1.DataSource = dataFiltrada;

                if (dataFiltrada.Rows.Count == 0)
                {
                    MessageBox.Show("Nenhum email enviado encontrado.");
                }
                else
                {
                    MessageBox.Show($"Exibindo {dataFiltrada.Rows.Count} registros com emails enviados.");
                }
            }
            catch 
            {
                MessageBox.Show("Erro ao filtrar emails enviados: ");
            }
        }
    }
}