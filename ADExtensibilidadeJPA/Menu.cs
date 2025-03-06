using ErpBS100;
using Primavera.Extensibility.BusinessEntities;
using Primavera.Extensibility.CustomForm;
using StdBE100;
using StdPlatBS100;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class Menu : Form
    {
        public string _ID;
        public string IdSelecionado;
        public Menu(ErpBS100.ErpBS bSO, StdPlatBS100.StdBSInterfPub pSO, string idSelecionado)
        {
            InitializeComponent();
            ConfigurarEstiloControles();
            CriarTabelaTrabalhadores();
            CriarFormularioTrabalhador();
            CriarBotaoAdicionar();
            BSO = bSO;
            PSO = pSO;
            IdSelecionado = idSelecionado;

            if (IdSelecionado != "")
            {
                DaValores();
            }
        }

        private void CriarBotaoAdicionar()
        {
            // Verificar se o botão já existe para evitar duplicação
            if (tabPage2.Controls.ContainsKey("btnAdicionarTrabalhador")) return;

            Button btnAdicionar = new Button
            {
                Name = "btnAdicionarTrabalhador",
                Text = "Adicionar Trabalhador",
                Location = new Point(10, 530), // Garante que fica abaixo da tabela
                Size = new Size(180, 30), // Define um tamanho adequado
                BackColor = Color.LightBlue,
                Font = new Font("Arial", 10, FontStyle.Bold)
            };

            // Evento de clique para mostrar o formulário
            btnAdicionar.Click += (s, e) =>
            {
                Panel panel = tabPage2.Controls["panelFormulario"] as Panel;
                if (panel != null)
                {
                    // Limpar campos ao abrir o formulário
                    foreach (Control c in panel.Controls)
                    {
                        if (c is TextBox txt) txt.Clear();
                        else if (c is DateTimePicker dtp) dtp.Value = DateTime.Today;
                        else if (c is CheckBox chk) chk.Checked = false;
                    }
                    panel.Visible = true;
                }
            };

            // Adicionar o botão à tabPage2
            tabPage2.Controls.Add(btnAdicionar);
        }
        private void DaValores()
        {
            Dictionary<string, string> entidade = new Dictionary<string, string>();
            GetEntidadesID(ref entidade);
            if (entidade.Count > 0)
            {
                SetInfoEntidades(entidade);
            }
        }


        private void GetEntidadesID(ref Dictionary<string, string> entidade)
        {
            // Consulta SQL para pegar os dados
            var query = $@"SELECT * FROM Geral_Entidade WHERE CDU_TrataSGS = 0 AND Id='{IdSelecionado}'";
            var dados = BSO.Consulta(query);

            // Iniciando a leitura dos dados
            dados.Inicio();

            // Verificando se há resultados
            if (dados.NumLinhas() > 0)
            {
                // Definindo as colunas esperadas na consulta
                string[] colunas = new string[] { "Codigo", "Nome", "NIPC", "AlvaraNumero", "AlvaraValidade", "CDU_NaoDivFinancas",
                                          "CDU_NaoDivSegSocial", "CDU_FolhaPagSegSocial", "CDU_ReciboApoliceAT",
                                          "CDU_ReciboRC", "CDU_Caminho", "CDU_ReciboPagSegSocial", "CDU_ApoliceAT",
                                          "CDU_ApoliceRC", "CDU_HorarioTrabalho", "CDU_DecTrabIlegais",
                                          "CDU_DecRespEstaleiro", "CDU_DecConhecimPSS", "Morada", "Localidade",
                                          "CodPostal", "CodPostalLocal", "EntidadeId", "id", "CDU_AnexoFinancas",
                                          "CDU_AnexoSegSocial", "CDU_FolhaPag" };

                // Iterando sobre as linhas dos dados
                for (int i = 0; i < dados.NumLinhas(); i++)
                {
                    // Preenchendo o dicionário com os valores de cada coluna
                    foreach (var coluna in colunas)
                    {
                        // Obtendo o valor de cada coluna para o tipo string e armazenando no dicionário
                        var valor = dados.DaValor<string>(coluna);
                        entidade[coluna] = valor; // Atribui o valor à chave correspondente
                    }

                    // Avançando para a próxima linha de dados
                    dados.Seguinte();
                }
            }
        }

        private void ConfigurarEstiloControles()
        {
            // Configuração das cores dos controles para uma aparência mais moderna
            this.BackColor = System.Drawing.Color.WhiteSmoke;

            // Configurar estilo do DataGridView
            dataGridView1.BorderStyle = BorderStyle.None;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;

            // Destacar botões
            BTF4.FlatStyle = FlatStyle.Flat;
            btnGravarObra.FlatStyle = FlatStyle.Flat;

            // Configurar os botões para anexar documentos específicos
            btnAnexoFinancas.FlatStyle = FlatStyle.Flat;
            btnAnexoFinancas.BackColor = Color.LightBlue;
            btnAnexoSegSocial.FlatStyle = FlatStyle.Flat;
            btnAnexoSegSocial.BackColor = Color.LightBlue;

            // Configurar painéis
            AlertaValidadeAlvara.BackColor = Color.Red;

            // Configurar valores iniciais para os DateTimePickers
            // Se a data atual não for apropriada como valor padrão, pode definir um valor mínimo
            TXT_NaoDivFinancas.Value = DateTime.Today;
            TXT_NaoDivSegSocial.Value = DateTime.Today;
            TXT_FolhaPagSegSocial.Value = DateTime.Today;
            TXT_AlvaraValidade.Value = DateTime.Today;

            // Permitir limpar as datas (opcional)
            TXT_NaoDivFinancas.ShowCheckBox = true;
            TXT_NaoDivSegSocial.ShowCheckBox = true;
            TXT_FolhaPagSegSocial.ShowCheckBox = true;
            TXT_AlvaraValidade.ShowCheckBox = true;

            // Definir o formato de exibição para mostrar a data completa
            TXT_NaoDivFinancas.Format = DateTimePickerFormat.Short;
            TXT_NaoDivSegSocial.Format = DateTimePickerFormat.Short;
            TXT_FolhaPagSegSocial.Format = DateTimePickerFormat.Short;
        }
        private void CriarTabelaTrabalhadores()
        {
            // Criar DataGridView com estilo moderno
            DataGridView dgvTrabalhadores = new DataGridView
            {
                Name = "dgvTrabalhadores",
                Size = new Size(680, 250),
                Location = new Point(8, 10),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackgroundColor = System.Drawing.Color.White,
                GridColor = System.Drawing.Color.LightGray,
                EnableHeadersVisualStyles = false,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = System.Drawing.Color.FromArgb(59, 89, 152),
                    ForeColor = System.Drawing.Color.White,
                    Font = new Font("Calibri", 9F, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                ColumnHeadersHeight = 30,
                RowTemplate = { Height = 25 },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = System.Drawing.Color.FromArgb(240, 242, 245)
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("Calibri", 9F),
                    SelectionBackColor = System.Drawing.Color.FromArgb(192, 202, 221),
                    SelectionForeColor = System.Drawing.Color.Black
                },
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            };

            // Adicionar colunas conforme imagem de referência
            dgvTrabalhadores.Columns.Add("Empresa", "Empresa");
            dgvTrabalhadores.Columns.Add("CategoriaFuncao", "Categoria/Função");
            dgvTrabalhadores.Columns.Add("Contribuinte", "Contribuinte");
            dgvTrabalhadores.Columns.Add("SegurancaSocial", "Segurança Social");
            dgvTrabalhadores.Columns.Add("CCidadaoNTitulo", "CC/Passaporte/T.Res./República");
            dgvTrabalhadores.Columns.Add("FichaMedica", "Ficha de Aptidão Médica");
            dgvTrabalhadores.Columns.Add("ContratoMesa", "Contrato de Mesa");
            dgvTrabalhadores.Columns.Add("FormacaoTrabalhador", "Trabalhador Estrangeiro");
            dgvTrabalhadores.Columns.Add("FormacaoInformacao", "Formação/Informação");
            dgvTrabalhadores.Columns.Add("EPIs", "EPIs");
            dgvTrabalhadores.Columns.Add("EntradaObra", "Entrada Obra");
            dgvTrabalhadores.Columns.Add("SaidaObra", "Saída Obra");
            dgvTrabalhadores.Columns.Add("AutorizacaoEntrada", "Autorização de Entrada em Obra");

            // Configurar largura das colunas
            dgvTrabalhadores.Columns["Empresa"].Width = 120;
            dgvTrabalhadores.Columns["CategoriaFuncao"].Width = 80;
            dgvTrabalhadores.Columns["Contribuinte"].Width = 90;
            dgvTrabalhadores.Columns["SegurancaSocial"].Width = 90;
            dgvTrabalhadores.Columns["CCidadaoNTitulo"].Width = 130;
            dgvTrabalhadores.Columns["FichaMedica"].Width = 80;
            dgvTrabalhadores.Columns["ContratoMesa"].Width = 80;
            dgvTrabalhadores.Columns["FormacaoTrabalhador"].Width = 80;
            dgvTrabalhadores.Columns["FormacaoInformacao"].Width = 90;
            dgvTrabalhadores.Columns["EPIs"].Width = 60;
            dgvTrabalhadores.Columns["EntradaObra"].Width = 80;
            dgvTrabalhadores.Columns["SaidaObra"].Width = 80;
            dgvTrabalhadores.Columns["AutorizacaoEntrada"].Width = 120;

            // Adicionar botões de ação
            DataGridViewButtonColumn btnEditar = new DataGridViewButtonColumn
            {
                Name = "Editar",
                HeaderText = "Editar",
                Text = "✏️",
                UseColumnTextForButtonValue = true,
                Width = 50
            };

            DataGridViewButtonColumn btnRemover = new DataGridViewButtonColumn
            {
                Name = "Remover",
                HeaderText = "Remover",
                Text = "❌",
                UseColumnTextForButtonValue = true,
                Width = 60
            };

            dgvTrabalhadores.Columns.Add(btnEditar);
            dgvTrabalhadores.Columns.Add(btnRemover);

            // Configurar scrollbar
            dgvTrabalhadores.ScrollBars = ScrollBars.Both;

            // Adicionar evento para edição e remoção
            dgvTrabalhadores.CellClick += dgvTrabalhadores_CellClick;

            // Adicionar à tabPage2
            tabPage2.Controls.Add(dgvTrabalhadores);

            // Adicionar label de título
            Label lblTitulo = new Label
            {
                Text = "Lista de Trabalhadores",
                Font = new Font("Calibri", 12F, FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                AutoSize = true,
                Location = new Point(8, 270)
            };
            tabPage2.Controls.Add(lblTitulo);
        }
        private void CriarFormularioTrabalhador()
        {
            // Painel principal com gradiente
            Panel panelFormulario = new Panel
            {
                Name = "panelFormulario",
                Size = new Size(680, 340),
                Location = new Point(8, 290),
                Visible = false,
                BackColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = true,
            };

            // Barra de título
            Panel titleBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 35,
                BackColor = System.Drawing.Color.FromArgb(59, 89, 152)
            };
            Label titleLabel = new Label
            {
                Text = "Dados do Trabalhador",
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Calibri", 12F, FontStyle.Bold),
                Location = new Point(10, 7),
                AutoSize = true
            };
            titleBar.Controls.Add(titleLabel);
            panelFormulario.Controls.Add(titleBar);

            // Área de conteúdo
            Panel contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15)
            };
            panelFormulario.Controls.Add(contentPanel);

            // Primeira coluna - Dados básicos
            int yPos = 50;
            int labelWidth = 130;
            int controlWidth = 180;
            int controlHeight = 25;
            int spacing = 30;

            // Empresa
            Label lblEmpresa = new Label
            {
                Text = "Empresa:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            TextBox txtEmpresa = new TextBox
            {
                Name = "txtEmpresa",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight - 5,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Categoria/Função
            Label lblFuncao = new Label
            {
                Text = "Categoria/Função:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            TextBox txtFuncao = new TextBox
            {
                Name = "txtFuncao",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight - 5,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Contribuinte
            Label lblContribuinte = new Label
            {
                Text = "Contribuinte:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            TextBox txtContribuinte = new TextBox
            {
                Name = "txtContribuinte",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight - 5,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Segurança Social
            Label lblSegurancaSocial = new Label
            {
                Text = "Segurança Social:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            TextBox txtSegurancaSocial = new TextBox
            {
                Name = "txtSegurancaSocial",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight - 5,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // CC/Passaporte
            Label lblCCidadao = new Label
            {
                Text = "CC/Passaporte:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            TextBox txtCCidadao = new TextBox
            {
                Name = "txtCCidadao",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight - 5,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Ficha Médica
            Label lblFichaMedica = new Label
            {
                Text = "Ficha Médica:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbFichaMedica = new ComboBox
            {
                Name = "cmbFichaMedica",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbFichaMedica.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbFichaMedica.SelectedIndex = 0;

            // Segunda coluna
            yPos = 50;
            int col2X = 340;

            // Contrato Mesa
            Label lblContratoMesa = new Label
            {
                Text = "Contrato Mesa:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbContratoMesa = new ComboBox
            {
                Name = "cmbContratoMesa",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbContratoMesa.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbContratoMesa.SelectedIndex = 0;
            yPos += spacing;

            // Trabalhador Estrangeiro
            Label lblTrabEstrangeiro = new Label
            {
                Text = "Trab. Estrangeiro:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbTrabEstrangeiro = new ComboBox
            {
                Name = "cmbTrabEstrangeiro",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbTrabEstrangeiro.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbTrabEstrangeiro.SelectedIndex = 0;
            yPos += spacing;

            // Formação/Informação
            Label lblFormacaoInfo = new Label
            {
                Text = "Formação/Info:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbFormacaoInfo = new ComboBox
            {
                Name = "cmbFormacaoInfo",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbFormacaoInfo.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbFormacaoInfo.SelectedIndex = 0;
            yPos += spacing;

            // EPIs
            Label lblEPIs = new Label
            {
                Text = "EPIs:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbEPIs = new ComboBox
            {
                Name = "cmbEPIs",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbEPIs.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbEPIs.SelectedIndex = 0;
            yPos += spacing;

            // Entrada Obra
            Label lblEntrada = new Label
            {
                Text = "Entrada Obra:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpEntrada = new DateTimePicker
            {
                Name = "dtpEntrada",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Saída Obra
            Label lblSaida = new Label
            {
                Text = "Saída Obra:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpSaida = new DateTimePicker
            {
                Name = "dtpSaida",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };

            // Autorização de Entrada
            CheckBox chkAutorizado = new CheckBox
            {
                Text = "Autorização de Entrada em Obra",
                Name = "chkAutorizado",
                Location = new Point(10, 260),
                Font = new Font("Calibri", 9.5F),
                AutoSize = true
            };

            // Área dos botões
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                BackColor = System.Drawing.Color.FromArgb(240, 242, 245)
            };

            Button btnSalvar = new Button
            {
                Text = "Guardar",
                Location = new Point(panelFormulario.Width - 200, 13),
                Size = new Size(90, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(59, 89, 152),
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Calibri", 9.5F, FontStyle.Bold)
            };
            btnSalvar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 89, 152);
            btnSalvar.Click += btnSalvar_Click;

            Button btnCancelar = new Button
            {
                Text = "Cancelar",
                Location = new Point(panelFormulario.Width - 100, 13),
                Size = new Size(90, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                Font = new Font("Calibri", 9.5F, FontStyle.Bold)
            };
            btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 89, 152);
            btnCancelar.Click += (s, e) => panelFormulario.Visible = false;

            buttonPanel.Controls.Add(btnSalvar);
            buttonPanel.Controls.Add(btnCancelar);
            panelFormulario.Controls.Add(buttonPanel);

            // Adicionar todos os controles ao painel de conteúdo
            contentPanel.Controls.AddRange(new Control[] {
                lblEmpresa, txtEmpresa,
                lblFuncao, txtFuncao,
                lblContribuinte, txtContribuinte,
                lblSegurancaSocial, txtSegurancaSocial,
                lblCCidadao, txtCCidadao,
                lblFichaMedica, cmbFichaMedica,
                lblContratoMesa, cmbContratoMesa,
                lblTrabEstrangeiro, cmbTrabEstrangeiro,
                lblFormacaoInfo, cmbFormacaoInfo,
                lblEPIs, cmbEPIs,
                lblEntrada, dtpEntrada,
                lblSaida, dtpSaida,
                chkAutorizado
            });

            // Adicionar o formulário à tabPage2
            tabPage2.Controls.Add(panelFormulario);
        }
        private void btnSalvar_Click(object sender, EventArgs e)
        {
            Panel panel = tabPage2.Controls["panelFormulario"] as Panel;
            Panel contentPanel = panel.Controls[1] as Panel;

            // Obter referências para todos os controles
            TextBox txtEmpresa = contentPanel.Controls["txtEmpresa"] as TextBox;
            TextBox txtFuncao = contentPanel.Controls["txtFuncao"] as TextBox;
            TextBox txtContribuinte = contentPanel.Controls["txtContribuinte"] as TextBox;
            TextBox txtSegurancaSocial = contentPanel.Controls["txtSegurancaSocial"] as TextBox;
            TextBox txtCCidadao = contentPanel.Controls["txtCCidadao"] as TextBox;
            ComboBox cmbFichaMedica = contentPanel.Controls["cmbFichaMedica"] as ComboBox;
            ComboBox cmbContratoMesa = contentPanel.Controls["cmbContratoMesa"] as ComboBox;
            ComboBox cmbTrabEstrangeiro = contentPanel.Controls["cmbTrabEstrangeiro"] as ComboBox;
            ComboBox cmbFormacaoInfo = contentPanel.Controls["cmbFormacaoInfo"] as ComboBox;
            ComboBox cmbEPIs = contentPanel.Controls["cmbEPIs"] as ComboBox;
            DateTimePicker dtpEntrada = contentPanel.Controls["dtpEntrada"] as DateTimePicker;
            DateTimePicker dtpSaida = contentPanel.Controls["dtpSaida"] as DateTimePicker;
            CheckBox chkAutorizado = contentPanel.Controls["chkAutorizado"] as CheckBox;

            // Adicionar dados à tabela
            DataGridView dgv = tabPage2.Controls["dgvTrabalhadores"] as DataGridView;

            // Criar uma nova linha com todos os campos
            DataGridViewRow row = new DataGridViewRow();
            dgv.Rows.Add(
                txtEmpresa.Text,
                txtFuncao.Text,
                txtContribuinte.Text,
                txtSegurancaSocial.Text,
                txtCCidadao.Text,
                cmbFichaMedica.Text,
                cmbContratoMesa.Text,
                cmbTrabEstrangeiro.Text,
                cmbFormacaoInfo.Text,
                cmbEPIs.Text,
                dtpEntrada.Value.ToShortDateString(),
                dtpSaida.Value.ToShortDateString(),
                chkAutorizado.Checked ? "Sim" : "Não"
            );

            // Aplicar estilo à linha adicionada
            int lastRowIndex = dgv.Rows.Count - 1;
            if (lastRowIndex % 2 == 0)
            {
                dgv.Rows[lastRowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.White;
            }
            else
            {
                dgv.Rows[lastRowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 242, 245);
            }

            // Ocultar o formulário
            panel.Visible = false;

            // Mostrar mensagem de sucesso
            MessageBox.Show("Trabalhador adicionado com sucesso!", "Sucesso",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void dgvTrabalhadores_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;

            if (e.RowIndex >= 0)
            {
                if (dgv.Columns[e.ColumnIndex].Name == "Editar")
                {
                    Panel panel = tabPage2.Controls["panelFormulario"] as Panel;
                    Panel contentPanel = panel.Controls[1] as Panel;

                    // Obter referências para todos os controles
                    TextBox txtEmpresa = contentPanel.Controls["txtEmpresa"] as TextBox;
                    TextBox txtFuncao = contentPanel.Controls["txtFuncao"] as TextBox;
                    TextBox txtContribuinte = contentPanel.Controls["txtContribuinte"] as TextBox;
                    TextBox txtSegurancaSocial = contentPanel.Controls["txtSegurancaSocial"] as TextBox;
                    TextBox txtCCidadao = contentPanel.Controls["txtCCidadao"] as TextBox;
                    ComboBox cmbFichaMedica = contentPanel.Controls["cmbFichaMedica"] as ComboBox;
                    ComboBox cmbContratoMesa = contentPanel.Controls["cmbContratoMesa"] as ComboBox;
                    ComboBox cmbTrabEstrangeiro = contentPanel.Controls["cmbTrabEstrangeiro"] as ComboBox;
                    ComboBox cmbFormacaoInfo = contentPanel.Controls["cmbFormacaoInfo"] as ComboBox;
                    ComboBox cmbEPIs = contentPanel.Controls["cmbEPIs"] as ComboBox;
                    DateTimePicker dtpEntrada = contentPanel.Controls["dtpEntrada"] as DateTimePicker;
                    DateTimePicker dtpSaida = contentPanel.Controls["dtpSaida"] as DateTimePicker;
                    CheckBox chkAutorizado = contentPanel.Controls["chkAutorizado"] as CheckBox;

                    // Preencher formulário com os dados da linha
                    txtEmpresa.Text = dgv.Rows[e.RowIndex].Cells["Empresa"].Value?.ToString() ?? "";
                    txtFuncao.Text = dgv.Rows[e.RowIndex].Cells["CategoriaFuncao"].Value?.ToString() ?? "";
                    txtContribuinte.Text = dgv.Rows[e.RowIndex].Cells["Contribuinte"].Value?.ToString() ?? "";
                    txtSegurancaSocial.Text = dgv.Rows[e.RowIndex].Cells["SegurancaSocial"].Value?.ToString() ?? "";
                    txtCCidadao.Text = dgv.Rows[e.RowIndex].Cells["CCidadaoNTitulo"].Value?.ToString() ?? "";

                    // Selecionar os itens nas ComboBoxes
                    SelectComboBoxItem(cmbFichaMedica, dgv.Rows[e.RowIndex].Cells["FichaMedica"].Value?.ToString());
                    SelectComboBoxItem(cmbContratoMesa, dgv.Rows[e.RowIndex].Cells["ContratoMesa"].Value?.ToString());
                    SelectComboBoxItem(cmbTrabEstrangeiro, dgv.Rows[e.RowIndex].Cells["FormacaoTrabalhador"].Value?.ToString());
                    SelectComboBoxItem(cmbFormacaoInfo, dgv.Rows[e.RowIndex].Cells["FormacaoInformacao"].Value?.ToString());
                    SelectComboBoxItem(cmbEPIs, dgv.Rows[e.RowIndex].Cells["EPIs"].Value?.ToString());

                    // Configurar DateTimePickers
                    string entradaStr = dgv.Rows[e.RowIndex].Cells["EntradaObra"].Value?.ToString();
                    if (DateTime.TryParse(entradaStr, out DateTime entradaDate))
                    {
                        dtpEntrada.Value = entradaDate;
                    }

                    string saidaStr = dgv.Rows[e.RowIndex].Cells["SaidaObra"].Value?.ToString();
                    if (DateTime.TryParse(saidaStr, out DateTime saidaDate))
                    {
                        dtpSaida.Value = saidaDate;
                    }

                    // Configurar CheckBox
                    string autorizadoStr = dgv.Rows[e.RowIndex].Cells["AutorizacaoEntrada"].Value?.ToString();
                    chkAutorizado.Checked = autorizadoStr == "Sim";

                    // Mostrar o formulário
                    panel.Visible = true;
                    panel.BringToFront();
                }
                else if (dgv.Columns[e.ColumnIndex].Name == "Remover")
                {
                    DialogResult result = MessageBox.Show(
                        "Tem certeza que deseja remover este trabalhador?",
                        "Confirmar Remoção",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        dgv.Rows.RemoveAt(e.RowIndex);
                        MessageBox.Show("Trabalhador removido com sucesso!", "Sucesso",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        // Método auxiliar para selecionar item na ComboBox
        private void SelectComboBoxItem(ComboBox comboBox, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                comboBox.SelectedIndex = 0;
                return;
            }

            for (int i = 0; i < comboBox.Items.Count; i++)
            {
                if (comboBox.Items[i].ToString() == value)
                {
                    comboBox.SelectedIndex = i;
                    return;
                }
            }

            comboBox.SelectedIndex = 0;
        }
        private void AtualizarLabelsAnexos()
        {
            // Atualiza o texto nos labels que mostram os anexos específicos
            lblAnexoFinancas.Text = string.IsNullOrEmpty(caminhoAnexoFinancas) ?
                "Nenhum anexo" : System.IO.Path.GetFileName(caminhoAnexoFinancas);

            lblAnexoSegSocial.Text = string.IsNullOrEmpty(caminhoAnexoSegSocial) ?
                "Nenhum anexo" : System.IO.Path.GetFileName(caminhoAnexoSegSocial);

            lblFolhaPagSS.Text = string.IsNullOrEmpty(caminhoAnexoFolhaPag) ?
                "Nenhum anexo" : System.IO.Path.GetFileName(caminhoAnexoFolhaPag);
        }

        private void BTF4_Click(object sender, EventArgs e)
        {
            CarregarDadosEntidade();
        }

        private void CarregarDadosEntidade()
        {
            Dictionary<string, string> entidade = new Dictionary<string, string>();
            GetEntidades(ref entidade);

            if (entidade.Count > 0)
            {
                SetInfoEntidades(entidade);
            }
        }

        // Variáveis para armazenar os caminhos dos anexos específicos
        private string caminhoAnexoFinancas = "";
        private string caminhoAnexoSegSocial = "";
        private string caminhoAnexoFolhaPag = "";

        public ErpBS BSO { get; private set; }
        public StdBSInterfPub PSO { get; private set; }

        private void SetInfoEntidades(Dictionary<string, string> entidade)
        {
            _ID = entidade["id"];
            TXT_Codigo.Text = entidade["Codigo"];
            TXT_Nome.Text = entidade["Nome"];
            TXT_Contribuinte.Text = entidade["NIPC"];
            TXT_Alvara.Text = entidade["AlvaraNumero"];

            // Tratamento da validade do alvará como um DateTimePicker
            if (!string.IsNullOrEmpty(entidade["AlvaraValidade"]))
            {
                try
                {
                    TXT_AlvaraValidade.Value = Convert.ToDateTime(entidade["AlvaraValidade"]);
                    TXT_AlvaraValidade.Checked = true;
                }
                catch
                {
                    TXT_AlvaraValidade.Checked = false;
                }
            }
            else
            {
                TXT_AlvaraValidade.Checked = false;
            }

            // Converter strings para DateTime para os campos de data
            if (!string.IsNullOrEmpty(entidade["CDU_NaoDivFinancas"]))
            {
                TXT_NaoDivFinancas.Value = Convert.ToDateTime(entidade["CDU_NaoDivFinancas"]);
                TXT_NaoDivFinancas.Checked = true;
            }
            else
            {
                TXT_NaoDivFinancas.Checked = false;
            }

            if (!string.IsNullOrEmpty(entidade["CDU_NaoDivSegSocial"]))
            {
                TXT_NaoDivSegSocial.Value = Convert.ToDateTime(entidade["CDU_NaoDivSegSocial"]);
                TXT_NaoDivSegSocial.Checked = true;
            }
            else
            {
                TXT_NaoDivSegSocial.Checked = false;
            }

            if (!string.IsNullOrEmpty(entidade["CDU_FolhaPagSegSocial"]))
            {
                TXT_FolhaPagSegSocial.Value = Convert.ToDateTime(entidade["CDU_FolhaPagSegSocial"]);
                TXT_FolhaPagSegSocial.Checked = true;
            }
            else
            {
                TXT_FolhaPagSegSocial.Checked = false;
            }
            TXT_ReciboApoliceAT.Text = entidade["CDU_ReciboApoliceAT"];
            TXT_ReciboRC.Text = entidade["CDU_ReciboRC"];

            // Recupera o caminho da pasta
            txtCaminhoPasta.Text = entidade["CDU_Caminho"];

            // Recupera os caminhos dos anexos específicos
            caminhoAnexoFinancas = entidade["CDU_AnexoFinancas"] ?? "";
            caminhoAnexoSegSocial = entidade["CDU_AnexoSegSocial"] ?? "";
            caminhoAnexoFolhaPag = entidade["CDU_FolhaPag"] ?? "";

            // Atualiza os labels de anexos específicos
            AtualizarLabelsAnexos();

            // Atualiza a lista de documentos ao carregar a entidade
            AtualizarListaDocumentos();

            // Recupera os valores do banco de dados
            string reciboPagSegSocial = entidade["CDU_ReciboPagSegSocial"];
            string apoliceAT = entidade["CDU_ApoliceAT"];
            string apoliceRC = entidade["CDU_ApoliceRC"];
            string horarioTrabalho = entidade["CDU_HorarioTrabalho"];
            string decTrabIlegais = entidade["CDU_DecTrabIlegais"];
            string decRespEstaleiro = entidade["CDU_DecRespEstaleiro"];
            string decConhecimPSS = entidade["CDU_DecConhecimPSS"];

            PreencherComboBox(cb_ReciboPagSegSocial, reciboPagSegSocial);
            PreencherComboBox(cb_ApoliceAT, apoliceAT);
            PreencherComboBox(cb_ApoliceRC, apoliceRC);
            PreencherComboBox(cb_HorarioTrabalho, horarioTrabalho);
            PreencherComboBox(cb_DecTrabIlegais, decTrabIlegais);
            PreencherComboBox(cb_DecRespEstaleiro, decRespEstaleiro);
            PreencherComboBox(cb_DecConhecimPSS, decConhecimPSS);

            var moradaCompleta = $"{entidade["Morada"]}, {entidade["Localidade"]}, {entidade["CodPostal"]}, {entidade["CodPostalLocal"]}";

            if (moradaCompleta == ", , , ")
            {
                moradaCompleta = "";
            }
            else
            {
                TXT_Sede.Text = moradaCompleta;
            }

            CarregarObrasComboBox(entidade);
        }

        private void cb_Obras_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Obras.SelectedItem is KeyValuePair<string, string> obraSelecionada)
            {
                string codigoObraSelecionada = obraSelecionada.Key;
                string queryGetObras = $@"SELECT * FROM TDU_AD_Obras 
                                  WHERE CDU_Obra = '{codigoObraSelecionada}'";

                var DBObras = BSO.Consulta(queryGetObras);

                dataGridView1.Rows.Clear();

                if (DBObras.NumLinhas() > 0)
                {
                    DBObras.Inicio();

                    while (!DBObras.NoFim())
                    {
                        dataGridView1.Rows.Add(
                            DBObras.DaValor<string>("CDU_EntradaObra"),
                            DBObras.DaValor<string>("CDU_SaidaObra"),
                            DBObras.DaValor<string>("CDU_ContratoSubempreitada"),
                            DBObras.DaValor<int>("CDU_AutorizacaoEntrada") == 1
                        );

                        int lastRowIndex = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[lastRowIndex].DefaultCellStyle.BackColor = Color.LightYellow;

                        DBObras.Seguinte();
                    }
                }
            }
        }

        private void CarregarObrasComboBox(Dictionary<string, string> entidade)
        {
            var BDObras = GetObrasSumbempreiteiro(entidade["EntidadeId"]);

            cb_Obras.Items.Clear(); // Limpa antes de adicionar novos itens

            while (!BDObras.NoFim())
            {
                string codigo = BDObras.DaValor<string>("Codigo");
                string descricao = BDObras.DaValor<string>("Descricao");

                // Adiciona um item ao ComboBox
                cb_Obras.Items.Add(new KeyValuePair<string, string>(codigo, $"{codigo} - {descricao}"));

                BDObras.Seguinte();
            }

            cb_Obras.DisplayMember = "Value"; // O que será exibido
            cb_Obras.ValueMember = "Key"; // O valor interno
        }

        private StdBELista GetObrasSumbempreiteiro(string entidadeId)
        {
            var query = $@"SELECT * FROM COP_Obras 
                            WHERE Tipo = 'S' AND EntidadeIDA = '{entidadeId}'";
            var BDObras = BSO.Consulta(query);
            return BDObras;
        }

        private void PreencherComboBox(ComboBox comboBox, string valorBanco)
        {
            // Defina a coleção de opções que você deseja no ComboBox
            var options = new List<string> { "C", "N/C", "N/A" };

            // Define o DataSource do ComboBox
            comboBox.DataSource = options;

            // Verifica se o valor retornado é NULL ou vazio
            if (string.IsNullOrEmpty(valorBanco))
            {
                // Se for NULL ou vazio, seleciona a opção "N/A" (terceira opção)
                comboBox.SelectedItem = options[2]; // "N/A"
            }
            else
            {
                // Caso contrário, verifica se o valor está na lista
                if (options.Contains(valorBanco))
                {
                    comboBox.SelectedItem = valorBanco;
                }
                else
                {
                    // Se o valor não estiver na lista, pode-se definir um valor padrão (por exemplo, "N/A")
                    comboBox.SelectedItem = options[2]; // "N/A"
                }
            }
        }

        private void btnSelecionarPasta_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Selecione a pasta para os documentos fiscais";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtCaminhoPasta.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void btnAnexarDocumento_Click(object sender, EventArgs e)
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro uma pasta para guardar os documentos.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png|Documentos Word|*.doc;*.docx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Selecionar Documentos para Anexar";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        foreach (string sourceFile in openFileDialog.FileNames)
                        {
                            string fileName = System.IO.Path.GetFileName(sourceFile);
                            string destFile = System.IO.Path.Combine(txtCaminhoPasta.Text, fileName);

                            // Verifica se o arquivo já existe
                            if (System.IO.File.Exists(destFile))
                            {
                                DialogResult result = MessageBox.Show(
                                    $"O arquivo {fileName} já existe na pasta de destino. Deseja substituí-lo?",
                                    "Arquivo já existe",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question);

                                if (result == DialogResult.No)
                                    continue;
                            }

                            // Copia o arquivo para a pasta de destino
                            System.IO.File.Copy(sourceFile, destFile, true);
                        }

                        MessageBox.Show("Documentos anexados com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos (se implementada)
                        AtualizarListaDocumentos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documentos: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void AtualizarListaDocumentos()
        {

            // Verifica se o caminho da pasta existe
            if (string.IsNullOrEmpty(txtCaminhoPasta.Text) || !System.IO.Directory.Exists(txtCaminhoPasta.Text))
                return;

            try
            {
                // Obtém todos os arquivos da pasta
                string[] arquivos = System.IO.Directory.GetFiles(txtCaminhoPasta.Text);

                // Adiciona cada arquivo à lista
                foreach (string arquivo in arquivos)
                {
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao listar documentos: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetEntidades(ref Dictionary<string, string> entidade)
        {
            string NomeLista = "Entidades";
            string Campos = "Codigo,Nome, NIPC, AlvaraNumero, AlvaraValidade, CDU_NaoDivFinancas, CDU_NaoDivSegSocial, CDU_FolhaPagSegSocial, CDU_ReciboApoliceAT, CDU_ReciboRC, CDU_Caminho, CDU_ReciboPagSegSocial, CDU_ApoliceAT, CDU_ApoliceRC, CDU_HorarioTrabalho, CDU_DecTrabIlegais, CDU_DecRespEstaleiro, CDU_DecConhecimPSS, Morada, Localidade ,CodPostal,CodPostalLocal,EntidadeId,id,CDU_AnexoFinancas,CDU_AnexoSegSocial,CDU_FolhaPag";
            string Tabela = "Geral_Entidade (NOLOCK)";
            string Where = "CDU_TrataSGS = 0";
            string CamposF4 = "Codigo,Nome, NIPC, AlvaraNumero, AlvaraValidade, CDU_NaoDivFinancas, CDU_NaoDivSegSocial, CDU_FolhaPagSegSocial, CDU_ReciboApoliceAT, CDU_ReciboRC, CDU_Caminho, CDU_ReciboPagSegSocial, CDU_ApoliceAT, CDU_ApoliceRC, CDU_HorarioTrabalho, CDU_DecTrabIlegais, CDU_DecRespEstaleiro, CDU_DecConhecimPSS, Morada, Localidade ,CodPostal,CodPostalLocal,EntidadeId,id,CDU_AnexoFinancas,CDU_AnexoSegSocial,CDU_FolhaPag";
            string orderby = "Codigo, Nome";

            List<string> ResQuery = new List<string>();

            OpenF4List(Campos, Tabela, Where, CamposF4, orderby, NomeLista, ref ResQuery);

            if (ResQuery.Count > 0)
            {
                string[] colunas = CamposF4.Split(',');
                for (int i = 0; i < colunas.Length; i++)
                {
                    if (i < ResQuery.Count)
                    {
                        entidade[colunas[i].Trim()] = ResQuery[i].ToString();
                    }
                }
            }
        }

        private void OpenF4List(string campos, string tabela, string where, string camposF4, string orderby, string nomeLista, ref List<string> resQuery)
        {
            string strSQL = "select distinct " + campos + " FROM " + tabela;

            if (where.Length > 0)
            {
                strSQL += " WHERE " + where;
            }

            strSQL += " Order by " + orderby;
            string result = Convert.ToString(PSO.Listas.GetF4SQL(nomeLista, strSQL, camposF4));

            if (!string.IsNullOrEmpty(result))
            {
                string[] itemQuery = result.Split('\t');
                resQuery.AddRange(itemQuery);
            }
        }

        private void btnAnexoFinancas_Click(object sender, EventArgs e)
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro uma pasta para guardar os documentos.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;
                openFileDialog.Title = "Selecionar Documento da Certidão de Não Dívida às Finanças";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "NaoDivFinancas_" + TXT_Codigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
                                          System.IO.Path.GetExtension(sourceFile);
                        string destFile = System.IO.Path.Combine(txtCaminhoPasta.Text, fileName);

                        // Verifica se o arquivo já existe
                        if (System.IO.File.Exists(destFile))
                        {
                            DialogResult result = MessageBox.Show(
                                $"O arquivo {fileName} já existe na pasta de destino. Deseja substituí-lo?",
                                "Arquivo já existe",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (result == DialogResult.No)
                                return;
                        }

                        // Copia o arquivo para a pasta de destino
                        System.IO.File.Copy(sourceFile, destFile, true);

                        // Atualiza o caminho do anexo
                        caminhoAnexoFinancas = destFile;

                        // Atualiza o label
                        AtualizarLabelsAnexos();

                        MessageBox.Show("Documento anexado com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos
                        AtualizarListaDocumentos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnAnexoSegSocial_Click(object sender, EventArgs e)
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro uma pasta para guardar os documentos.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;
                openFileDialog.Title = "Selecionar Documento da Certidão de Não Dívida à Segurança Social";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "NaoDivSegSocial_" + TXT_Codigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
                                          System.IO.Path.GetExtension(sourceFile);
                        string destFile = System.IO.Path.Combine(txtCaminhoPasta.Text, fileName);

                        // Verifica se o arquivo já existe
                        if (System.IO.File.Exists(destFile))
                        {
                            DialogResult result = MessageBox.Show(
                                $"O arquivo {fileName} já existe na pasta de destino. Deseja substituí-lo?",
                                "Arquivo já existe",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (result == DialogResult.No)
                                return;
                        }

                        // Copia o arquivo para a pasta de destino
                        System.IO.File.Copy(sourceFile, destFile, true);

                        // Atualiza o caminho do anexo
                        caminhoAnexoSegSocial = destFile;

                        // Atualiza o label
                        AtualizarLabelsAnexos();

                        MessageBox.Show("Documento anexado com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos
                        AtualizarListaDocumentos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void visualizarAnexoFinancas_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(caminhoAnexoFinancas) || !System.IO.File.Exists(caminhoAnexoFinancas))
            {
                MessageBox.Show("Não existe anexo para a certidão de não dívida às Finanças.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(caminhoAnexoFinancas);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void visualizarAnexoSegSocial_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(caminhoAnexoSegSocial) || !System.IO.File.Exists(caminhoAnexoSegSocial))
            {
                MessageBox.Show("Não existe anexo para a certidão de não dívida à Segurança Social.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(caminhoAnexoSegSocial);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void visualizarFolhaPag_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(caminhoAnexoFolhaPag) || !System.IO.File.Exists(caminhoAnexoFolhaPag))
            {
                MessageBox.Show("Não existe anexo para a Folha Pag.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(caminhoAnexoFolhaPag);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BT_Salvar_Click_Click(object sender, EventArgs e)
        {
            try
            {
                // Atualiza a tabela Geral_Entidade
                var querySalvar = $@"
            UPDATE Geral_Entidade
            SET 
                NIPC = '{TXT_Contribuinte.Text}', 
                AlvaraNumero = '{TXT_Alvara.Text}', 
                AlvaraValidade = '{(TXT_AlvaraValidade.Checked ? TXT_AlvaraValidade.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_NaoDivFinancas = '{(TXT_NaoDivFinancas.Checked ? TXT_NaoDivFinancas.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_NaoDivSegSocial = '{(TXT_NaoDivSegSocial.Checked ? TXT_NaoDivSegSocial.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_FolhaPagSegSocial = '{(TXT_FolhaPagSegSocial.Checked ? TXT_FolhaPagSegSocial.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_ReciboApoliceAT = '{TXT_ReciboApoliceAT.Text}', 
                CDU_ReciboRC = '{TXT_ReciboRC.Text}', 
                CDU_Caminho = '{txtCaminhoPasta.Text}',
                CDU_ReciboPagSegSocial = '{cb_ReciboPagSegSocial.Text}', 
                CDU_ApoliceAT = '{cb_ApoliceAT.Text}', 
                CDU_ApoliceRC = '{cb_ApoliceRC.Text}', 
                CDU_HorarioTrabalho = '{cb_HorarioTrabalho.Text}', 
                CDU_DecTrabIlegais = '{cb_DecTrabIlegais.Text}', 
                CDU_DecRespEstaleiro = '{cb_DecRespEstaleiro.Text}', 
                CDU_DecConhecimPSS = '{cb_DecConhecimPSS.Text}',
                CDU_AnexoFinancas = '{caminhoAnexoFinancas}',
                CDU_AnexoSegSocial = '{caminhoAnexoSegSocial}',
                CDU_FolhaPag = '{caminhoAnexoFolhaPag}'
            WHERE ID = '{_ID}';
        ";

                BSO.DSO.ExecuteSQL(querySalvar);
                MessageBox.Show("Dados salvos com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar os dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Verifica se há linhas no DataGridView
                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Não há dados para salvar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Verifica se uma obra foi selecionada no ComboBox
                if (cb_Obras.SelectedItem == null)
                {
                    MessageBox.Show("Selecione uma obra antes de salvar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Obtém o código da obra selecionada
                string codigoObraSelecionada = ((KeyValuePair<string, string>)cb_Obras.SelectedItem).Key;

                // Percorre cada linha do DataGridView
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Garante que a linha não esteja vazia
                    if (row.Cells[0].Value != null)
                    {
                        string entradaObra = row.Cells[0].Value.ToString();
                        string saidaObra = row.Cells[1].Value.ToString();
                        string contratoSubempreitada = row.Cells[2].Value.ToString();
                        bool autorizacaoEntrada = Convert.ToBoolean(row.Cells[3].Value);
                        Guid id = Guid.NewGuid();

                        // Monta a query de inserção
                        string queryUpsert = $@"
    IF EXISTS (
        SELECT 1 FROM TDU_AD_Obras 
        WHERE CDU_Obra = '{codigoObraSelecionada}' 
        AND CDU_EntradaObra = '{entradaObra}'
        AND CDU_SaidaObra = '{saidaObra}'
        AND CDU_ContratoSubempreitada = '{contratoSubempreitada}'
    )
    BEGIN
        UPDATE TDU_AD_Obras 
        SET CDU_AutorizacaoEntrada = {(autorizacaoEntrada ? 1 : 0)}
        WHERE CDU_Obra = '{codigoObraSelecionada}' 
        AND CDU_EntradaObra = '{entradaObra}'
        AND CDU_SaidaObra = '{saidaObra}'
        AND CDU_ContratoSubempreitada = '{contratoSubempreitada}';
    END
    ELSE
    BEGIN
        INSERT INTO TDU_AD_Obras 
        (CDU_Codigo, CDU_Obra, CDU_EntradaObra, CDU_SaidaObra, CDU_ContratoSubempreitada, CDU_AutorizacaoEntrada) 
        VALUES 
        ('{id}', '{codigoObraSelecionada}', '{entradaObra}', '{saidaObra}', '{contratoSubempreitada}', {(autorizacaoEntrada ? 1 : 0)});
    END";

                        BSO.DSO.ExecuteSQL(queryUpsert);
                    }
                }

                MessageBox.Show("Registros adicionados com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar os dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAnexoFolhaPag_Click(object sender, EventArgs e)
        {

            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro uma pasta para guardar os documentos.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;
                openFileDialog.Title = "Selecionar Documento da Folha Pag.";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "FolhaPagSS_" + TXT_Codigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
                                          System.IO.Path.GetExtension(sourceFile);
                        string destFile = System.IO.Path.Combine(txtCaminhoPasta.Text, fileName);

                        // Verifica se o arquivo já existe
                        if (System.IO.File.Exists(destFile))
                        {
                            DialogResult result = MessageBox.Show(
                                $"O arquivo {fileName} já existe na pasta de destino. Deseja substituí-lo?",
                                "Arquivo já existe",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (result == DialogResult.No)
                                return;
                        }

                        // Copia o arquivo para a pasta de destino
                        System.IO.File.Copy(sourceFile, destFile, true);

                        // Atualiza o caminho do anexo
                        caminhoAnexoFolhaPag = destFile;

                        // Atualiza o label
                        AtualizarLabelsAnexos();

                        MessageBox.Show("Documento anexado com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos
                        AtualizarListaDocumentos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
        }
    }
}
