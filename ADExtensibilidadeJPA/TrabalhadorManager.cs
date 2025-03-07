using System;
using System.Drawing;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public class TrabalhadorManager
    {
        private TabPage _tabPage;
        private DataGridView _dgvTrabalhadores;
        private Panel _panelFormulario;

        public TrabalhadorManager(TabPage tabPage)
        {
            _tabPage = tabPage;

            // Inicializar componentes para trabalhadores
            CriarTabelaTrabalhadores();
            CriarFormularioTrabalhador();
            CriarBotaoAdicionar();
        }

        private void CriarBotaoAdicionar()
        {
            // Verificar se o botão já existe para evitar duplicação
            if (_tabPage.Controls.ContainsKey("btnAdicionarTrabalhador")) return;

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
                Panel panel = _tabPage.Controls["panelFormulario"] as Panel;
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
            _tabPage.Controls.Add(btnAdicionar);
        }

        private void CriarTabelaTrabalhadores()
        {
            // Criar DataGridView com estilo moderno
            _dgvTrabalhadores = new DataGridView
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
            _dgvTrabalhadores.Columns.Add("Empresa", "Empresa");
            _dgvTrabalhadores.Columns.Add("CategoriaFuncao", "Categoria/Função");
            _dgvTrabalhadores.Columns.Add("Contribuinte", "Contribuinte");
            _dgvTrabalhadores.Columns.Add("SegurancaSocial", "Segurança Social");
            _dgvTrabalhadores.Columns.Add("CCidadaoNTitulo", "CC/Passaporte/T.Res./República");
            _dgvTrabalhadores.Columns.Add("FichaMedica", "Ficha de Aptidão Médica");
            _dgvTrabalhadores.Columns.Add("ContratoMesa", "Contrato de Mesa");
            _dgvTrabalhadores.Columns.Add("FormacaoTrabalhador", "Trabalhador Estrangeiro");
            _dgvTrabalhadores.Columns.Add("FormacaoInformacao", "Formação/Informação");
            _dgvTrabalhadores.Columns.Add("EPIs", "EPIs");
            _dgvTrabalhadores.Columns.Add("EntradaObra", "Entrada Obra");
            _dgvTrabalhadores.Columns.Add("SaidaObra", "Saída Obra");
            _dgvTrabalhadores.Columns.Add("AutorizacaoEntrada", "Autorização de Entrada em Obra");

            // Configurar largura das colunas
            _dgvTrabalhadores.Columns["Empresa"].Width = 120;
            _dgvTrabalhadores.Columns["CategoriaFuncao"].Width = 80;
            _dgvTrabalhadores.Columns["Contribuinte"].Width = 90;
            _dgvTrabalhadores.Columns["SegurancaSocial"].Width = 90;
            _dgvTrabalhadores.Columns["CCidadaoNTitulo"].Width = 130;
            _dgvTrabalhadores.Columns["FichaMedica"].Width = 80;
            _dgvTrabalhadores.Columns["ContratoMesa"].Width = 80;
            _dgvTrabalhadores.Columns["FormacaoTrabalhador"].Width = 80;
            _dgvTrabalhadores.Columns["FormacaoInformacao"].Width = 90;
            _dgvTrabalhadores.Columns["EPIs"].Width = 60;
            _dgvTrabalhadores.Columns["EntradaObra"].Width = 80;
            _dgvTrabalhadores.Columns["SaidaObra"].Width = 80;
            _dgvTrabalhadores.Columns["AutorizacaoEntrada"].Width = 120;

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

            _dgvTrabalhadores.Columns.Add(btnEditar);
            _dgvTrabalhadores.Columns.Add(btnRemover);

            // Configurar scrollbar
            _dgvTrabalhadores.ScrollBars = ScrollBars.Both;

            // Adicionar evento para edição e remoção
            _dgvTrabalhadores.CellClick += dgvTrabalhadores_CellClick;

            // Adicionar à tabPage2
            _tabPage.Controls.Add(_dgvTrabalhadores);

            // Adicionar label de título
            Label lblTitulo = new Label
            {
                Text = "Lista de Trabalhadores",
                Font = new Font("Calibri", 12F, FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                AutoSize = true,
                Location = new Point(8, 270)
            };
            _tabPage.Controls.Add(lblTitulo);
        }

        private void CriarFormularioTrabalhador()
        {
            // Painel principal com gradiente
            _panelFormulario = new Panel
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
            _panelFormulario.Controls.Add(titleBar);

            // Área de conteúdo
            Panel contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15)
            };
            _panelFormulario.Controls.Add(contentPanel);

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
                Location = new Point(_panelFormulario.Width - 200, 13),
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
                Location = new Point(_panelFormulario.Width - 100, 13),
                Size = new Size(90, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                Font = new Font("Calibri", 9.5F, FontStyle.Bold)
            };
            btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 89, 152);
            btnCancelar.Click += (s, e) => _panelFormulario.Visible = false;

            buttonPanel.Controls.Add(btnSalvar);
            buttonPanel.Controls.Add(btnCancelar);
            _panelFormulario.Controls.Add(buttonPanel);

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
            _tabPage.Controls.Add(_panelFormulario);
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            Panel contentPanel = _panelFormulario.Controls[1] as Panel;

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

            // Criar uma nova linha com todos os campos
            DataGridViewRow row = new DataGridViewRow();
            _dgvTrabalhadores.Rows.Add(
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
            int lastRowIndex = _dgvTrabalhadores.Rows.Count - 1;
            if (lastRowIndex % 2 == 0)
            {
                _dgvTrabalhadores.Rows[lastRowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.White;
            }
            else
            {
                _dgvTrabalhadores.Rows[lastRowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 242, 245);
            }

            // Ocultar o formulário
            _panelFormulario.Visible = false;

            // Mostrar mensagem de sucesso
            MessageBox.Show("Trabalhador adicionado com sucesso!", "Sucesso",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // New handler for cell click event from the Menu class
        public void HandleCellClick(DataGridView dgv, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (dgv.Columns[e.ColumnIndex].Name == "Editar")
                {
                    Panel contentPanel = _panelFormulario.Controls[1] as Panel;

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
                    _panelFormulario.Visible = true;
                    _panelFormulario.BringToFront();
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

        // Keep the original method for compatibility, but delegate to the new handler
        private void dgvTrabalhadores_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            HandleCellClick(dgv, e);
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
    }
}