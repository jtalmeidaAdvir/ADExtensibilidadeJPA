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
                Size = new Size(780, 250),
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

            // Adicionar colunas conforme imagem de referência atualizada
            _dgvTrabalhadores.Columns.Add("Empresa", "Empresa");
            _dgvTrabalhadores.Columns.Add("NomeCompleto", "Nome Completo");
            _dgvTrabalhadores.Columns.Add("CategoriaFuncao", "Categoria/Função");
            _dgvTrabalhadores.Columns.Add("Contribuinte", "Contribuinte");
            _dgvTrabalhadores.Columns.Add("SegurancaSocial", "Segurança Social");
            _dgvTrabalhadores.Columns.Add("CCidadaoNTitulo", "CC/Passaporte/Título");
            _dgvTrabalhadores.Columns.Add("CCValidade", "CC Validade");
            _dgvTrabalhadores.Columns.Add("FichaAptidao", "Ficha de Aptidão Médica");
            _dgvTrabalhadores.Columns.Add("FichaEPI", "Ficha de EPI");
            _dgvTrabalhadores.Columns.Add("FichaAssociado", "Ficha de Associado");
            _dgvTrabalhadores.Columns.Add("ContaMesa", "Conta Mesa");
            _dgvTrabalhadores.Columns.Add("MapaSS", "Consta no Mapa SS / Inscrito?");
            _dgvTrabalhadores.Columns.Add("TrabalhadorEstrangeiro", "Trabalhador Estrangeiro");
            _dgvTrabalhadores.Columns.Add("VistoNumero", "Trab. Estrangeiro - Visto número");
            _dgvTrabalhadores.Columns.Add("VistoValidade", "Trab. Estrangeiro - Visto Validade");
            _dgvTrabalhadores.Columns.Add("ContratoACTData", "Trab. Estrangeiro - Contrato Carimb. ACT - Data");
            _dgvTrabalhadores.Columns.Add("FormacaoSobrepresion", "Formação/Sobrepressão");
            _dgvTrabalhadores.Columns.Add("FormacaoAcolhimentoData", "Formação/Informação Acolhimento Data");
            _dgvTrabalhadores.Columns.Add("FormacaoEspecificaData", "Formação/Informação Específica Data");
            _dgvTrabalhadores.Columns.Add("Contacto", "Contacto");
            _dgvTrabalhadores.Columns.Add("ACAUDEmSr", "ACAUDEM Sr");
            _dgvTrabalhadores.Columns.Add("TrAutoriz", "Tr. Autoriz");
            _dgvTrabalhadores.Columns.Add("Cadastro1AvisoData", "Cadastro 1.º Aviso Data");
            _dgvTrabalhadores.Columns.Add("Cadastro2AvisoData", "Cadastro 2.º Aviso Data");
            _dgvTrabalhadores.Columns.Add("EntradaObra", "Entrada Obra");
            _dgvTrabalhadores.Columns.Add("SaidaObra", "Saída Obra");
            _dgvTrabalhadores.Columns.Add("AutorizacaoEntrada", "Autorização de Entrada em Obra");

            // Configurar largura das colunas
            _dgvTrabalhadores.Columns["Empresa"].Width = 100;
            _dgvTrabalhadores.Columns["NomeCompleto"].Width = 150;
            _dgvTrabalhadores.Columns["CategoriaFuncao"].Width = 80;
            _dgvTrabalhadores.Columns["Contribuinte"].Width = 80;
            _dgvTrabalhadores.Columns["SegurancaSocial"].Width = 80;
            _dgvTrabalhadores.Columns["CCidadaoNTitulo"].Width = 80;
            _dgvTrabalhadores.Columns["CCValidade"].Width = 80;
            _dgvTrabalhadores.Columns["FichaAptidao"].Width = 80;
            _dgvTrabalhadores.Columns["FichaEPI"].Width = 80;
            _dgvTrabalhadores.Columns["FichaAssociado"].Width = 80;
            _dgvTrabalhadores.Columns["ContaMesa"].Width = 80;
            _dgvTrabalhadores.Columns["MapaSS"].Width = 100;
            _dgvTrabalhadores.Columns["TrabalhadorEstrangeiro"].Width = 80;
            _dgvTrabalhadores.Columns["VistoNumero"].Width = 80;
            _dgvTrabalhadores.Columns["VistoValidade"].Width = 80;
            _dgvTrabalhadores.Columns["ContratoACTData"].Width = 100;
            _dgvTrabalhadores.Columns["FormacaoSobrepresion"].Width = 80;
            _dgvTrabalhadores.Columns["FormacaoAcolhimentoData"].Width = 100;
            _dgvTrabalhadores.Columns["FormacaoEspecificaData"].Width = 100;
            _dgvTrabalhadores.Columns["Contacto"].Width = 80;
            _dgvTrabalhadores.Columns["ACAUDEmSr"].Width = 80;
            _dgvTrabalhadores.Columns["TrAutoriz"].Width = 80;
            _dgvTrabalhadores.Columns["Cadastro1AvisoData"].Width = 80;
            _dgvTrabalhadores.Columns["Cadastro2AvisoData"].Width = 80;
            _dgvTrabalhadores.Columns["EntradaObra"].Width = 80;
            _dgvTrabalhadores.Columns["SaidaObra"].Width = 80;
            _dgvTrabalhadores.Columns["AutorizacaoEntrada"].Width = 110;

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
                Size = new Size(_tabPage.Width - 20, 440),
                Location = new Point(8, 290),
                Visible = false,
                BackColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = false, // Desativamos aqui porque o contentPanel terá o scroll
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
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

            // Área de conteúdo com suporte a scroll horizontal e vertical
            Panel contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15),
                AutoScroll = true,
                Width = 1200, // Definir uma largura maior para garantir que todos os campos sejam visíveis com scroll
                MinimumSize = new Size(1200, 0),
                AutoScrollMargin = new Size(20, 20),
                AutoScrollMinSize = new Size(1200, 600) // Garante dimensões mínimas para ativar o scroll
            };
            _panelFormulario.Controls.Add(contentPanel);

            // Primeira coluna - Dados básicos
            int yPos = 50;
            int labelWidth = 130;
            int controlWidth = 150;
            int controlHeight = 25;
            int spacing = 28;

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

            // Nome Completo
            Label lblNomeCompleto = new Label
            {
                Text = "Nome Completo:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            TextBox txtNomeCompleto = new TextBox
            {
                Name = "txtNomeCompleto",
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

            // CC Validade
            Label lblCCValidade = new Label
            {
                Text = "CC Validade:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpCCValidade = new DateTimePicker
            {
                Name = "dtpCCValidade",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Ficha de Aptidão
            Label lblFichaAptidao = new Label
            {
                Text = "Ficha Aptidão:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbFichaAptidao = new ComboBox
            {
                Name = "cmbFichaAptidao",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbFichaAptidao.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbFichaAptidao.SelectedIndex = 0;
            yPos += spacing;

            // Ficha de EPI
            Label lblFichaEPI = new Label
            {
                Text = "Ficha de EPI:",
                Location = new Point(10, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbFichaEPI = new ComboBox
            {
                Name = "cmbFichaEPI",
                Location = new Point(labelWidth + 10, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbFichaEPI.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbFichaEPI.SelectedIndex = 0;
            yPos += spacing;

            // Segunda coluna
            yPos = 50;
            int col2X = 320;

            // Ficha de Associado
            Label lblFichaAssociado = new Label
            {
                Text = "Ficha Associado:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbFichaAssociado = new ComboBox
            {
                Name = "cmbFichaAssociado",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbFichaAssociado.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbFichaAssociado.SelectedIndex = 0;
            yPos += spacing;

            // Conta Mesa
            Label lblContaMesa = new Label
            {
                Text = "Conta Mesa:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbContaMesa = new ComboBox
            {
                Name = "cmbContaMesa",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbContaMesa.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbContaMesa.SelectedIndex = 0;
            yPos += spacing;

            // Mapa SS
            Label lblMapaSS = new Label
            {
                Text = "Mapa SS/Inscrito:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbMapaSS = new ComboBox
            {
                Name = "cmbMapaSS",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbMapaSS.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbMapaSS.SelectedIndex = 0;
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

            // Visto Número
            Label lblVistoNumero = new Label
            {
                Text = "Visto Número:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            TextBox txtVistoNumero = new TextBox
            {
                Name = "txtVistoNumero",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight - 5,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Visto Validade
            Label lblVistoValidade = new Label
            {
                Text = "Visto Validade:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpVistoValidade = new DateTimePicker
            {
                Name = "dtpVistoValidade",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Contrato ACT Data
            Label lblContratoACTData = new Label
            {
                Text = "Contrato ACT Data:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpContratoACTData = new DateTimePicker
            {
                Name = "dtpContratoACTData",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Formação/Sobrepressão
            Label lblFormacaoSobrepresion = new Label
            {
                Text = "Formação/Sobrep.:",
                Location = new Point(col2X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbFormacaoSobrepresion = new ComboBox
            {
                Name = "cmbFormacaoSobrepresion",
                Location = new Point(col2X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbFormacaoSobrepresion.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbFormacaoSobrepresion.SelectedIndex = 0;
            yPos += spacing;

            // Terceira coluna
            yPos = 50;
            int col3X = 600;

            // Formação Acolhimento Data
            Label lblFormacaoAcolhimentoData = new Label
            {
                Text = "Form. Acolh. Data:",
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpFormacaoAcolhimentoData = new DateTimePicker
            {
                Name = "dtpFormacaoAcolhimentoData",
                Location = new Point(col3X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Formação Específica Data
            Label lblFormacaoEspecificaData = new Label
            {
                Text = "Form. Espec. Data:",
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpFormacaoEspecificaData = new DateTimePicker
            {
                Name = "dtpFormacaoEspecificaData",
                Location = new Point(col3X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Contacto
            Label lblContacto = new Label
            {
                Text = "Contacto:",
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            TextBox txtContacto = new TextBox
            {
                Name = "txtContacto",
                Location = new Point(col3X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight - 5,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // ACAUDEM Sr
            Label lblACAUDEmSr = new Label
            {
                Text = "ACAUDEM Sr:",
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbACAUDEmSr = new ComboBox
            {
                Name = "cmbACAUDEmSr",
                Location = new Point(col3X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbACAUDEmSr.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbACAUDEmSr.SelectedIndex = 0;
            yPos += spacing;

            // Tr. Autoriz
            Label lblTrAutoriz = new Label
            {
                Text = "Tr. Autoriz:",
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            ComboBox cmbTrAutoriz = new ComboBox
            {
                Name = "cmbTrAutoriz",
                Location = new Point(col3X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Calibri", 9F)
            };
            cmbTrAutoriz.Items.AddRange(new object[] { "C", "N/C", "N/A" });
            cmbTrAutoriz.SelectedIndex = 0;
            yPos += spacing;

            // Cadastro 1º Aviso Data
            Label lblCadastro1AvisoData = new Label
            {
                Text = "Cadastro 1º Aviso:",
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpCadastro1AvisoData = new DateTimePicker
            {
                Name = "dtpCadastro1AvisoData",
                Location = new Point(col3X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Cadastro 2º Aviso Data
            Label lblCadastro2AvisoData = new Label
            {
                Text = "Cadastro 2º Aviso:",
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpCadastro2AvisoData = new DateTimePicker
            {
                Name = "dtpCadastro2AvisoData",
                Location = new Point(col3X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Entrada Obra
            Label lblEntrada = new Label
            {
                Text = "Entrada Obra:",
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpEntrada = new DateTimePicker
            {
                Name = "dtpEntrada",
                Location = new Point(col3X + labelWidth, yPos),
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
                Location = new Point(col3X, yPos),
                Size = new Size(labelWidth, controlHeight),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Calibri", 9F)
            };
            DateTimePicker dtpSaida = new DateTimePicker
            {
                Name = "dtpSaida",
                Location = new Point(col3X + labelWidth, yPos),
                Width = controlWidth,
                Height = controlHeight,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Calibri", 9F)
            };
            yPos += spacing;

            // Autorização de Entrada
            CheckBox chkAutorizado = new CheckBox
            {
                Text = "Autorização de Entrada em Obra",
                Name = "chkAutorizado",
                Location = new Point(10, yPos),
                Size = new Size(controlWidth + labelWidth, controlHeight),
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
                lblNomeCompleto, txtNomeCompleto,
                lblFuncao, txtFuncao,
                lblContribuinte, txtContribuinte,
                lblSegurancaSocial, txtSegurancaSocial,
                lblCCidadao, txtCCidadao,
                lblCCValidade, dtpCCValidade,
                lblFichaAptidao, cmbFichaAptidao,
                lblFichaEPI, cmbFichaEPI,
                lblFichaAssociado, cmbFichaAssociado,
                lblContaMesa, cmbContaMesa,
                lblMapaSS, cmbMapaSS,
                lblTrabEstrangeiro, cmbTrabEstrangeiro,
                lblVistoNumero, txtVistoNumero,
                lblVistoValidade, dtpVistoValidade,
                lblContratoACTData, dtpContratoACTData,
                lblFormacaoSobrepresion, cmbFormacaoSobrepresion,
                lblFormacaoAcolhimentoData, dtpFormacaoAcolhimentoData,
                lblFormacaoEspecificaData, dtpFormacaoEspecificaData,
                lblContacto, txtContacto,
                lblACAUDEmSr, cmbACAUDEmSr,
                lblTrAutoriz, cmbTrAutoriz,
                lblCadastro1AvisoData, dtpCadastro1AvisoData,
                lblCadastro2AvisoData, dtpCadastro2AvisoData,
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
            TextBox txtNomeCompleto = contentPanel.Controls["txtNomeCompleto"] as TextBox;
            TextBox txtFuncao = contentPanel.Controls["txtFuncao"] as TextBox;
            TextBox txtContribuinte = contentPanel.Controls["txtContribuinte"] as TextBox;
            TextBox txtSegurancaSocial = contentPanel.Controls["txtSegurancaSocial"] as TextBox;
            TextBox txtCCidadao = contentPanel.Controls["txtCCidadao"] as TextBox;
            DateTimePicker dtpCCValidade = contentPanel.Controls["dtpCCValidade"] as DateTimePicker;
            ComboBox cmbFichaAptidao = contentPanel.Controls["cmbFichaAptidao"] as ComboBox;
            ComboBox cmbFichaEPI = contentPanel.Controls["cmbFichaEPI"] as ComboBox;
            ComboBox cmbFichaAssociado = contentPanel.Controls["cmbFichaAssociado"] as ComboBox;
            ComboBox cmbContaMesa = contentPanel.Controls["cmbContaMesa"] as ComboBox;
            ComboBox cmbMapaSS = contentPanel.Controls["cmbMapaSS"] as ComboBox;
            ComboBox cmbTrabEstrangeiro = contentPanel.Controls["cmbTrabEstrangeiro"] as ComboBox;
            TextBox txtVistoNumero = contentPanel.Controls["txtVistoNumero"] as TextBox;
            DateTimePicker dtpVistoValidade = contentPanel.Controls["dtpVistoValidade"] as DateTimePicker;
            DateTimePicker dtpContratoACTData = contentPanel.Controls["dtpContratoACTData"] as DateTimePicker;
            ComboBox cmbFormacaoSobrepresion = contentPanel.Controls["cmbFormacaoSobrepresion"] as ComboBox;
            DateTimePicker dtpFormacaoAcolhimentoData = contentPanel.Controls["dtpFormacaoAcolhimentoData"] as DateTimePicker;
            DateTimePicker dtpFormacaoEspecificaData = contentPanel.Controls["dtpFormacaoEspecificaData"] as DateTimePicker;
            TextBox txtContacto = contentPanel.Controls["txtContacto"] as TextBox;
            ComboBox cmbACAUDEmSr = contentPanel.Controls["cmbACAUDEmSr"] as ComboBox;
            ComboBox cmbTrAutoriz = contentPanel.Controls["cmbTrAutoriz"] as ComboBox;
            DateTimePicker dtpCadastro1AvisoData = contentPanel.Controls["dtpCadastro1AvisoData"] as DateTimePicker;
            DateTimePicker dtpCadastro2AvisoData = contentPanel.Controls["dtpCadastro2AvisoData"] as DateTimePicker;
            DateTimePicker dtpEntrada = contentPanel.Controls["dtpEntrada"] as DateTimePicker;
            DateTimePicker dtpSaida = contentPanel.Controls["dtpSaida"] as DateTimePicker;
            CheckBox chkAutorizado = contentPanel.Controls["chkAutorizado"] as CheckBox;

            // Criar uma nova linha com todos os campos
            DataGridViewRow row = new DataGridViewRow();
            _dgvTrabalhadores.Rows.Add(
                txtEmpresa.Text,
                txtNomeCompleto.Text,
                txtFuncao.Text,
                txtContribuinte.Text,
                txtSegurancaSocial.Text,
                txtCCidadao.Text,
                dtpCCValidade.Value.ToShortDateString(),
                cmbFichaAptidao.Text,
                cmbFichaEPI.Text,
                cmbFichaAssociado.Text,
                cmbContaMesa.Text,
                cmbMapaSS.Text,
                cmbTrabEstrangeiro.Text,
                txtVistoNumero.Text,
                dtpVistoValidade.Value.ToShortDateString(),
                dtpContratoACTData.Value.ToShortDateString(),
                cmbFormacaoSobrepresion.Text,
                dtpFormacaoAcolhimentoData.Value.ToShortDateString(),
                dtpFormacaoEspecificaData.Value.ToShortDateString(),
                txtContacto.Text,
                cmbACAUDEmSr.Text,
                cmbTrAutoriz.Text,
                dtpCadastro1AvisoData.Value.ToShortDateString(),
                dtpCadastro2AvisoData.Value.ToShortDateString(),
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
                    TextBox txtNomeCompleto = contentPanel.Controls["txtNomeCompleto"] as TextBox;
                    TextBox txtFuncao = contentPanel.Controls["txtFuncao"] as TextBox;
                    TextBox txtContribuinte = contentPanel.Controls["txtContribuinte"] as TextBox;
                    TextBox txtSegurancaSocial = contentPanel.Controls["txtSegurancaSocial"] as TextBox;
                    TextBox txtCCidadao = contentPanel.Controls["txtCCidadao"] as TextBox;
                    DateTimePicker dtpCCValidade = contentPanel.Controls["dtpCCValidade"] as DateTimePicker;
                    ComboBox cmbFichaAptidao = contentPanel.Controls["cmbFichaAptidao"] as ComboBox;
                    ComboBox cmbFichaEPI = contentPanel.Controls["cmbFichaEPI"] as ComboBox;
                    ComboBox cmbFichaAssociado = contentPanel.Controls["cmbFichaAssociado"] as ComboBox;
                    ComboBox cmbContaMesa = contentPanel.Controls["cmbContaMesa"] as ComboBox;
                    ComboBox cmbMapaSS = contentPanel.Controls["cmbMapaSS"] as ComboBox;
                    ComboBox cmbTrabEstrangeiro = contentPanel.Controls["cmbTrabEstrangeiro"] as ComboBox;
                    TextBox txtVistoNumero = contentPanel.Controls["txtVistoNumero"] as TextBox;
                    DateTimePicker dtpVistoValidade = contentPanel.Controls["dtpVistoValidade"] as DateTimePicker;
                    DateTimePicker dtpContratoACTData = contentPanel.Controls["dtpContratoACTData"] as DateTimePicker;
                    ComboBox cmbFormacaoSobrepresion = contentPanel.Controls["cmbFormacaoSobrepresion"] as ComboBox;
                    DateTimePicker dtpFormacaoAcolhimentoData = contentPanel.Controls["dtpFormacaoAcolhimentoData"] as DateTimePicker;
                    DateTimePicker dtpFormacaoEspecificaData = contentPanel.Controls["dtpFormacaoEspecificaData"] as DateTimePicker;
                    TextBox txtContacto = contentPanel.Controls["txtContacto"] as TextBox;
                    ComboBox cmbACAUDEmSr = contentPanel.Controls["cmbACAUDEmSr"] as ComboBox;
                    ComboBox cmbTrAutoriz = contentPanel.Controls["cmbTrAutoriz"] as ComboBox;
                    DateTimePicker dtpCadastro1AvisoData = contentPanel.Controls["dtpCadastro1AvisoData"] as DateTimePicker;
                    DateTimePicker dtpCadastro2AvisoData = contentPanel.Controls["dtpCadastro2AvisoData"] as DateTimePicker;
                    DateTimePicker dtpEntrada = contentPanel.Controls["dtpEntrada"] as DateTimePicker;
                    DateTimePicker dtpSaida = contentPanel.Controls["dtpSaida"] as DateTimePicker;
                    CheckBox chkAutorizado = contentPanel.Controls["chkAutorizado"] as CheckBox;

                    // Preencher formulário com os dados da linha
                    txtEmpresa.Text = dgv.Rows[e.RowIndex].Cells["Empresa"].Value?.ToString() ?? "";
                    txtNomeCompleto.Text = dgv.Rows[e.RowIndex].Cells["NomeCompleto"].Value?.ToString() ?? "";
                    txtFuncao.Text = dgv.Rows[e.RowIndex].Cells["CategoriaFuncao"].Value?.ToString() ?? "";
                    txtContribuinte.Text = dgv.Rows[e.RowIndex].Cells["Contribuinte"].Value?.ToString() ?? "";
                    txtSegurancaSocial.Text = dgv.Rows[e.RowIndex].Cells["SegurancaSocial"].Value?.ToString() ?? "";
                    txtCCidadao.Text = dgv.Rows[e.RowIndex].Cells["CCidadaoNTitulo"].Value?.ToString() ?? "";

                    // Configurar DateTimePicker para CC Validade
                    string ccValidadeStr = dgv.Rows[e.RowIndex].Cells["CCValidade"].Value?.ToString();
                    if (DateTime.TryParse(ccValidadeStr, out DateTime ccValidadeDate))
                    {
                        dtpCCValidade.Value = ccValidadeDate;
                    }

                    // Selecionar os itens nas ComboBoxes
                    SelectComboBoxItem(cmbFichaAptidao, dgv.Rows[e.RowIndex].Cells["FichaAptidao"].Value?.ToString());
                    SelectComboBoxItem(cmbFichaEPI, dgv.Rows[e.RowIndex].Cells["FichaEPI"].Value?.ToString());
                    SelectComboBoxItem(cmbFichaAssociado, dgv.Rows[e.RowIndex].Cells["FichaAssociado"].Value?.ToString());
                    SelectComboBoxItem(cmbContaMesa, dgv.Rows[e.RowIndex].Cells["ContaMesa"].Value?.ToString());
                    SelectComboBoxItem(cmbMapaSS, dgv.Rows[e.RowIndex].Cells["MapaSS"].Value?.ToString());
                    SelectComboBoxItem(cmbTrabEstrangeiro, dgv.Rows[e.RowIndex].Cells["TrabalhadorEstrangeiro"].Value?.ToString());

                    // Preencher campos relacionados a trabalhador estrangeiro
                    txtVistoNumero.Text = dgv.Rows[e.RowIndex].Cells["VistoNumero"].Value?.ToString() ?? "";

                    string vistoValidadeStr = dgv.Rows[e.RowIndex].Cells["VistoValidade"].Value?.ToString();
                    if (DateTime.TryParse(vistoValidadeStr, out DateTime vistoValidadeDate))
                    {
                        dtpVistoValidade.Value = vistoValidadeDate;
                    }

                    string contratoACTStr = dgv.Rows[e.RowIndex].Cells["ContratoACTData"].Value?.ToString();
                    if (DateTime.TryParse(contratoACTStr, out DateTime contratoACTDate))
                    {
                        dtpContratoACTData.Value = contratoACTDate;
                    }

                    // Preencher campos de formação
                    SelectComboBoxItem(cmbFormacaoSobrepresion, dgv.Rows[e.RowIndex].Cells["FormacaoSobrepresion"].Value?.ToString());

                    string formacaoAcolhimentoStr = dgv.Rows[e.RowIndex].Cells["FormacaoAcolhimentoData"].Value?.ToString();
                    if (DateTime.TryParse(formacaoAcolhimentoStr, out DateTime formacaoAcolhimentoDate))
                    {
                        dtpFormacaoAcolhimentoData.Value = formacaoAcolhimentoDate;
                    }

                    string formacaoEspecificaStr = dgv.Rows[e.RowIndex].Cells["FormacaoEspecificaData"].Value?.ToString();
                    if (DateTime.TryParse(formacaoEspecificaStr, out DateTime formacaoEspecificaDate))
                    {
                        dtpFormacaoEspecificaData.Value = formacaoEspecificaDate;
                    }

                    // Outros campos
                    txtContacto.Text = dgv.Rows[e.RowIndex].Cells["Contacto"].Value?.ToString() ?? "";
                    SelectComboBoxItem(cmbACAUDEmSr, dgv.Rows[e.RowIndex].Cells["ACAUDEmSr"].Value?.ToString());
                    SelectComboBoxItem(cmbTrAutoriz, dgv.Rows[e.RowIndex].Cells["TrAutoriz"].Value?.ToString());

                    // Cadastro avisos
                    string cadastro1AvisoStr = dgv.Rows[e.RowIndex].Cells["Cadastro1AvisoData"].Value?.ToString();
                    if (DateTime.TryParse(cadastro1AvisoStr, out DateTime cadastro1AvisoDate))
                    {
                        dtpCadastro1AvisoData.Value = cadastro1AvisoDate;
                    }

                    string cadastro2AvisoStr = dgv.Rows[e.RowIndex].Cells["Cadastro2AvisoData"].Value?.ToString();
                    if (DateTime.TryParse(cadastro2AvisoStr, out DateTime cadastro2AvisoDate))
                    {
                        dtpCadastro2AvisoData.Value = cadastro2AvisoDate;
                    }

                    // Entrada e saída obra
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