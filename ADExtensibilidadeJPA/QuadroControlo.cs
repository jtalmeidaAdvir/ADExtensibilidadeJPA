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

        private void CriaCampos()
        {
            var validacampos = $@"IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_TrataSGS')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_TrataSGS BIT;
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_EmailEnviado')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_EmailEnviado BIT;
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_DataEnvio')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_DataEnvio NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_NaoDivFinancas')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_NaoDivFinancas NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_NaoDivSegSocial')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_NaoDivSegSocial NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_FolhaPagSegSocial')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_FolhaPagSegSocial NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_ReciboPagSegSocial')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_ReciboPagSegSocial NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_ReciboApoliceAT')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_ReciboApoliceAT NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_ReciboRC')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_ReciboRC NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_Caminho')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_Caminho NVARCHAR(1000);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_ApoliceAT')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_ApoliceAT NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_ApoliceRC')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_ApoliceRC NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_HorarioTrabalho')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_HorarioTrabalho NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_DecTrabIlegais')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_DecTrabIlegais NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_DecRespEstaleiro')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_DecRespEstaleiro NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_DecConhecimPSS')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_DecConhecimPSS NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_AnexoFinancas')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_AnexoFinancas NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_AnexoSegSocial')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_AnexoSegSocial NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_FolhaPag')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_FolhaPag NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_AnexoApoliceAT')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_AnexoApoliceAT NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_AnexoApoliceRC')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_AnexoApoliceRC NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_AnexoHorarioTrabalho')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_AnexoHorarioTrabalho NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_AnexoD')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_AnexoD NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_DecTrabEmigr')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_DecTrabEmigr NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_InscricaoSS')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_InscricaoSS NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_AnexoDStatus')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_AnexoDStatus NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_DecTrabEmigrStatus')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_DecTrabEmigrStatus NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_InscricaoSSStatus')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_InscricaoSSStatus NVARCHAR(255);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_CaminhoTRab')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_CaminhoTRab NVARCHAR(1000);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_CaminhoEqui')
BEGIN
    ALTER TABLE Geral_Entidade ADD CDU_CaminhoEqui NVARCHAR(1000);
END

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'TDU_AD_Trabalhadores')
BEGIN
    CREATE TABLE TDU_AD_Trabalhadores (
        CDU_ValidadeDocumento DATE NULL,
        CDU_NIF NVARCHAR(50) NULL,
        CDU_NumSS NVARCHAR(50) NULL,
        CDU_FichaAptidao BIT  NULL,
        CDU_CaminhoFichaAptidao NVARCHAR(500) NULL,
        CDU_Credenciacao BIT  NULL,
        CDU_DescCredenciacao NVARCHAR(255) NULL,
        CDU_CaminhoCredenciacao NVARCHAR(500) NULL,
        CDU_FichaEPI BIT  NULL,
        CDU_CaminhoFichaEPI NVARCHAR(500) NULL,
        CDU_Status NVARCHAR(50) NULL,
        CDU_Observacoes NVARCHAR(500) NULL,
        CDU_Caminho NVARCHAR(500) NULL,
        CDU_AnexoCartaoCidadao INT NULL,
        CDU_ValidadeCartaoCidadao DATE NULL,
        nome NVARCHAR(255) NULL,
        categoria NVARCHAR(255) NULL,
        contribuinte NVARCHAR(255) NULL,
        seguranca_social NVARCHAR(255) NULL,
        anexo1 BIT NULL,
        anexo2 BIT NULL,
        anexo3 BIT NULL,
        anexo4 BIT NULL,
        anexo5 BIT NULL,
        id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        id_empresa NVARCHAR(255) NULL,
        caminho1 NVARCHAR(255) NULL,
        caminho2 NVARCHAR(255) NULL,
        caminho3 NVARCHAR(255) NULL,
        caminho4 NVARCHAR(255) NULL,
        caminho5 NVARCHAR(255) NULL
    );
END


IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_Link')
BEGIN
    -- Caso a coluna não exista, cria a coluna CDU_Link com o tipo nvarchar(max)
    ALTER TABLE Geral_Entidade
    ADD CDU_Link nvarchar(max);
END;
";
            BSO.DSO.ExecuteSQL(validacampos);
        }

        private void QuadroControlo_Load(object sender, EventArgs e)
        {
            CriaCampos();
            ConfigurarInterface();
            DadosLista();
            
        }

        private TextBox txtFiltro;
        private Button btnFiltrar;
        private Button btnLimparFiltro;
        private Button btnFiltrarEnviados;
        private Button btnAtualizar;
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
            topPanel.Controls.Add(Bt_Validades);
            BT_Editar.Location = new System.Drawing.Point(10, 9);
            Bt_Email.Location = new System.Drawing.Point(110, 9);
            Bt_Validades.Location = new System.Drawing.Point(600, 9);

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

            // 
            btnAtualizar = new Button
            {
                Text = "Atualizar",
                Location = new System.Drawing.Point(720, 11),
                Size = new System.Drawing.Size(70, 25),
                Font = new System.Drawing.Font("Calibri", 9.5F, System.Drawing.FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                BackColor = System.Drawing.Color.White
            };
            btnAtualizar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(59, 89, 152);
            btnAtualizar.Click += BtnAtualizar_Click;
            panelFiltro.Controls.Add(btnAtualizar);

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
            EstilizarBotao(Bt_Validades, "Validades");

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
                string query = "SELECT id, Nome, CDU_EmailEnviado, CDU_DataEnvio FROM Geral_Entidade WHERE CDU_TrataSGS = 1";
                StdBELista dt = BSO.Consulta(query);

                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("ID", typeof(string));
                dataTable.Columns.Add("Nome", typeof(string));
                dataTable.Columns.Add("EmailEnviadoColumn", typeof(bool));
                dataTable.Columns.Add("DataEnvioColumn", typeof(DateTime));
                dataTable.Columns.Add("Autorizado Em Obra", typeof(bool));
                dataTable.Columns.Add("Documentos Expirados", typeof(bool));


                dt.Inicio();
                while (!dt.NoFim())
                {

                    string id = dt.Valor("id")?.ToString() ?? string.Empty;
                    string nome = dt.Valor("Nome")?.ToString() ?? string.Empty;

                    //bool emailEnviado = bool.TryParse(dt.Valor("CDU_EmailEnviado")?.ToString(), out bool result) ? result : false;
                    var emailEnviadostring = dt.DaValor<string>("CDU_EmailEnviado");
                    bool emailEnviado = emailEnviadostring == "1";

                    //MessageBox.Show(emailEnviadostring);
                    DateTime dataEnvio = DateTime.TryParse(dt.Valor("CDU_DataEnvio")?.ToString(), out DateTime envio) ? envio : DateTime.MinValue;

                    //Verifica
                    var queryauto = $@"SELECT * FROM TDU_AD_Autorizacoes WHERE ID_Entidade = '{id}';";
                    var autorizado = BSO.Consulta(queryauto);
                    bool auto = false;
                    if (autorizado.NumLinhas() > 0)
                    {
                        auto = true;
                    }

                    bool caducado = VerificaDocumentos(id);


                    dataTable.Rows.Add(id, nome, emailEnviado, dataEnvio, auto, caducado);

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

        private bool VerificaDocumentos(string id)
        {
            var queryentidade = $@"SELECT 
    CASE 
        WHEN EXISTS (
            SELECT 1 
            FROM Geral_Entidade 
            WHERE ID = '{id}'
            AND (
                (CDU_ValidadeFinancas < GETDATE() AND CDU_ValidadeFinancas IS NOT NULL) OR
                (CDU_ValidadeSegSocial < GETDATE() AND CDU_ValidadeSegSocial IS NOT NULL) OR
                (CDU_ValidadeFolhaPag < GETDATE() AND CDU_ValidadeFolhaPag IS NOT NULL) OR
                (CDU_ValidadeComprovativoPagamento < GETDATE() AND CDU_ValidadeComprovativoPagamento IS NOT NULL) OR
                (CDU_ValidadeSeguroRC < GETDATE() AND CDU_ValidadeSeguroRC IS NOT NULL) OR
                (CDU_ValidadeSeguroAT < GETDATE() AND CDU_ValidadeSeguroAT IS NOT NULL) OR
                (CDU_ValidadeAlvara < GETDATE() AND CDU_ValidadeAlvara IS NOT NULL) OR
                (CDU_ValidadeCertidaoPermanente < GETDATE() AND CDU_ValidadeCertidaoPermanente IS NOT NULL)     
            )
        ) 
        THEN 'Sim' 
        ELSE 'Não' 
    END AS TemDataVencida;
";

            var entiexist = BSO.Consulta(queryentidade);


            var querytrab = $@"WITH DataExtraida AS (
    SELECT 
        -- Extraindo e convertendo a data no formato DD/MM/YYYY para o formato YYYY-MM-DD
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho1, CHARINDEX('Válido até&#58; ', caminho1) + 16, 10))), 103) AS DATE) AS Data_Caminho1,
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho2, CHARINDEX('Válido até&#58; ', caminho2) + 16, 10))), 103) AS DATE) AS Data_Caminho2,
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho3, CHARINDEX('Válido até&#58; ', caminho3) + 16, 10))), 103) AS DATE) AS Data_Caminho3,
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho4, CHARINDEX('Válido até&#58; ', caminho4) + 16, 10))), 103) AS DATE) AS Data_Caminho4,
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho5, CHARINDEX('Válido até&#58; ', caminho5) + 16, 10))), 103) AS DATE) AS Data_Caminho5
    FROM TDU_AD_Trabalhadores
	WHERE id_empresa = '{id}'
)
SELECT
    Data_Caminho1,
    Data_Caminho2,
    Data_Caminho3,
    Data_Caminho4,
    Data_Caminho5,

    -- Verificação Final para qualquer data expirada, excluindo NULL e 1900-01-01
    CASE
        WHEN 
            (
                -- Verificando se qualquer data é expirada e tratando NULL e 1900-01-01
                (Data_Caminho1 <= CAST(GETDATE() AS DATE) AND Data_Caminho1 <> '1900-01-01' AND Data_Caminho1 IS NOT NULL)
                OR (Data_Caminho2 <= CAST(GETDATE() AS DATE) AND Data_Caminho2 <> '1900-01-01' AND Data_Caminho2 IS NOT NULL)
                OR (Data_Caminho3 <= CAST(GETDATE() AS DATE) AND Data_Caminho3 <> '1900-01-01' AND Data_Caminho3 IS NOT NULL)
                OR (Data_Caminho4 <= CAST(GETDATE() AS DATE) AND Data_Caminho4 <> '1900-01-01' AND Data_Caminho4 IS NOT NULL)
                OR (Data_Caminho5 <= CAST(GETDATE() AS DATE) AND Data_Caminho5 <> '1900-01-01' AND Data_Caminho5 IS NOT NULL)
            )
        THEN 'Sim'
        ELSE 'Não'
    END AS Verificacao_Final
FROM DataExtraida

";

            var trabexit = BSO.Consulta(querytrab);


            var queryEqui = $@"WITH DataExtraida AS (
    SELECT 
        -- Extraindo e convertendo a data no formato DD/MM/YYYY para o formato YYYY-MM-DD
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho5, CHARINDEX('Válido até&#58; ', caminho5) + 16, 10))), 103) AS DATE) AS Data_Caminho5
    FROM TDU_AD_Equipamentos
	WHERE id_empresa = '{id}'
)
SELECT
    Data_Caminho5,

    -- Verificação Final para qualquer data expirada, excluindo NULL e 1900-01-01
    CASE
        WHEN 
            (
                -- Verificando se qualquer data é expirada e tratando NULL e 1900-01-01
                (Data_Caminho5 <= CAST(GETDATE() AS DATE) AND Data_Caminho5 <> '1900-01-01' AND Data_Caminho5 IS NOT NULL)
            )
        THEN 'Sim'
        ELSE 'Não'
    END AS Verificacao_Final
FROM DataExtraida

";

            var equiexit = BSO.Consulta(queryEqui);


            var queryauto = $@"WITH DataExtraida AS (
    SELECT 
        -- Extraindo e convertendo a data no formato DD/MM/YYYY para o formato YYYY-MM-DD
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho1, CHARINDEX('Válido até&#58; ', caminho1) + 16, 10))), 103) AS DATE) AS Data_Caminho1,
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho2, CHARINDEX('Válido até&#58; ', caminho2) + 16, 10))), 103) AS DATE) AS Data_Caminho2,
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho3, CHARINDEX('Válido até&#58; ', caminho3) + 16, 10))), 103) AS DATE) AS Data_Caminho3,
        TRY_CAST(CONVERT(DATE, LTRIM(RTRIM(SUBSTRING(caminho4, CHARINDEX('Válido até&#58; ', caminho4) + 16, 10))), 103) AS DATE) AS Data_Caminho4
    FROM TDU_AD_Autorizacoes
	WHERE ID_Entidade = '{id}'
)
SELECT
    Data_Caminho1,
    Data_Caminho2,
    Data_Caminho3,
    Data_Caminho4,

    -- Verificação Final para qualquer data expirada, excluindo NULL e 1900-01-01
    CASE
        WHEN 
            (
                -- Verificando se qualquer data é expirada e tratando NULL e 1900-01-01
                (Data_Caminho1 <= CAST(GETDATE() AS DATE) AND Data_Caminho1 <> '1900-01-01' AND Data_Caminho1 IS NOT NULL)
                OR (Data_Caminho2 <= CAST(GETDATE() AS DATE) AND Data_Caminho2 <> '1900-01-01' AND Data_Caminho2 IS NOT NULL)
                OR (Data_Caminho3 <= CAST(GETDATE() AS DATE) AND Data_Caminho3 <> '1900-01-01' AND Data_Caminho3 IS NOT NULL)
                OR (Data_Caminho4 <= CAST(GETDATE() AS DATE) AND Data_Caminho4 <> '1900-01-01' AND Data_Caminho4 IS NOT NULL)
            )
        THEN 'Sim'
        ELSE 'Não'
    END AS Verificacao_Final
FROM DataExtraida


";

            var autoexit = BSO.Consulta(queryauto);


            var resultenti = entiexist.DaValor<string>("TemDataVencida");


            var num = trabexit.NumLinhas();
            var num2 = equiexit.NumLinhas();
            var num3 = autoexit.NumLinhas();

            autoexit.Inicio();
            for (int i = 0; i < num3; i++)
            {
                var resultauto = autoexit.DaValor<string>("Verificacao_Final");
                if (resultauto == "Sim")
                {
                    return true;
                }

                autoexit.Seguinte();
            }

            equiexit.Inicio();
            for (int i = 0; i < num2; i++)
            {
                var resultequi = equiexit.DaValor<string>("Verificacao_Final");
                if (resultequi == "Sim")
                {
                    return true;
                }

                equiexit.Seguinte();
            }


            trabexit.Inicio();
            for (int i = 0; i < num; i++)
            {
                var resulttrab = trabexit.DaValor<string>("Verificacao_Final");
                if (resulttrab == "Sim")
                {
                    return true;
                }

                trabexit.Seguinte();
            }


            if (resultenti == "Sim")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        private void BT_Editar_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    string idSelecionado = dataGridView1.SelectedRows[0].Cells["ID"].Value?.ToString();
                    GestaoSubempreitada menuForm = new GestaoSubempreitada(BSO, PSO, idSelecionado);
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
                SELECT ec.Email, ge.CDU_Link
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
                    var link = dt.DaValor<string>("CDU_Link");
                    // Se não houver e-mail, exibir mensagem e retornar
                    if (string.IsNullOrEmpty(email))
                    {
                        var resultado = MessageBox.Show(
                            "Esta entidade não tem um e-mail registado. Deseja criar um agora?",
                            "E-mail não encontrado",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );
                        var enviado = false;
                        if (resultado == DialogResult.Yes)
                        {


                            CriarEmail criarEmail = new CriarEmail(BSO,idSelecionado);
                            if (criarEmail.ShowDialog() == DialogResult.OK)
                            {
                                string emailInserido = criarEmail.Email;
                                email = emailInserido;
                            }



                            // Iniciando o Outlook
                            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                            MailItem emailItem = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);


                            // Definindo o assunto e o corpo do e-mail
                            emailItem.Subject = "Documentação para entrada obra";
                            emailItem.Body = $@"Exmos. Senhores,


No seguimento da indicação de que será subempreiteiro da JPA CONSTRUTORA na empreitada supracitada, solicitamos o envio/anexação da documentação referente à Vossa empresa, aos Vossos colaboradores e aos Equipamentos previstos para a empreitada, conforme a listagem abaixo.

Para colocar a documentação solicitada, por favor, aceda ao seguinte link:
{link}

 

DOCUMENTAÇÃO RELATIVA À EMPRESA:

IMPIC (Alvará de funcionamento)
Apólice de Seguro de Acidentes de Trabalho
Apólice de Seguro de Responsabilidade Civil
Último recibo pago de Seguro de Acidentes de Trabalho (Deve ser atualizado, pois é mensal)
Último recibo pago de Seguro de Responsabilidade Civil
Horário de Trabalho, onde conste o nome e local da empreitada
Folha de Remunerações da Segurança Social atualizada, onde conste o nome dos trabalhadores e o comprovativo de pagamento da TSU
Inscrição na Segurança Social de trabalhadores novos, caso não estejam descritos na última folha de remunerações da Segurança Social
Declaração de não dívida às Finanças
Declaração de não dívida à Segurança Social
Declaração de trabalhadores emigrantes
Declaração de aceitação do PSS ou PTRE (Para subempreitadas)
 

DOCUMENTAÇÃO RELATIVA A TRABALHADORES

Elementos/dados de identificação do trabalhador:
  - N.º B.I./Cartão de cidadão ou título de residência (caso o trabalhador seja estrangeiro) e validade
  - N.º contribuinte
  - N.º segurança social

Registo de posse de Equipamento de Proteção Individual (EPI´s) com validade inferior a 2 anos
Ficha de Aptidão Médica
Contrato de trabalho (com carimbo da ACT) – trabalhadores estrangeiros (nacionalidades referidas pela ACT)
Passaporte e Visto de Permanência ou manifestação de interesse atualizados - trabalhadores estrangeiros
Declaração de aptidão de manobrador (trabalhadores que manobram equipamentos)
DOCUMENTAÇÃO RELATIVA A EQUIPAMENTOS:

Declaração CE de conformidade e manual do equipamento
Seguro do equipamento
Seguro de responsabilidade civil atualizado
Ficha da última revisão
Declaração da empresa a garantir que o equipamento realizou as revisões/manutenções, conforme o plano de revisões/manutenções
Último relatório de Bom Funcionamento do equipamento de acordo com o Decreto-lei 50/2005
 

VIATURAS (Será solicitado se necessário):

Inspeção
Seguro
Documento Único
Nota: A documentação deverá obrigatoriamente ser enviada 48 horas antes da entrada em obra.

 

Equipa de Subempreiteiro

Deve cumprir as obrigações previstas no artigo 22.º do mesmo Decreto-Lei. Deve também garantir que as empresas por si subcontratadas cumpram este mesmo artigo 22.º, bem como o artigo 23.º, no caso da existência de trabalhadores independentes.

É proibido o consumo de bebidas alcoólicas durante o período e no local de trabalho, não sendo permitida a permanência no local de trabalho com uma taxa de álcool igual ou superior a 0,5g/L, nem a presença de estupefacientes.

Recomendações básicas de HST a serem seguidas durante a execução dos trabalhos:

Apenas poderão estar em obra técnicos abrangidos pela apólice do seguro de Acidentes de Trabalho e aptos para a realização dos trabalhos, conforme a Ficha de Aptidão Médica e registo de intervenientes aprovado no PSS.
Recorrer ao uso dos EPCs (Equipamentos de Proteção Coletiva) e EPIs (Equipamentos de Proteção Individual) de acordo com a recomendação deste documento.
Devem ser divulgados a todos os colaboradores em obra os riscos associados à sua atividade/tarefa e respetivas medidas preventivas.
Não é permitida a execução de trabalhos com riscos especiais por parte de trabalhadores isolados.
Todos os colaboradores devem conhecer e respeitar as regras de uso de máquinas e equipamentos, de acordo com o DL 50/2005.
As escadas utilizadas devem ser certificadas e estar em bom estado de conservação (degraus antiderrapantes).
Todos os colaboradores devem conhecer os procedimentos de emergência.
Todos os subempreiteiros devem procurar manter o estaleiro em boa ordem e estado de salubridade.
Todos os subempreiteiros devem eliminar, reciclar ou evacuar resíduos e escombros.

Com os melhores cumprimentos,

";


                            // Definindo o e-mail do destinatário
                            emailItem.To = email;

                            // Abre o Outlook para o usuário revisar o e-mail antes de enviar
                            emailItem.Display();

                            // Atualizando os campos na tabela após o envio do e-mail (pode ser modificado dependendo da necessidade)
            

                            MessageBox.Show("E-mail foi aberto no Outlook. Por favor, revise e envie.");

                            // Recarregar os dados para mostrar as alterações
                            DadosLista();
                        
                            enviado = true;
                        }
                        else
                        {
                            MessageBox.Show("Operação cancelada. Nenhum e-mail foi criado.");
                            enviado = false;
                        }
                        if (enviado) {
                            string updateQuery = $@"
                UPDATE Geral_Entidade 
                SET CDU_EmailEnviado = 1, CDU_DataEnvio = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}'
                WHERE id = '{idSelecionado}'";
                            BSO.DSO.ExecuteSQL(updateQuery);
                        }
                        return;
                    }



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
        private void BtnAtualizar_Click(object sender, EventArgs e)
        {
            DadosLista();
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

        private void Bt_Validades_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    string idSelecionado = dataGridView1.SelectedRows[0].Cells["ID"].Value?.ToString();
                    Validades menuForm = new Validades(BSO, PSO, idSelecionado);
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
    }
}