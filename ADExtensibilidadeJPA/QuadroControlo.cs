using Microsoft.Office.Interop.Outlook;  // Para o Outlook
using Microsoft.Office.Interop;
using Primavera.Extensibility.CustomForm;
using StdBE100;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Net.Http;
using System.Threading.Tasks;
using static System.Windows.Forms.LinkLabel;
using DocumentFormat.OpenXml.Office2010.Excel;
using PrimaveraSDK;
using PRISDK100;
using PriTextBoxF4100;
using System.Runtime.InteropServices;

namespace ADExtensibilidadeJPA
{
    public partial class QuadroControlo : CustomForm
    {
        private bool controlsInitialized;
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
            AjustarFillComBaseNosHeadersECelulas();
            AdicionarCheckBoxCabecalho();

        }

        private System.Windows.Forms.TextBox txtFiltro;
        private System.Windows.Forms.Button btnFiltrar;
        private System.Windows.Forms.Button btnLimparFiltro;
        private Button btnFiltrarEnviados;
        private Button btnAtualizar;
        private DataTable dataOriginal;

        CheckBox cbHeader;

        void AdicionarCheckBoxCabecalho()
        {
            cbHeader = new CheckBox();
            cbHeader.Size = new Size(18, 18);
            cbHeader.BackColor = System.Drawing.Color.Transparent;

            cbHeader.CheckedChanged += (s, e) =>
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                        row.Cells[" "].Value = cbHeader.Checked;
                }
            };

            dataGridView1.Controls.Add(cbHeader);
            ReposicionarCheckBoxCabecalho();

            dataGridView1.ColumnWidthChanged += (s, e) => ReposicionarCheckBoxCabecalho();
            dataGridView1.Scroll += (s, e) => ReposicionarCheckBoxCabecalho();
            dataGridView1.SizeChanged += (s, e) => ReposicionarCheckBoxCabecalho();
            this.Resize += (s, e) => ReposicionarCheckBoxCabecalho(); // opcional
        }

        void ReposicionarCheckBoxCabecalho()
        {
            if (!dataGridView1.Columns.Contains(" "))
                return;

            int indexColuna = dataGridView1.Columns[" "].Index;
            Rectangle cabecalho = dataGridView1.GetCellDisplayRectangle(indexColuna, -1, true);
            cbHeader.Location = new Point(
                cabecalho.Location.X + (cabecalho.Width - cbHeader.Width) / 2,
                cabecalho.Location.Y + (cabecalho.Height - cbHeader.Height) / 2
            );
        }

        private void AjustarFillComBaseNosHeadersECelulas()
        {
            float totalLargura = 0;
            Dictionary<DataGridViewColumn, float> larguras = new Dictionary<DataGridViewColumn, float>();

            using (Graphics g = dataGridView1.CreateGraphics())
            {
                foreach (DataGridViewColumn coluna in dataGridView1.Columns)
                {
                    float larguraMaxima = 0;

                    // Medir texto do cabeçalho
                    SizeF tamanhoHeader = g.MeasureString(coluna.HeaderText, dataGridView1.ColumnHeadersDefaultCellStyle.Font);
                    larguraMaxima = tamanhoHeader.Width;

                    // Medir texto das células visíveis
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            object valor = row.Cells[coluna.Index].Value;
                            if (valor != null)
                            {
                                SizeF tamanhoCelula = g.MeasureString(valor.ToString(), dataGridView1.DefaultCellStyle.Font);
                                if (tamanhoCelula.Width > larguraMaxima)
                                    larguraMaxima = tamanhoCelula.Width;
                            }
                        }
                    }

                    // Adiciona margem
                    larguraMaxima += 20; // padding horizontal extra

                    larguras[coluna] = larguraMaxima;
                    totalLargura += larguraMaxima;
                }
            }

            // Aplicar FillWeight proporcional
            foreach (var item in larguras)
            {
                float proporcao = (item.Value / totalLargura) * 100f;
                item.Key.FillWeight = proporcao;
            }

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }

        private void ConfigurarInterface()
        {

            ToolTip toolTip = new ToolTip();

            // Configuração do formulário principal com gradiente visual moderno
            this.BackColor = System.Drawing.Color.FromArgb(240, 242, 245);

            // Criar um painel de topo com gradiente (para os botões)
            System.Windows.Forms.Panel topPanel = new System.Windows.Forms.Panel
            {
                Height = 45,
                Dock = DockStyle.Top,
                BackColor = System.Drawing.Color.White
            };
            this.Controls.Add(topPanel);



            // Adicionar controles de filtro
            System.Windows.Forms.Panel panelFiltro = new System.Windows.Forms.Panel
            {
                Height = 45,
                Dock = DockStyle.Top, // Vai logo abaixo do painel de topo
                BackColor = System.Drawing.Color.FromArgb(59, 89, 152)
            };
            this.Controls.Add(panelFiltro);


            // Mover os botões para o painel de topo e reestilizar
            //panelFiltro.Controls.Add(BT_Editar);


            panelFiltro.Controls.Add(Bt_Email);
            panelFiltro.Controls.Add(Bt_Validades);
            panelFiltro.Controls.Add(Bt_Avisos);
            panelFiltro.Controls.Add(Bt_imprimir);
            panelFiltro.Controls.Add(Bt_imprimir2);
            panelFiltro.Controls.Add(BT_ImprimirJPA);
            panelFiltro.Controls.Add(BT_CriarTrabalhadores);
            BT_Editar.Location = new System.Drawing.Point(10, 9);
            Bt_Email.Location = new System.Drawing.Point(10, 9);
            Bt_Validades.Location = new System.Drawing.Point(330, 9);
            Bt_Avisos.Location = new System.Drawing.Point(170, 9);
            Bt_imprimir.Location = new System.Drawing.Point(490, 9);
            Bt_imprimir2.Location = new System.Drawing.Point(980, 9);
            BT_ImprimirJPA.Location = new System.Drawing.Point(650, 9);
            BT_CriarTrabalhadores.Location = new System.Drawing.Point(810, 9);




            // Label para o filtro
            Label lblFiltro = new Label
            {
                Text = "Filtrar por Nome:",
                Font = new System.Drawing.Font("Calibri", 10F),
                ForeColor = System.Drawing.Color.FromArgb(59, 89, 152),
                AutoSize = true,
                Location = new System.Drawing.Point(10, 14)
            };
            topPanel.Controls.Add(lblFiltro);

            // Textbox para o filtro
            txtFiltro = new TextBox
            {
                Location = new System.Drawing.Point(120, 12),
                Size = new System.Drawing.Size(300, 23),
                Font = new System.Drawing.Font("Calibri", 10F),
                BorderStyle = BorderStyle.FixedSingle
            };
            topPanel.Controls.Add(txtFiltro);

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
            topPanel.Controls.Add(btnFiltrar);

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
            topPanel.Controls.Add(btnLimparFiltro);

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
            topPanel.Controls.Add(btnFiltrarEnviados);

            // Botão Atualizar
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
            topPanel.Controls.Add(btnAtualizar);

            // Ajustar posição do DataGridView
            dataGridView1.Location = new System.Drawing.Point(10, 140);
            dataGridView1.Size = new System.Drawing.Size(780, 320);

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
            EstilizarBotao(Bt_Email, "Solicitar Documentação");
            EstilizarBotao(Bt_Validades, "Consulta");
            EstilizarBotao(Bt_Avisos, "Alerta de Caducidade");
            EstilizarBotao(Bt_imprimir, "Exportar");
            EstilizarBotao(Bt_imprimir2, "Exportar TESTE");
            EstilizarBotao(BT_ImprimirJPA, "Exportar JPA");
            EstilizarBotao(BT_CriarTrabalhadores, "Criar Trabalhadores");


            // Adicionar ToolTip nos botões
            toolTip.SetToolTip(Bt_Email, "Clique aqui para solicitar a documentação necessária para a entrada em obra das subempreitadas selecionadas, por email.");
            toolTip.SetToolTip(Bt_Validades, "Clique aqui para consultar as subempreitadas selecionadas.");
            toolTip.SetToolTip(Bt_Avisos, "Clique aqui para alertar sobre documentos caducados das subempreitadas selecionadas, por email.");
            toolTip.SetToolTip(Bt_imprimir, "Clique aqui para Imprimir das subempreitadas selecionadas.");
            toolTip.SetToolTip(Bt_imprimir2, "Clique aqui para Imprimir das subempreitadas selecionadas.");
            toolTip.SetToolTip(BT_ImprimirJPA, "Clique aqui para Imprimir das subempreitadas selecionadas.");

            // Adicionar painel inferior com informações ou estatísticas
            System.Windows.Forms.Panel bottomPanel = new System.Windows.Forms.Panel
            {
                Height = 30,
                Dock = DockStyle.Bottom,  // Colocando no fundo
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




            dataGridView1.UseWaitCursor = false;
            dataGridView1.Cursor = Cursors.Default;
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
            botao.Size = new System.Drawing.Size(150, 28);



            // Adicionar efeito de hover ao botão
            botao.MouseEnter += (s, e) =>
            {
                botao.BackColor = System.Drawing.Color.FromArgb(59, 89, 152);
                botao.ForeColor = System.
           Drawing.Color.White;
            };
            botao.MouseLeave += (s, e) =>
            {
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
                dataTable.Columns.Add(" ", typeof(bool));
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


                    dataTable.Rows.Add(id, false, nome, emailEnviado, dataEnvio, auto, caducado);

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

                dataGridView1.Columns["Nome"].ReadOnly = true;
                dataGridView1.Columns["DataEnvioColumn"].ReadOnly = true;
                dataGridView1.Columns["EmailEnviadoColumn"].ReadOnly = true;
                dataGridView1.Columns["Autorizado Em Obra"].ReadOnly = true;
                dataGridView1.Columns["Documentos Expirados"].ReadOnly = true;

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

                this.Cursor = Cursors.Default;
                dataGridView1.Cursor = Cursors.Default;
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
(CDU_ValidadeReciboSeguroAT < GETDATE() AND CDU_ValidadeReciboSeguroAT IS NOT NULL) OR
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
                bool enviado = false; // Variável para controlar se algum e-mail foi enviado
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Verificando se a coluna 'Selecione' está marcada como true
                    if (row.Cells[" "].Value != null && (bool)row.Cells[" "].Value)
                    {
                        string idSelecionado = row.Cells["ID"].Value?.ToString();
                        string nome = row.Cells["Nome"].Value?.ToString();

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

                            if (resultado == DialogResult.Yes)
                            {
                                CriarEmail criarEmail = new CriarEmail(BSO, idSelecionado);
                                if (criarEmail.ShowDialog() == DialogResult.OK)
                                {
                                    email = criarEmail.Email;
                                    var updateentidadeemail = $@"INSERT INTO Geral_Entidade_Contactos (ID, EntidadeID, Email, TipoContacto, Contacto)
                                                        VALUES (
                                                            NEWID(),
                                                            CAST('{idSelecionado}' AS UNIQUEIDENTIFIER),
                                                            '{email}',
                                                            'Geral',
                                                            '219999999'  -- ou outro valor de contacto obrigatório
                                                        );
                                                        ";
                                    BSO.DSO.ExecuteSQL(updateentidadeemail);
                                }





                            }
                        }


                        // Se o e-mail for válido, iniciar o Outlook e criar o e-mail
                        if (!string.IsNullOrEmpty(email))
                        {
                            // Iniciando o Outlook
                            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                            MailItem emailItem = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);


                            // Definindo o assunto e o corpo do e-mail
                            emailItem.Subject = "Documentação para entrada obra";
                            emailItem.Body = $@"Exmos. Senhores,

No seguimento da indicação de que será subempreiteiro da JPA CONSTRUTORA na empreitada supracitada, solicitamos o envio/anexação da documentação referente à Vossa empresa, aos Vossos colaboradores e aos Equipamentos previstos para a empreitada, conforme a listagem abaixo.

Para colocar a documentação solicitada, por favor, aceda ao seguinte link:
{link}
📁 DOCUMENTAÇÃO RELATIVA À EMPRESA

Alvará / Certificado de Empreiteiro de Obras Públicas
Apólice de Seguro de Acidentes de Trabalho
Último recibo pago do Seguro de Acidentes de Trabalho
Apólice de Seguro de Responsabilidade Civil
Último recibo pago do Seguro de Responsabilidade Civil
Folha de Remunerações da Segurança Social atualizada (com nomes dos trabalhadores e comprovativo de pagamento da TSU)
Inscrição na Segurança Social de trabalhadores novos (caso não constem na última folha de remunerações)
Declaração de não dívida às Finanças
Declaração de não dívida à Segurança Social
Horário de Trabalho (com indicação do nome e local da empreitada)
Declaração de trabalhadores emigrantes (se aplicável)
Declaração de aceitação do PSS ou PTRE (para subempreitadas)

👷‍♂️ DOCUMENTAÇÃO RELATIVA AOS TRABALHADORES

Registo de Colaborador:
- N.º do B.I./Cartão de Cidadão ou Título de Residência (e validade)
- N.º de Contribuinte
- N.º de Segurança Social
- Ficha de Aptidão Médica
- Ficha de Equipamentos de Proteção Individual (EPI’s) com validade inferior a 2 anos
- Formação específica
- Passaporte e Visto de Permanência ou Manifestação de Interesse atualizados – trabalhadores estrangeiros
- Contrato de trabalho com carimbo da ACT – obrigatório para trabalhadores estrangeiros

🛠️ DOCUMENTAÇÃO RELATIVA A EQUIPAMENTOS

- Declaração CE de conformidade e registo de marcação CE
- Manual de Instruções em Português
- Registo de manutenção e revisão
- Lista de verificação conforme Decreto Lei nº 50/2005 de 25 de fevereiro
- Seguro do equipamento/Seguro de Responsabilidade Civil atualizado
- Formação em manobrador (para operadores de equipamentos)

⚠️ NOTA IMPORTANTE

A documentação deverá ser enviada obrigatoriamente até 48 horas antes da entrada em obra.

🛑 PROIBIÇÕES NO LOCAL DE TRABALHO

- É proibido o consumo de bebidas alcoólicas durante o período e no local de trabalho.
- Não é permitida a presença com taxa de álcool ≥ 0,5g/L, nem sob influência de estupefacientes.

✅ RECOMENDAÇÕES BÁSICAS DE HST A CUMPRIR EM OBRA

- Apenas técnicos abrangidos por seguro de acidentes de trabalho e com Ficha de Aptidão Médica válida podem estar em obra.
- Utilização obrigatória de EPCs e EPIs conforme indicado.
- Riscos e medidas preventivas devem ser comunicados a todos os trabalhadores.
- Trabalhos com riscos especiais não podem ser executados isoladamente.
- Equipamentos e máquinas devem ser operados conforme o DL 50/2005.
- Escadas devem estar certificadas e em bom estado.
- Todos devem conhecer os procedimentos de emergência.
- Estaleiro deve ser mantido em ordem e salubridade.
- Subempreiteiros são responsáveis por eliminar, reciclar ou evacuar resíduos e entulhos.

🔧 INSTRUÇÕES ADICIONAIS PARA SUBEMPREITEIROS

A equipa do subempreiteiro deverá cumprir as obrigações previstas no Artigo 22.º do Decreto-Lei aplicável, e assegurar que eventuais empresas subcontratadas também o façam.
Caso existam trabalhadores independentes, aplica-se igualmente o Artigo 23.º do mesmo decreto.
";


                            // Definindo o e-mail do destinatário
                            emailItem.To = email;

                            // Abre o Outlook para o usuário revisar o e-mail antes de enviar
                            emailItem.Display();

                            // Marcando que um e-mail foi enviado para esta linha
                            enviado = true;

                            // Atualizando os campos na tabela após o envio
                            string updateQuery = $@"
                        UPDATE Geral_Entidade 
                        SET CDU_EmailEnviado = 1, CDU_DataEnvio = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}'
                        WHERE id = '{idSelecionado}'";
                            BSO.DSO.ExecuteSQL(updateQuery);
                        }
                    }
                }

                // Caso nenhum e-mail tenha sido enviado, avisamos o usuário
                if (!enviado)
                {
                    MessageBox.Show("Nenhuma linha selecionada ou nenhum e-mail foi enviado.");
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
            //   dataGridView1.DataSource = dataOriginal;
            DadosLista();
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
                List<string> idsSelecionados = new List<string>();

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[" "].Value != null && (bool)row.Cells[" "].Value)
                    {
                        string id = row.Cells["ID"].Value?.ToString();
                        if (!string.IsNullOrEmpty(id))
                            idsSelecionados.Add(id);
                    }
                }

                if (idsSelecionados.Count > 0)
                {

                    Validades menuForm = new Validades(BSO, PSO, idsSelecionados);
                    menuForm.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Por favor, selecione pelo menos uma empresa com a caixa marcada.");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao editar: " + ex.Message);
            }
        }

        private void Bt_Avisos_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells[" "].Value) == true)
                {
                    string id = row.Cells["id"].Value.ToString();
                    string nomeEntidade = row.Cells["Nome"].Value.ToString();

                    // Verifica se deve ignorar os alertas
                    var ignoraAlerta = BSO.Consulta($"SELECT CDU_IgnoraAlerta FROM Geral_Entidade WHERE ID = '{id}'").DaValor<int>("CDU_IgnoraAlerta");
                    if (ignoraAlerta == 1)
                        continue;

                    List<string> documentosEmpresa = VerificaDocumentosDetalhados(id);
                    Dictionary<string, List<string>> documentosTrabalhadores = VerificaDocumentosTrabalhadores(id);
                    Dictionary<string, List<string>> documentosEquipamentos = VerificaDocumentosEquipamentos(id);
                    Dictionary<string, List<string>> documentosAutorizacoes = VerificaDocumentosAutorizacoes(id);
                    var link = BSO.Consulta($"SELECT CDU_Link FROM Geral_Entidade WHERE ID = '{id}'").DaValor<string>("CDU_Link");

                    if (documentosEmpresa.Count > 0 ||
                        documentosTrabalhadores.Count > 0 ||
                        documentosEquipamentos.Count > 0 ||
                        documentosAutorizacoes.Count > 0)
                    {
                        StringBuilder corpo = new StringBuilder();
                        corpo.AppendLine("Prezado(a),");
                        corpo.AppendLine();
                        corpo.AppendLine($"A entidade \"{nomeEntidade}\" tem documentos caducados.");
                        corpo.AppendLine("Para colocar a documentação solicitada, por favor, aceda ao seguinte link:");
                        corpo.AppendLine(link);
                        corpo.AppendLine();

                        if (documentosEmpresa.Any())
                        {
                            corpo.AppendLine("📁 Documentos da Empresa:");
                            foreach (var doc in documentosEmpresa)
                            {
                                corpo.AppendLine($"- {doc}");
                            }
                            corpo.AppendLine();
                        }

                        if (documentosTrabalhadores.Any())
                        {
                            corpo.AppendLine("👷 Documentos por Trabalhador:");
                            foreach (var trabalhador in documentosTrabalhadores)
                            {
                                corpo.AppendLine($"\n{trabalhador.Key}:");
                                foreach (var doc in trabalhador.Value)
                                {
                                    corpo.AppendLine($"- {doc}");
                                }
                            }
                            corpo.AppendLine();
                        }

                        if (documentosEquipamentos.Any())
                        {
                            corpo.AppendLine("🔧 Documentos de Equipamentos:");
                            foreach (var equipamento in documentosEquipamentos)
                            {
                                corpo.AppendLine($"\n{equipamento.Key}:");
                                foreach (var doc in equipamento.Value)
                                {
                                    corpo.AppendLine($"- {doc}");
                                }
                            }
                            corpo.AppendLine();
                        }

                        if (documentosAutorizacoes.Any())
                        {
                            corpo.AppendLine("🔑 Documentos de Autorizações:");
                            foreach (var autorizacao in documentosAutorizacoes)
                            {
                                corpo.AppendLine($"\n{autorizacao.Key}:");
                                foreach (var doc in autorizacao.Value)
                                {
                                    corpo.AppendLine($"- {doc}");
                                }
                            }
                            corpo.AppendLine();
                        }

                        corpo.AppendLine("\nPor favor, regularize esta situação com urgência.");
                        corpo.AppendLine("\nObrigado.");

                        Outlook.Application outlookApp = new Outlook.Application();
                        Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                        mailItem.To = "departamento@email.pt";
                        mailItem.Subject = $"Alerta Documentos Caducados - {nomeEntidade}";

                        mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                        mailItem.Display(); // Gera a assinatura

                        string existingBody = mailItem.HTMLBody;
                        string customBody = corpo.ToString().Replace("\n", "<br>");

                        mailItem.HTMLBody = customBody + "<br><br>" + existingBody;

                        mailItem.Display(); // Mostra o email com a assinatura incluída
                    }
                    else
                    {
                        //MessageBox.Show($"Não há documentos caducados para a entidade \"{nomeEntidade}\".");
                    }
                }
            }
        }


        private List<string> VerificaDocumentosDetalhados(string id)
        {
            List<string> caducados = new List<string>();

            // Reutilizar a lógica da tua função anterior, mas guardar os nomes dos documentos caducados
            var campos = new Dictionary<string, string>()
            {
                {"CDU_ValidadeFinancas", "Finanças"},
                {"CDU_ValidadeSegSocial", "Segurança Social"},
                {"CDU_ValidadeFolhaPag", "Folha de Pagamento"},
                {"CDU_ValidadeComprovativoPagamento", "Comprovativo de Pagamento"},
                {"CDU_ValidadeReciboSeguroAT", "Seguro AT"},
                {"CDU_ValidadeSeguroRC", "Seguro RC"},
                {"CDU_ValidadeSeguroAT", "condições Seguro AT"},
                {"CDU_ValidadeAlvara", "Alvará"},
                {"CDU_ValidadeCertidaoPermanente", "Certidão Permanente"}
            };

            string query = $"SELECT {string.Join(",", campos.Keys)} FROM Geral_Entidade WHERE ID = '{id}'";
            var res = BSO.Consulta(query);


            res.Inicio();

            foreach (var campo in campos)
            {
                DateTime validade;
                if (DateTime.TryParse(res.DaValor<string>(campo.Key), out validade))
                {
                    if (validade < DateTime.Now && validade != DateTime.MinValue)
                    {
                        caducados.Add(campo.Value);
                    }
                }
            }

            // Podes replicar esta mesma lógica para trabalhadores, equipamentos e autorizações se quiseres mais detalhe

            return caducados;
        }

        private Dictionary<string, List<string>> VerificaDocumentosTrabalhadores(string idEmpresa)
        {
            var resultado = new Dictionary<string, List<string>>();

            var camposTrabalhadores = new Dictionary<string, string>()
    {
        {"caminho1", "Cartão de cidadão ou residência"},
        {"caminho2", "Ficha Medica"},
        {"caminho3", "Credenciacao"},
        {"caminho4", "Trabalhoss especializados"},
        {"caminho5", "Ficha Destribuiçao"}
    };

            // Supondo que tens um campo com o nome ou identificador do trabalhador
            string querytrab = $@"SELECT Nome, {string.Join(",", camposTrabalhadores.Keys)} FROM TDU_AD_Trabalhadores WHERE id_empresa = '{idEmpresa}'";
            var resTrab = BSO.Consulta(querytrab);

            resTrab.Inicio();
            for (int i = 0; i < resTrab.NumLinhas(); i++)
            {
                string nomeTrab = resTrab.DaValor<string>("nome");

                var documentosCaducados = new List<string>();

                foreach (var campo in camposTrabalhadores)
                {
                    string valorOriginal = resTrab.DaValor<string>(campo.Key);
                    if (string.IsNullOrWhiteSpace(valorOriginal)) continue;

                    string valorDecodificado = WebUtility.HtmlDecode(valorOriginal);

                    var match = Regex.Match(valorDecodificado, @"\d{2}[\/\-]\d{2}[\/\-]\d{4}");
                    if (match.Success)
                    {
                        if (DateTime.TryParse(match.Value, out DateTime validade))
                        {
                            if (validade < DateTime.Now && validade != DateTime.MinValue)
                            {
                                documentosCaducados.Add(campo.Value);
                            }
                        }
                    }
                }

                if (documentosCaducados.Any())
                {
                    resultado[nomeTrab] = documentosCaducados;
                }

                resTrab.Seguinte(); // move para o próximo registo
            }


            return resultado;
        }

        private Dictionary<string, List<string>> VerificaDocumentosEquipamentos(string idEmpresa)
        {
            var resultado = new Dictionary<string, List<string>>();

            var camposEquipamentos = new Dictionary<string, string>()
    {
        {"caminho5", "Outro Documento Relevante"}
    };

            string queryEquip = $@"SELECT marca, {string.Join(",", camposEquipamentos.Keys)} FROM TDU_AD_Equipamentos WHERE id_empresa = '{idEmpresa}'";
            var resEquip = BSO.Consulta(queryEquip);

            resEquip.Inicio();
            for (int i = 0; i < resEquip.NumLinhas(); i++)
            {
                string nomeEquip = resEquip.DaValor<string>("marca")?.Trim();
                if (string.IsNullOrEmpty(nomeEquip)) nomeEquip = "(Sem Nome)";

                var documentosCaducados = new List<string>();

                foreach (var campo in camposEquipamentos)
                {
                    string valorOriginal = resEquip.DaValor<string>(campo.Key);
                    if (string.IsNullOrWhiteSpace(valorOriginal)) continue;

                    string valorDecodificado = WebUtility.HtmlDecode(valorOriginal);

                    var match = Regex.Match(valorDecodificado, @"\d{2}[\/\-]\d{2}[\/\-]\d{4}");
                    if (match.Success)
                    {
                        if (DateTime.TryParse(match.Value, out DateTime validade))
                        {
                            if (validade < DateTime.Now && validade != DateTime.MinValue)
                            {
                                documentosCaducados.Add(campo.Value);
                            }
                        }
                    }
                }

                if (documentosCaducados.Any())
                {
                    resultado[nomeEquip] = documentosCaducados;
                }

                resEquip.Seguinte();
            }

            return resultado;
        }

        private Dictionary<string, List<string>> VerificaDocumentosAutorizacoes(string idEmpresa)
        {
            var resultado = new Dictionary<string, List<string>>();

            var camposAutorizacoes = new Dictionary<string, string>()
            {
                {"caminho1", "Contrato"},
                {"caminho2", "Horário de trabalho da empreitada"},
                {"caminho3", "Declaração de adesão ao PSS"},
                {"caminho4", "Declaração do resposável no estaleiro"}
                // Adicione mais campos conforme necessário para as autorizações.
            };

            // Supondo que tens um campo com o nome ou identificador da autorização
            string queryAutorizacoes = $@"SELECT Codigo_Obra, {string.Join(",", camposAutorizacoes.Keys)} FROM TDU_AD_Autorizacoes WHERE ID_Entidade = '{idEmpresa}'";
            var resAutorizacoes = BSO.Consulta(queryAutorizacoes);

            resAutorizacoes.Inicio();
            for (int i = 0; i < resAutorizacoes.NumLinhas(); i++)
            {
                string nomeAutorizacao = resAutorizacoes.DaValor<string>("Codigo_Obra");

                var documentosCaducados = new List<string>();

                foreach (var campo in camposAutorizacoes)
                {
                    string valorOriginal = resAutorizacoes.DaValor<string>(campo.Key);
                    if (string.IsNullOrWhiteSpace(valorOriginal)) continue;

                    string valorDecodificado = WebUtility.HtmlDecode(valorOriginal);

                    var match = Regex.Match(valorDecodificado, @"\d{2}[\/\-]\d{2}[\/\-]\d{4}");
                    if (match.Success)
                    {
                        if (DateTime.TryParse(match.Value, out DateTime validade))
                        {
                            if (validade < DateTime.Now && validade != DateTime.MinValue)
                            {
                                documentosCaducados.Add(campo.Value);
                            }
                        }
                    }
                }

                if (documentosCaducados.Any())
                {
                    resultado[nomeAutorizacao] = documentosCaducados;
                }

                resAutorizacoes.Seguinte(); // move para o próximo registo
            }

            return resultado;
        }

        private void EnviarEmailOutlook(string destinatario, string assunto, string corpo)
        {
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.To = destinatario;
            mailItem.Subject = assunto;
            mailItem.Body = corpo;
            mailItem.Display(); // Mostra o Outlook com o email preenchido, mas não envia automaticamente
        }

        private void Bt_imprimir_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> idsSelecionados = new List<string>();

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[" "].Value != null && (bool)row.Cells[" "].Value)
                    {
                        string id = row.Cells["ID"].Value?.ToString();
                        if (!string.IsNullOrEmpty(id))
                            idsSelecionados.Add(id);
                    }
                }



                if (idsSelecionados.Count > 0)
                {


                    // 1. Cria uma cópia da lista original para verificação de autorização
                    var idsParaVerificacao = new List<string>(idsSelecionados);
                    idsParaVerificacao.Remove("2A8C7ECD-309B-49F9-A337-203B45CED948"); // remove se estiver por algum motivo

                    // 2. Verifica autorização sem o id padrão
                    Dictionary<string, List<string>> autorizacoes;
                    string obraComum;
                    var autorizado = VerificaAutorizacao(idsParaVerificacao, out autorizacoes, out obraComum);
                    if (!autorizado)
                    {
                        return;
                    }

                    // 3. Adiciona o ID padrão no início da lista (se ainda não estiver)
                    string idPadrao = "2A8C7ECD-309B-49F9-A337-203B45CED948";
                    if (!idsSelecionados.Contains(idPadrao))
                    {
                        idsSelecionados.Insert(0, idPadrao); // insere na primeira posição
                    }
                    else
                    {
                        // opcional: move para o início se já existir em outra posição
                        idsSelecionados.Remove(idPadrao);
                        idsSelecionados.Insert(0, idPadrao);
                    }

                    // 4. Continua com a exportação
                    //ExportarParaExcelNovo(idsSelecionados, obraComum);
                    ExportarParaExcel(idsSelecionados, obraComum);
                }
                else
                {
                    MessageBox.Show("Por favor, selecione pelo menos uma empresa com a caixa marcada.");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao exportar para Excel: " + ex.Message);
            }
        }
        private void Bt_imprimir2_Click(object sender, EventArgs e)
        {

            try
            {
                List<string> idsSelecionados = new List<string>();

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[" "].Value != null && (bool)row.Cells[" "].Value)
                    {
                        string id = row.Cells["ID"].Value?.ToString();
                        if (!string.IsNullOrEmpty(id))
                            idsSelecionados.Add(id);
                    }
                }



                if (idsSelecionados.Count > 0)
                {


                    // 1. Cria uma cópia da lista original para verificação de autorização
                    var idsParaVerificacao = new List<string>(idsSelecionados);
                    idsParaVerificacao.Remove("2A8C7ECD-309B-49F9-A337-203B45CED948"); // remove se estiver por algum motivo

                    // 2. Verifica autorização sem o id padrão
                    Dictionary<string, List<string>> autorizacoes;
                    string obraComum;
                    var autorizado = VerificaAutorizacao(idsParaVerificacao, out autorizacoes, out obraComum);
                    if (!autorizado)
                    {
                        return;
                    }

                    // 3. Adiciona o ID padrão no início da lista (se ainda não estiver)
                    string idPadrao = "2A8C7ECD-309B-49F9-A337-203B45CED948";
                    if (!idsSelecionados.Contains(idPadrao))
                    {
                        idsSelecionados.Insert(0, idPadrao); // insere na primeira posição
                    }
                    else
                    {
                        // opcional: move para o início se já existir em outra posição
                        idsSelecionados.Remove(idPadrao);
                        idsSelecionados.Insert(0, idPadrao);
                    }

                    // 4. Continua com a exportação
                    ExportarParaExcelNovo(idsSelecionados, obraComum);
                    //ExportarParaExcel(idsSelecionados, obraComum);
                }
                else
                {
                    MessageBox.Show("Por favor, selecione pelo menos uma empresa com a caixa marcada.");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao exportar para Excel: " + ex.Message);
            }
        }
        private bool VerificaAutorizacao(List<string> idsSelecionados, out Dictionary<string, List<string>> autorizacoes, out string obraComum)
        {
            List<string> semAutorizacao = new List<string>();
            autorizacoes = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> obrasPaiPorEntidade = new Dictionary<string, List<string>>();

            foreach (string id in idsSelecionados)
            {
                var result = BSO.Consulta($"SELECT Codigo_Obra FROM TDU_AD_Autorizacoes WHERE ID_Entidade = '{id}'");
                int num = result.NumLinhas();

                if (num == 0)
                {
                    semAutorizacao.Add(id);
                }
                else
                {
                    List<string> obras = new List<string>();
                    List<string> obrasPai = new List<string>();

                    result.Inicio();
                    for (int i = 1; i <= num; i++)
                    {
                        string codigoObra = result.DaValor<string>("Codigo_Obra");
                        obras.Add(codigoObra);

                        // Consulta ObraPaiID da obra
                        var resultPai = BSO.Consulta($"SELECT ObraPaiID FROM COP_Obras WHERE Codigo = '{codigoObra}'");
                        if (resultPai.NumLinhas() > 0)
                        {
                            resultPai.Inicio();
                            string obraPaiId = resultPai.DaValor<string>("ObraPaiID");
                            obrasPai.Add(obraPaiId);
                        }

                        result.Seguinte();
                    }

                    autorizacoes[id] = obras;
                    obrasPaiPorEntidade[id] = obrasPai;
                }
            }

            if (semAutorizacao.Count > 0)
            {
                string msg = "As seguintes entidades não possuem autorização em nenhuma obra:\n" +
                             string.Join("\n", semAutorizacao);
                MessageBox.Show(msg, "Autorização Ausente", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                obraComum = null;
                return false;
            }

            // Verificar se todas as entidades compartilham pelo menos um ObraPaiID em comum
            var obrasPaiComuns = obrasPaiPorEntidade.Values.Aggregate((prev, next) => prev.Intersect(next).ToList());

            if (obrasPaiComuns.Count == 0)
            {
                MessageBox.Show("As entidades selecionadas possuem autorizações, mas não têm nenhuma obra pai em comum.", "Obras Divergentes", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                obraComum = null;
                return false;
            }

            // Se tiverem um ObraPaiID em comum, agora buscamos uma das obras associadas a ele (poderia ser a primeira)
            string obraPaiIdComum = obrasPaiComuns.First();

            var resultObra = BSO.Consulta($"SELECT Codigo FROM COP_Obras WHERE ObraPaiID = '{obraPaiIdComum}'");
            if (resultObra.NumLinhas() > 0)
            {
                resultObra.Inicio();
                obraComum = resultObra.DaValor<string>("Codigo");
            }
            else
            {
                obraComum = null;
            }

            return true;
        }




        private void ExportarParaExcel(List<string> idsSelecionados, string codigoObra)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add();
                var numidsempresa = idsSelecionados.Count;
                // Buscar dados reais da obra
                string queryObra = $@"SELECT COP.EntidadeIDA,GE.Nome,COP.CDU_LocalObra FROM COP_Obras AS COP
                                    INNER JOIN Geral_Entidade AS GE ON COP.EntidadeIDA = GE.EntidadeId
                                    WHERE COP.Codigo = '{codigoObra}'";
                var dadosObra = BSO.Consulta(queryObra);
                string descricaoObra = "", donoObra = "", entidadeExecutante = "";
                if (!dadosObra.Vazia())
                {
                    dadosObra.Inicio();
                    descricaoObra = dadosObra.Valor("CDU_LocalObra")?.ToString() ?? "";
                    donoObra = dadosObra.Valor("Nome")?.ToString() ?? "";
                    //entidadeExecutante = dadosObra.Valor("EntidadeExecutante")?.ToString() ?? "";
                }
                // Obter dados da empresa
                string idsFormatados = string.Join(",", idsSelecionados.Select(id => $"'{id}'"));

                string queryEmpresa = $"SELECT Nome FROM Geral_Entidade WHERE id IN ({idsFormatados})";

                StdBELista dtEmpresa = BSO.Consulta(queryEmpresa);
                string nomeEmpresa = "";

                dtEmpresa.Inicio();
                if (!dtEmpresa.NoFim())
                {
                    nomeEmpresa = dtEmpresa.Valor("Nome")?.ToString() ?? "";
                }

                // Criar nova folha para cada empresa
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.Add();
                // Limitar o nome da folha a 31 caracteres e remover caracteres inválidos
                string nomeEmpresaLimpo = nomeEmpresa.Replace("/", "").Replace("\\", "").Replace("?", "").Replace("*", "").Replace("[", "").Replace("]", "").Replace(":", "");
                string nomeFolha = $"Empresa";
                if (nomeFolha.Length > 31)
                {
                    nomeFolha = nomeFolha.Substring(0, 31);
                }
                worksheet.Name = nomeFolha;

                int linhaAtual = 1;

                // Adicionar cabeçalho da obra no topo da folha
                worksheet.Cells[linhaAtual, 1] = $"OBRA: {descricaoObra}";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Merge();
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Size = 14;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                linhaAtual += 2;

                worksheet.Cells[linhaAtual, 1] = $"DONO DE OBRA: {donoObra}";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Merge();
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Bold = true;
                linhaAtual += 2;

                worksheet.Cells[linhaAtual, 1] = "ENTIDADE EXECUTANTE: JOAQUIM PEIXOTO AZEVEDO & FILHOS LDA";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Merge();
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Bold = true;
                linhaAtual += 3;
                // Cabeçalhos principais
                worksheet.Cells[linhaAtual, 1] = "EMPRESA";
                worksheet.Cells[linhaAtual, 4] = "Alvará";
                worksheet.Cells[linhaAtual, 5] = "Contribuinte";
                worksheet.Cells[linhaAtual, 6] = "Não Div. Finanças";
                worksheet.Cells[linhaAtual, 7] = "Não Div. Seg. Social";
                worksheet.Cells[linhaAtual, 8] = "Folha Pag. Seg. Social";
                worksheet.Cells[linhaAtual, 9] = "Recibo de Pag. Seg. Social";
                worksheet.Cells[linhaAtual, 10] = "Apólice AT";
                worksheet.Cells[linhaAtual, 11] = "Recibo Apólice AT";
                worksheet.Cells[linhaAtual, 12] = "Apólice RC";
                worksheet.Cells[linhaAtual, 13] = "Recibo RC";
                worksheet.Cells[linhaAtual, 14] = "Horário de Trabalho";
                worksheet.Cells[linhaAtual, 15] = "Condições do seguro de responsabilidade civil";


                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 15]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 15]].Interior.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                linhaAtual++;

                // Sub-cabeçalhos
                worksheet.Cells[linhaAtual, 1] = "N.º";
                worksheet.Cells[linhaAtual, 2] = "Nome";
                worksheet.Cells[linhaAtual, 3] = "Sede";
                worksheet.Cells[linhaAtual, 4] = "N.º";
                worksheet.Cells[linhaAtual, 5] = "N.º";
                worksheet.Cells[linhaAtual, 6] = "Validade";
                worksheet.Cells[linhaAtual, 7] = "Validade";
                worksheet.Cells[linhaAtual, 8] = "Validade";
                worksheet.Cells[linhaAtual, 9] = "C; N/C; N/A";
                worksheet.Cells[linhaAtual, 10] = "C ; N/C ; N/A";
                worksheet.Cells[linhaAtual, 11] = "Validade";
                worksheet.Cells[linhaAtual, 12] = "C; N/C; N/A";
                worksheet.Cells[linhaAtual, 13] = "Validade";
                worksheet.Cells[linhaAtual, 14] = "C ; N/C ; N/A";
                worksheet.Cells[linhaAtual, 15] = "C ; N/C ; N/A";

                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 15]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 15]].Interior.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                linhaAtual++;

                // Preencher empresas
                int numeroEmpresa = 1;

                foreach (string id in idsSelecionados)
                {
                    string query = $@"SELECT 
            Nome,AlvaraNumero, NIPC, Morada, CDU_ValidadeAlvara, CDU_ValidadeFinancas, CDU_ValidadeSegSocial, CDU_ValidadeFolhaPag,
            CDU_ValidadeComprovativoPagamento, CDU_ValidadeReciboSeguroAT, CDU_ValidadeSeguroRC,
            CDU_ValidadeHorarioTrabalho, CDU_ValidadeSeguroAT, CDU_ValidadeSeguroResposabilidadeCivil
            FROM Geral_Entidade WHERE id = '{id}'";

                    StdBELista empresa = BSO.Consulta(query);
                    empresa.Inicio();








                    if (!empresa.NoFim())
                    {
                        string nome = empresa.Valor("Nome")?.ToString() ?? "";
                        string alvara = empresa.Valor("AlvaraNumero")?.ToString() ?? "";
                        string nif = empresa.Valor("NIPC")?.ToString() ?? "";
                        string morada = empresa.Valor("Morada")?.ToString() ?? "";
                        if (id == "2A8C7ECD-309B-49F9-A337-203B45CED948")
                        {

                            worksheet.Cells[linhaAtual, 1] = numeroEmpresa;
                            worksheet.Cells[linhaAtual, 2] = nome;
                            worksheet.Cells[linhaAtual, 3] = morada;
                            worksheet.Cells[linhaAtual, 4] = alvara;
                            worksheet.Cells[linhaAtual, 5] = nif;
                            worksheet.Cells[linhaAtual, 6] = "C";
                            worksheet.Cells[linhaAtual, 7] = "C";
                            worksheet.Cells[linhaAtual, 8] = "C";
                            worksheet.Cells[linhaAtual, 9] = "C";
                            worksheet.Cells[linhaAtual, 10] = "C";
                            worksheet.Cells[linhaAtual, 11] = "C";
                            worksheet.Cells[linhaAtual, 12] = "C";
                            worksheet.Cells[linhaAtual, 13] = "C";
                            worksheet.Cells[linhaAtual, 14] = "C";
                            worksheet.Cells[linhaAtual, 15] = "C";


                            linhaAtual++;
                            numeroEmpresa++;
                            continue; // pula para o próximo ID
                        }






                        DateTime.TryParse(empresa.Valor("CDU_ValidadeAlvara")?.ToString(), out DateTime validadeAlvara);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeFinancas")?.ToString(), out DateTime validadeFinancas);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeSegSocial")?.ToString(), out DateTime validadeSegSocial);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeFolhaPag")?.ToString(), out DateTime validadeFolhaPag);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeComprovativoPagamento")?.ToString(), out DateTime validadeComprovativoPagamento);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeReciboSeguroAT")?.ToString(), out DateTime validadeReciboSeguroAT);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeSeguroRC")?.ToString(), out DateTime validadeSeguroRC);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeHorarioTrabalho")?.ToString(), out DateTime validadeHorarioTrabalho);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeSeguroAT")?.ToString(), out DateTime validadeSeguroAT);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeSeguroResposabilidadeCivil")?.ToString(), out DateTime ValidadeSeguroResposabilidadeCivil);

                        worksheet.Cells[linhaAtual, 1] = numeroEmpresa;
                        worksheet.Cells[linhaAtual, 2] = nome;
                        worksheet.Cells[linhaAtual, 3] = morada;
                        worksheet.Cells[linhaAtual, 4] = alvara;
                        worksheet.Cells[linhaAtual, 5] = nif;
                        worksheet.Cells[linhaAtual, 6] = validadeAlvara.Year > 1 ? validadeAlvara.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 7] = validadeFinancas.Year > 1 ? validadeFinancas.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 8] = validadeSegSocial.Year > 1 ? validadeSegSocial.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 9] = validadeFolhaPag.Year > 1 ? validadeFolhaPag.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 10] = validadeComprovativoPagamento.Year > 1 ? "C" : "NC";
                        worksheet.Cells[linhaAtual, 11] = validadeSeguroAT.Year > 1 ? validadeSeguroAT.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 12] = validadeReciboSeguroAT.Year > 1 ? "C" : "NC";
                        worksheet.Cells[linhaAtual, 13] = validadeSeguroRC.Year > 1 ? validadeSeguroRC.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 14] = validadeHorarioTrabalho.Year > 1 ? "C" : "NC";
                        worksheet.Cells[linhaAtual, 15] = ValidadeSeguroResposabilidadeCivil.Year > 1 ? "C" : "NC";

                        linhaAtual++;
                        numeroEmpresa++;
                    }
                }

                linhaAtual += 2;



                // Dados dos Trabalhadores
                worksheet.Cells[linhaAtual, 1] = "TRABALHADORES";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Merge();
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                linhaAtual++;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                // Cabeçalhos dos trabalhadores
                worksheet.Cells[linhaAtual, 1] = "N.º"; // NOVA COLUNA
                worksheet.Cells[linhaAtual, 2] = "Nome Completo";
                worksheet.Cells[linhaAtual, 3] = "Empresa";
                worksheet.Cells[linhaAtual, 4] = "Categoria";
                worksheet.Cells[linhaAtual, 5] = "Contribuinte";
                worksheet.Cells[linhaAtual, 6] = "Nº Segurança Social";
                worksheet.Cells[linhaAtual, 7] = "Cartão de cidadão ou residencia";
                worksheet.Cells[linhaAtual, 8] = "Ficha de Aptidão para o Trabalho (FAT)";
                worksheet.Cells[linhaAtual, 9] = "Formação Profissional";
                worksheet.Cells[linhaAtual, 10] = "Trabalhos especializados";
                worksheet.Cells[linhaAtual, 11] = "Ficha de distribuição de EPI's";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Font.Bold = true;
                linhaAtual++;


                // Dados dos trabalhadores

                var queryTrabalhadoresExel = $@"
SELECT 
    t.nome,
    t.categoria,
    t.contribuinte,
    t.seguranca_social,
    t.email,
    t.anexo1,
    t.anexo2,
    t.anexo3,
    t.anexo4,
    t.anexo5,
    t.cBFormacaoProfissional,
    t.cBEspecializados,
    g.Nome AS nome_empresa
FROM 
    TDU_AD_Trabalhadores t
JOIN 
    Geral_Entidade g ON t.id_empresa = g.ID
WHERE 
    t.id_empresa IN ({idsFormatados});
";
                StdBELista dtTrabalhadores = BSO.Consulta(queryTrabalhadoresExel);


                int numeroTrabalhador = 1;

                dtTrabalhadores.Inicio();
                while (!dtTrabalhadores.NoFim())
                {


                    worksheet.Cells[linhaAtual, 1] = numeroTrabalhador; // N.º
                    worksheet.Cells[linhaAtual, 2] = dtTrabalhadores.Valor("nome")?.ToString() ?? "";
                    worksheet.Cells[linhaAtual, 3] = dtTrabalhadores.Valor("nome_empresa")?.ToString() ?? ""; ; // Empresa (pode preencher se quiser)
                    worksheet.Cells[linhaAtual, 4] = dtTrabalhadores.Valor("categoria")?.ToString() ?? "";
                    worksheet.Cells[linhaAtual, 5] = dtTrabalhadores.Valor("contribuinte")?.ToString() ?? "";
                    worksheet.Cells[linhaAtual, 6] = dtTrabalhadores.Valor("seguranca_social")?.ToString() ?? "";

                    var valorAnexo1 = dtTrabalhadores.Valor("anexo1")?.ToString();
                    worksheet.Cells[linhaAtual, 7] = valorAnexo1 == "True" ? "C" : "NC";

                    var valorAnexo2 = dtTrabalhadores.Valor("anexo2")?.ToString();
                    worksheet.Cells[linhaAtual, 8] = valorAnexo2 == "True" ? "C" : "NC";

                    // if o cBFormacaoProfissional for igual a NA coloca como NA se for igual a '' coloca NC se for igual a A coloca C
                    var valorcBFormacaoProfissional = dtTrabalhadores.Valor("cBFormacaoProfissional")?.ToString();
                    string valorFinal;

                    if (valorcBFormacaoProfissional == "NA")
                    {
                        valorFinal = "NA";
                    }
                    else if (string.IsNullOrEmpty(valorcBFormacaoProfissional))
                    {
                        valorFinal = "NC";
                    }
                    else if (valorcBFormacaoProfissional == "A")
                    {
                        valorFinal = "C";
                    }
                    else
                    {
                        // Caso queira tratar outros valores com "NC" por padrão
                        valorFinal = "NC";
                    }
                    worksheet.Cells[linhaAtual, 9] = valorFinal;


                    var valorcBEspecializados = dtTrabalhadores.Valor("cBEspecializados")?.ToString();
                    string valorFinalEspecializados;

                    if (valorcBEspecializados == "NA")
                    {
                        valorFinalEspecializados = "NA";
                    }
                    else if (string.IsNullOrEmpty(valorcBEspecializados))
                    {
                        valorFinalEspecializados = "NC";
                    }
                    else if (valorcBEspecializados == "A")
                    {
                        valorFinalEspecializados = "C";
                    }
                    else
                    {
                        // Caso queira tratar outros valores com "NC" por padrão
                        valorFinalEspecializados = "NC";
                    }

                    worksheet.Cells[linhaAtual, 10] = valorFinalEspecializados;

                    var valorAnexo5 = dtTrabalhadores.Valor("anexo5")?.ToString();
                    worksheet.Cells[linhaAtual, 11] = valorAnexo5 == "True" ? "C" : "NC";

                    linhaAtual++;
                    numeroTrabalhador++;
                    dtTrabalhadores.Seguinte();
                }

                var queryTrabalhadoresJPA = $@"SELECT COP.codigo ,COP_P.Funcionario,C.Descricao,F.NumContr,F.NumBeneficiario  FROM COP_Obras AS COP
   INNER JOIN COP_Obras_Pessoal AS COP_P ON COP.id = COP_P.ObraID 
   INNER JOIN GPR_Operadores AS O ON COP_P.ColaboradorID = O.IDOperador
   INNER JOIN Funcionarios AS F ON O.Funcionario = F.Codigo
   INNER JOIN Categorias AS C ON F.Categoria = C.Categoria
   WHERe COP.Codigo = '{ObraCodigo}'";



                StdBELista dtTrabalhadoresJPA = BSO.Consulta(queryTrabalhadoresJPA);
                dtTrabalhadoresJPA.Inicio();
                while (!dtTrabalhadoresJPA.NoFim())
                {
                    string funcionario = dtTrabalhadoresJPA.Valor("Funcionario")?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(funcionario))
                    {
                        string categoria = dtTrabalhadoresJPA.Valor("Descricao")?.ToString() ?? "N/A";
                        string numeroContr = dtTrabalhadoresJPA.Valor("NumContr")?.ToString() ?? "N/A";
                        string numeroBeneficiario = dtTrabalhadoresJPA.Valor("NumBeneficiario")?.ToString() ?? "N/A";

                        worksheet.Cells[linhaAtual, 1] = numeroTrabalhador;
                        worksheet.Cells[linhaAtual, 2] = funcionario;
                        worksheet.Cells[linhaAtual, 3] = "JPA";
                        worksheet.Cells[linhaAtual, 4] = categoria;
                        worksheet.Cells[linhaAtual, 5] = numeroContr;
                        worksheet.Cells[linhaAtual, 6] = numeroBeneficiario;
                        worksheet.Cells[linhaAtual, 7] = "C";
                        worksheet.Cells[linhaAtual, 8] = "C";
                        worksheet.Cells[linhaAtual, 9] = "C";
                        worksheet.Cells[linhaAtual, 10] = "C";
                        worksheet.Cells[linhaAtual, 11] = "C";

                        linhaAtual++;
                        numeroTrabalhador++;
                    }
                    dtTrabalhadoresJPA.Seguinte();
                }





                linhaAtual += 1;

                // Dados dos Equipamentos
                worksheet.Cells[linhaAtual, 1] = "EQUIPAMENTOS";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Merge(); // Atualizado para 9 colunas
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                linhaAtual++;

                // Cabeçalhos com N.º
                worksheet.Cells[linhaAtual, 1] = "N.º";
                worksheet.Cells[linhaAtual, 2] = "Marca";
                worksheet.Cells[linhaAtual, 3] = "Tipo";
                worksheet.Cells[linhaAtual, 4] = "Série";
                worksheet.Cells[linhaAtual, 5] = "Declaração de conformidade CE";
                worksheet.Cells[linhaAtual, 6] = "Lista de verificação conforme o Decreto-Lei n.º 50/2005";
                worksheet.Cells[linhaAtual, 7] = "Registos de Manutenção";
                worksheet.Cells[linhaAtual, 8] = "Manual de Instruções";
                worksheet.Cells[linhaAtual, 9] = "Seguro";

                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                linhaAtual++;


                // Consulta aos equipamentos
                string queryEquipamentos = $@"
                    SELECT marca, tipo, serie, anexo1, anexo2, anexo3, anexo4, anexo5 , cBConformidadeCE,cBDecreto_Lei,cBManutencao,cBManualInstrucoes,cBSeguro
                    FROM TDU_AD_Equipamentos 
                    WHERE id_empresa IN  ({idsFormatados})";

                StdBELista dtEquipamentos = BSO.Consulta(queryEquipamentos);

                int numeroEquipamento = 1;

                dtEquipamentos.Inicio();
                while (!dtEquipamentos.NoFim())
                {
                    worksheet.Cells[linhaAtual, 1] = numeroEquipamento; // N.º
                    worksheet.Cells[linhaAtual, 2] = dtEquipamentos.Valor("marca")?.ToString() ?? "";
                    worksheet.Cells[linhaAtual, 3] = dtEquipamentos.Valor("tipo")?.ToString() ?? "";
                    worksheet.Cells[linhaAtual, 4] = dtEquipamentos.Valor("serie")?.ToString() ?? "";

                    var valorcBConformidadeCE = dtEquipamentos.Valor("cBConformidadeCE")?.ToString();
                    string valorFinalConformidadeCE;

                    if (valorcBConformidadeCE == "NA")
                    {
                        valorFinalConformidadeCE = "NA";
                    }
                    else if (string.IsNullOrEmpty(valorcBConformidadeCE))
                    {
                        valorFinalConformidadeCE = "NC";
                    }
                    else if (valorcBConformidadeCE == "A")
                    {
                        valorFinalConformidadeCE = "C";
                    }
                    else
                    {
                        valorFinalConformidadeCE = "NC"; // padrão para outros valores
                    }

                    worksheet.Cells[linhaAtual, 5] = valorFinalConformidadeCE;


                    var valorcBDecreto_Lei = dtEquipamentos.Valor("cBDecreto_Lei")?.ToString();
                    string valorFinalDecretoLei;

                    if (valorcBDecreto_Lei == "NA")
                    {
                        valorFinalDecretoLei = "NA";
                    }
                    else if (string.IsNullOrEmpty(valorcBDecreto_Lei))
                    {
                        valorFinalDecretoLei = "NC";
                    }
                    else if (valorcBDecreto_Lei == "A")
                    {
                        valorFinalDecretoLei = "C";
                    }
                    else
                    {
                        valorFinalDecretoLei = "NC";
                    }

                    worksheet.Cells[linhaAtual, 6] = valorFinalDecretoLei;

                    var valorcBManutencao = dtEquipamentos.Valor("cBManutencao")?.ToString();
                    string valorFinalManutencao;

                    if (valorcBManutencao == "NA")
                    {
                        valorFinalManutencao = "NA";
                    }
                    else if (string.IsNullOrEmpty(valorcBManutencao))
                    {
                        valorFinalManutencao = "NC";
                    }
                    else if (valorcBManutencao == "A")
                    {
                        valorFinalManutencao = "C";
                    }
                    else
                    {
                        valorFinalManutencao = "NC";
                    }

                    worksheet.Cells[linhaAtual, 7] = valorFinalManutencao;

                    var valorcBManualInstrucoes = dtEquipamentos.Valor("cBManualInstrucoes")?.ToString();
                    string valorFinalManualInstrucoes;

                    if (valorcBManualInstrucoes == "NA")
                    {
                        valorFinalManualInstrucoes = "NA";
                    }
                    else if (string.IsNullOrEmpty(valorcBManualInstrucoes))
                    {
                        valorFinalManualInstrucoes = "NC";
                    }
                    else if (valorcBManualInstrucoes == "A")
                    {
                        valorFinalManualInstrucoes = "C";
                    }
                    else
                    {
                        valorFinalManualInstrucoes = "NC";
                    }

                    worksheet.Cells[linhaAtual, 8] = valorFinalManualInstrucoes;

                    var valorcBSeguro = dtEquipamentos.Valor("cBSeguro")?.ToString();
                    string valorFinalSeguro;

                    if (valorcBSeguro == "NA")
                    {
                        valorFinalSeguro = "NA";
                    }
                    else if (string.IsNullOrEmpty(valorcBSeguro))
                    {
                        valorFinalSeguro = "NC";
                    }
                    else if (valorcBSeguro == "A")
                    {
                        valorFinalSeguro = "C";
                    }
                    else
                    {
                        valorFinalSeguro = "NC";
                    }

                    worksheet.Cells[linhaAtual, 9] = valorFinalSeguro;

                    linhaAtual++;
                    numeroEquipamento++;
                    dtEquipamentos.Seguinte();
                }


                linhaAtual += 2;






                // Autofit das colunas
                worksheet.Columns.AutoFit();


                // Remover a folha em branco inicial
                Excel.Worksheet firstSheet = (Excel.Worksheet)workbook.Worksheets[1];
                if (workbook.Worksheets.Count > 1)
                {
                    firstSheet.Delete();
                }

            }
            catch (System.Exception ex)
            {

                MessageBox.Show("Erro ao criar ficheiro Excel: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Libertar recursos COM
                if (workbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
        }
        private void ExportarParaExcel2(List<string> idsSelecionados, string codigoObra)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add();
                var numidsempresa = idsSelecionados.Count;
                // Buscar dados reais da obra
                string queryObra = $@"SELECT COP.EntidadeIDA,GE.Nome,COP.CDU_LocalObra FROM COP_Obras AS COP
                                    INNER JOIN Geral_Entidade AS GE ON COP.EntidadeIDA = GE.EntidadeId
                                    WHERE COP.Codigo = '{codigoObra}'";
                var dadosObra = BSO.Consulta(queryObra);
                string descricaoObra = "", donoObra = "", entidadeExecutante = "";
                if (!dadosObra.Vazia())
                {
                    dadosObra.Inicio();
                    descricaoObra = dadosObra.Valor("CDU_LocalObra")?.ToString() ?? "";
                    donoObra = dadosObra.Valor("Nome")?.ToString() ?? "";
                    //entidadeExecutante = dadosObra.Valor("EntidadeExecutante")?.ToString() ?? "";
                }
                // Obter dados da empresa
                string idsFormatados = string.Join(",", idsSelecionados.Select(id => $"'{id}'"));

                string queryEmpresa = $"SELECT Nome FROM Geral_Entidade WHERE id IN ({idsFormatados})";

                StdBELista dtEmpresa = BSO.Consulta(queryEmpresa);
                string nomeEmpresa = "";

                dtEmpresa.Inicio();
                if (!dtEmpresa.NoFim())
                {
                    nomeEmpresa = dtEmpresa.Valor("Nome")?.ToString() ?? "";
                }

                // Criar nova folha para cada empresa
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.Add();
                // Limitar o nome da folha a 31 caracteres e remover caracteres inválidos
                string nomeEmpresaLimpo = nomeEmpresa.Replace("/", "").Replace("\\", "").Replace("?", "").Replace("*", "").Replace("[", "").Replace("]", "").Replace(":", "");
                string nomeFolha = $"Empresa";
                if (nomeFolha.Length > 31)
                {
                    nomeFolha = nomeFolha.Substring(0, 31);
                }
                worksheet.Name = nomeFolha;

                int linhaAtual = 1;

                // Adicionar cabeçalho da obra no topo da folha
                worksheet.Cells[linhaAtual, 1] = $"OBRA: {descricaoObra}";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Merge();
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Size = 14;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                linhaAtual += 2;

                worksheet.Cells[linhaAtual, 1] = $"DONO DE OBRA: {donoObra}";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Merge();
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Bold = true;
                linhaAtual += 2;

                worksheet.Cells[linhaAtual, 1] = "ENTIDADE EXECUTANTE: JOAQUIM PEIXOTO AZEVEDO & FILHOS LDA";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Merge();
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Bold = true;
                linhaAtual += 3;
                // Cabeçalhos principais
                worksheet.Cells[linhaAtual, 1] = "EMPRESA";
                worksheet.Cells[linhaAtual, 4] = "Alvará";
                worksheet.Cells[linhaAtual, 5] = "Contribuinte";
                worksheet.Cells[linhaAtual, 6] = "Não Div. Finanças";
                worksheet.Cells[linhaAtual, 7] = "Não Div. Seg. Social";
                worksheet.Cells[linhaAtual, 8] = "Folha Pag. Seg. Social";
                worksheet.Cells[linhaAtual, 9] = "Recibo de Pag. Seg. Social";
                worksheet.Cells[linhaAtual, 10] = "Apólice AT";
                worksheet.Cells[linhaAtual, 11] = "Recibo Apólice AT";
                worksheet.Cells[linhaAtual, 12] = "Apólice RC";
                worksheet.Cells[linhaAtual, 13] = "Recibo RC";
                worksheet.Cells[linhaAtual, 14] = "Horário de Trabalho";

                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Interior.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                linhaAtual++;

                // Sub-cabeçalhos
                worksheet.Cells[linhaAtual, 1] = "N.º";
                worksheet.Cells[linhaAtual, 2] = "Nome";
                worksheet.Cells[linhaAtual, 3] = "Sede";
                worksheet.Cells[linhaAtual, 4] = "N.º";
                worksheet.Cells[linhaAtual, 5] = "N.º";
                worksheet.Cells[linhaAtual, 6] = "Validade";
                worksheet.Cells[linhaAtual, 7] = "Validade";
                worksheet.Cells[linhaAtual, 8] = "Validade";
                worksheet.Cells[linhaAtual, 9] = "C; N/C; N/A";
                worksheet.Cells[linhaAtual, 10] = "C ; N/C ; N/A";
                worksheet.Cells[linhaAtual, 11] = "Validade";
                worksheet.Cells[linhaAtual, 12] = "C; N/C; N/A";
                worksheet.Cells[linhaAtual, 13] = "Validade";
                worksheet.Cells[linhaAtual, 14] = "C ; N/C ; N/A";

                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 14]].Interior.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                linhaAtual++;

                // Preencher empresas
                int numeroEmpresa = 1;

                foreach (string id in idsSelecionados)
                {
                    string query = $@"SELECT 
            Nome,AlvaraNumero, NIPC, Morada, CDU_ValidadeAlvara, CDU_ValidadeFinancas, CDU_ValidadeSegSocial, CDU_ValidadeFolhaPag,
            CDU_ValidadeComprovativoPagamento, CDU_ValidadeReciboSeguroAT, CDU_ValidadeSeguroRC,
            CDU_ValidadeHorarioTrabalho, CDU_ValidadeSeguroAT
            FROM Geral_Entidade WHERE id = '{id}'";

                    StdBELista empresa = BSO.Consulta(query);
                    empresa.Inicio();



                    if (!empresa.NoFim())
                    {
                        string nome = empresa.Valor("Nome")?.ToString() ?? "";
                        string alvara = empresa.Valor("AlvaraNumero")?.ToString() ?? "";
                        string nif = empresa.Valor("NIPC")?.ToString() ?? "";
                        string morada = empresa.Valor("Morada")?.ToString() ?? "";
                        if (id == "2A8C7ECD-309B-49F9-A337-203B45CED948")
                        {

                            worksheet.Cells[linhaAtual, 1] = numeroEmpresa;
                            worksheet.Cells[linhaAtual, 2] = nome;
                            worksheet.Cells[linhaAtual, 3] = morada;
                            worksheet.Cells[linhaAtual, 4] = alvara;
                            worksheet.Cells[linhaAtual, 5] = nif;
                            worksheet.Cells[linhaAtual, 6] = "C";
                            worksheet.Cells[linhaAtual, 7] = "C";
                            worksheet.Cells[linhaAtual, 8] = "C";
                            worksheet.Cells[linhaAtual, 9] = "C";
                            worksheet.Cells[linhaAtual, 10] = "C";
                            worksheet.Cells[linhaAtual, 11] = "C";
                            worksheet.Cells[linhaAtual, 12] = "C";
                            worksheet.Cells[linhaAtual, 13] = "C";
                            worksheet.Cells[linhaAtual, 14] = "C";


                            linhaAtual++;
                            numeroEmpresa++;
                            continue; // pula para o próximo ID
                        }






                        DateTime.TryParse(empresa.Valor("CDU_ValidadeAlvara")?.ToString(), out DateTime validadeAlvara);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeFinancas")?.ToString(), out DateTime validadeFinancas);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeSegSocial")?.ToString(), out DateTime validadeSegSocial);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeFolhaPag")?.ToString(), out DateTime validadeFolhaPag);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeComprovativoPagamento")?.ToString(), out DateTime validadeComprovativoPagamento);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeReciboSeguroAT")?.ToString(), out DateTime validadeReciboSeguroAT);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeSeguroRC")?.ToString(), out DateTime validadeSeguroRC);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeHorarioTrabalho")?.ToString(), out DateTime validadeHorarioTrabalho);
                        DateTime.TryParse(empresa.Valor("CDU_ValidadeSeguroAT")?.ToString(), out DateTime validadeSeguroAT);

                        worksheet.Cells[linhaAtual, 1] = numeroEmpresa;
                        worksheet.Cells[linhaAtual, 2] = nome;
                        worksheet.Cells[linhaAtual, 3] = morada;
                        worksheet.Cells[linhaAtual, 4] = alvara;
                        worksheet.Cells[linhaAtual, 5] = nif;
                        worksheet.Cells[linhaAtual, 6] = validadeAlvara.Year > 1 ? validadeAlvara.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 7] = validadeFinancas.Year > 1 ? validadeFinancas.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 8] = validadeSegSocial.Year > 1 ? validadeSegSocial.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 9] = validadeFolhaPag.Year > 1 ? validadeFolhaPag.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 10] = validadeComprovativoPagamento.Year > 1 ? "C" : "NC";
                        worksheet.Cells[linhaAtual, 11] = validadeSeguroAT.Year > 1 ? validadeSeguroAT.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 12] = validadeReciboSeguroAT.Year > 1 ? "C" : "NC";
                        worksheet.Cells[linhaAtual, 13] = validadeSeguroRC.Year > 1 ? validadeSeguroRC.ToString("dd/MM/yyyy") : "NC";
                        worksheet.Cells[linhaAtual, 14] = validadeHorarioTrabalho.Year > 1 ? "C" : "NC";

                        linhaAtual++;
                        numeroEmpresa++;
                    }
                }

                linhaAtual += 2;



                // Dados dos Trabalhadores
                worksheet.Cells[linhaAtual, 1] = "TRABALHADORES";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Merge();
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                linhaAtual++;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                // Cabeçalhos dos trabalhadores
                worksheet.Cells[linhaAtual, 1] = "N.º"; // NOVA COLUNA
                worksheet.Cells[linhaAtual, 2] = "Nome Completo";
                worksheet.Cells[linhaAtual, 3] = "Empresa";
                worksheet.Cells[linhaAtual, 4] = "Categoria";
                worksheet.Cells[linhaAtual, 5] = "Contribuinte";
                worksheet.Cells[linhaAtual, 6] = "Nº Segurança Social";
                worksheet.Cells[linhaAtual, 7] = "Cartão de cidadão ou residencia";
                worksheet.Cells[linhaAtual, 8] = "Ficha Médica de aptidão";
                worksheet.Cells[linhaAtual, 9] = "Credenciação do trabalhador";
                worksheet.Cells[linhaAtual, 10] = "Trabalhos especializados";
                worksheet.Cells[linhaAtual, 11] = "Ficha de distribuição de EPI's";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 11]].Font.Bold = true;
                linhaAtual++;


                // Dados dos trabalhadores



                int numeroTrabalhador = 1;




                var queryTrabalhadoresJPA = $@"SELECT COP.codigo ,COP_P.Funcionario,C.Descricao,F.NumContr,F.NumBeneficiario  FROM COP_Obras AS COP
   INNER JOIN COP_Obras_Pessoal AS COP_P ON COP.id = COP_P.ObraID 
   INNER JOIN GPR_Operadores AS O ON COP_P.ColaboradorID = O.IDOperador
   INNER JOIN Funcionarios AS F ON O.Funcionario = F.Codigo
   INNER JOIN Categorias AS C ON F.Categoria = C.Categoria
   WHERe COP.Codigo = '{codigoObra}'
";
                StdBELista dtTrabalhadoresJPA = BSO.Consulta(queryTrabalhadoresJPA);
                dtTrabalhadoresJPA.Inicio();
                while (!dtTrabalhadoresJPA.NoFim())
                {
                    string funcionario = dtTrabalhadoresJPA.Valor("Funcionario")?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(funcionario))
                    {
                        string categoria = dtTrabalhadoresJPA.Valor("Descricao")?.ToString() ?? "N/A";
                        string numeroContr = dtTrabalhadoresJPA.Valor("NumContr")?.ToString() ?? "N/A";

                        string numeroBeneficiario = dtTrabalhadoresJPA.Valor("NumBeneficiario")?.ToString() ?? "N/A";
                        worksheet.Cells[linhaAtual, 1] = numeroTrabalhador; // N.º
                        worksheet.Cells[linhaAtual, 2] = funcionario;
                        worksheet.Cells[linhaAtual, 3] = "JPA"; // Empresa (pode preencher se quiser)
                        worksheet.Cells[linhaAtual, 4] = categoria;
                        worksheet.Cells[linhaAtual, 5] = numeroContr;
                        worksheet.Cells[linhaAtual, 6] = numeroBeneficiario;
                        worksheet.Cells[linhaAtual, 7] = "C";
                        worksheet.Cells[linhaAtual, 8] = "C";
                        worksheet.Cells[linhaAtual, 9] = "C";
                        worksheet.Cells[linhaAtual, 10] = "C";
                        worksheet.Cells[linhaAtual, 11] = "C";

                        linhaAtual++;
                        numeroTrabalhador++;
                    }
                    dtTrabalhadoresJPA.Seguinte();
                }





                linhaAtual += 1;

                // Dados dos Equipamentos
                worksheet.Cells[linhaAtual, 1] = "EQUIPAMENTOS";
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Merge(); // Atualizado para 9 colunas
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                linhaAtual++;


                // Cabeçalhos com N.º
                worksheet.Cells[linhaAtual, 1] = "N.º";
                worksheet.Cells[linhaAtual, 2] = "Marca";
                worksheet.Cells[linhaAtual, 3] = "Tipo";
                worksheet.Cells[linhaAtual, 4] = "Série";
                worksheet.Cells[linhaAtual, 5] = "Anexo 1";
                worksheet.Cells[linhaAtual, 6] = "Anexo 2";
                worksheet.Cells[linhaAtual, 7] = "Anexo 3";
                worksheet.Cells[linhaAtual, 8] = "Anexo 4";
                worksheet.Cells[linhaAtual, 9] = "Anexo 5";

                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[linhaAtual, 1], worksheet.Cells[linhaAtual, 9]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                linhaAtual++;


                // Consulta aos equipamentos
                string queryEquipamentos = $@"
    SELECT marca, tipo, serie, anexo1, anexo2, anexo3, anexo4, anexo5 
    FROM TDU_AD_Equipamentos 
    WHERE id_empresa IN  ({idsFormatados})";

                StdBELista dtEquipamentos = BSO.Consulta(queryEquipamentos);

                int numeroEquipamento = 1;

                dtEquipamentos.Inicio();
                while (!dtEquipamentos.NoFim())
                {
                    worksheet.Cells[linhaAtual, 1] = numeroEquipamento; // N.º
                    worksheet.Cells[linhaAtual, 2] = dtEquipamentos.Valor("marca")?.ToString() ?? "";
                    worksheet.Cells[linhaAtual, 3] = dtEquipamentos.Valor("tipo")?.ToString() ?? "";
                    worksheet.Cells[linhaAtual, 4] = dtEquipamentos.Valor("serie")?.ToString() ?? "";

                    worksheet.Cells[linhaAtual, 5] = dtEquipamentos.Valor("anexo1")?.ToString() == "True" ? "C" : "NC";
                    worksheet.Cells[linhaAtual, 6] = dtEquipamentos.Valor("anexo2")?.ToString() == "True" ? "C" : "NC";
                    worksheet.Cells[linhaAtual, 7] = dtEquipamentos.Valor("anexo3")?.ToString() == "True" ? "C" : "NC";
                    worksheet.Cells[linhaAtual, 8] = dtEquipamentos.Valor("anexo4")?.ToString() == "True" ? "C" : "NC";
                    worksheet.Cells[linhaAtual, 9] = dtEquipamentos.Valor("anexo5")?.ToString() == "True" ? "C" : "NC";

                    linhaAtual++;
                    numeroEquipamento++;
                    dtEquipamentos.Seguinte();
                }


                linhaAtual += 2;






                // Autofit das colunas
                worksheet.Columns.AutoFit();


                // Remover a folha em branco inicial
                Excel.Worksheet firstSheet = (Excel.Worksheet)workbook.Worksheets[1];
                if (workbook.Worksheets.Count > 1)
                {
                    firstSheet.Delete();
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao criar ficheiro Excel: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Libertar recursos COM
                if (workbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
        }
        private void ExportarParaExcelNovo(System.Collections.Generic.List<string> idsSelecionados, string codigoObra)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application { Visible = true, DisplayAlerts = false };
                workbook = excelApp.Workbooks.Add();
                ws = (Excel.Worksheet)workbook.Worksheets[1];
                ws.Name = "Resumo Empr";

                // Helpers
                int ToOle(System.Drawing.Color c) => ColorTranslator.ToOle(c);
                void Borda(Excel.Range r) => r.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                void Negrito(Excel.Range r, bool v = true) => r.Font.Bold = v;
                void Centro(Excel.Range r) { r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; r.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; }
                void Wrap(Excel.Range r, bool v = true) => r.WrapText = v;
                Excel.Range R(int l1, int c1, int l2, int c2) => ws.Range[ws.Cells[l1, c1], ws.Cells[l2, c2]];

                // Linha 1: Marca + Título
                ws.Cells[1, 1] = "JPA";
                var rLogo = R(1, 1, 1, 1); Negrito(rLogo); rLogo.Font.Size = 14;

                ws.Cells[1, 4] = "Controlo de Documentos de Empresas, Trabalhadores e Máquinas/Equipamentos";
                var rTitulo = R(1, 4, 1, 23); rTitulo.Merge(); Negrito(rTitulo); Centro(rTitulo);
       
                codigoObra = f4TabelaSQL1.Text;
                var querydadosObra = $@"
SELECT * frOM COP_Obras WHERE Codigo = '{codigoObra}'
                    ";



                var dadosObra = BSO.Consulta(querydadosObra);
                dadosObra.NumLinhas();
                string entidadeIdA = "";
                if (dadosObra.NumLinhas() == 0)
                {
                    return;
                }
                else
                {
                     entidadeIdA = dadosObra.DaValor<string>("EntidadeIDA");
                }


                var dadosDonoObra = default(StdBELista);
                int ln = 3;
                string donoObra = "";
                if (!string.IsNullOrEmpty(entidadeIdA))
                {
                    var queryDonoObra = $"SELECT * FROM Geral_Entidade WHERE EntidadeId = '{entidadeIdA}'";
                     dadosDonoObra = BSO.Consulta(queryDonoObra);
                    R(ln, 2, ln, 4).Merge(); ws.Cells[ln, 2] = $"Designação da Empreitada: {dadosObra.DaValor<string>("Codigo")}"; Borda(R(ln, 2, ln, 4)); ln++;
                    R(ln, 2, ln, 4).Merge(); ws.Cells[ln, 2] = $"Dono de Obra: {dadosDonoObra.DaValor<string>("Nome")}"; Borda(R(ln, 2, ln, 4)); ln++;
                    R(ln, 2, ln, 4).Merge(); ws.Cells[ln, 2] = $"Entidade Executante: {dadosDonoObra.DaValor<string>("Nome")}"; Borda(R(ln, 2, ln, 4));
                    donoObra = dadosDonoObra.DaValor<string>("Nome");
                }
                else
                {
                    R(ln, 2, ln, 4).Merge(); ws.Cells[ln, 2] = $"Designação da Empreitada: "; Borda(R(ln, 2, ln, 4)); ln++;
                    R(ln, 2, ln, 4).Merge(); ws.Cells[ln, 2] = $"Dono de Obra: "; Borda(R(ln, 2, ln, 4)); ln++;
                    R(ln, 2, ln, 4).Merge(); ws.Cells[ln, 2] = $"Entidade Executante: "; Borda(R(ln, 2, ln, 4));
                    donoObra = "";
                }
                // Blocos topo (esquerda)
              
    

                // Elaborado por
                int lnQas = 3;
                R(lnQas, 6, lnQas, 8).Merge(); ws.Cells[lnQas, 6] = "Elaborado por (Téc. QAS):"; Centro(R(lnQas, 6, lnQas, 8)); Borda(R(lnQas, 6, lnQas, 8)); lnQas++;
                ws.Cells[lnQas, 6] = "Nome:"; Borda(R(lnQas, 6, lnQas, 6)); R(lnQas, 7, lnQas, 8).Merge(); Borda(R(lnQas, 7, lnQas, 8)); lnQas++;
                ws.Cells[lnQas, 6] = "Assinatura:"; Borda(R(lnQas, 6, lnQas, 6)); R(lnQas, 7, lnQas, 8).Merge(); Borda(R(lnQas, 7, lnQas, 8)); lnQas++;
                ws.Cells[lnQas, 6] = "Data:"; Borda(R(lnQas, 6, lnQas, 6)); R(lnQas, 7, lnQas, 8).Merge(); ws.Cells[lnQas, 7] = "14 de outubro de 2025"; Borda(R(lnQas, 7, lnQas, 8));

                // Verificado por
                int lnVer = 3;
                R(lnVer, 10, lnVer, 12).Merge(); ws.Cells[lnVer, 10] = "Verificado por (Dir. Técnico / Dir. Obra):"; Centro(R(lnVer, 10, lnVer, 12)); Borda(R(lnVer, 10, lnVer, 12)); lnVer++;
                ws.Cells[lnVer, 10] = "Nome:"; Borda(R(lnVer, 10, lnVer, 10)); R(lnVer, 11, lnVer, 12).Merge(); Borda(R(lnVer, 11, lnVer, 12)); lnVer++;
                ws.Cells[lnVer, 10] = "Assinatura:"; Borda(R(lnVer, 10, lnVer, 10)); R(lnVer, 11, lnVer, 12).Merge(); Borda(R(lnVer, 11, lnVer, 12)); lnVer++;
                ws.Cells[lnVer, 10] = "Data:"; Borda(R(lnVer, 10, lnVer, 10)); R(lnVer, 11, lnVer, 12).Merge(); ws.Cells[lnVer, 11] = "14 de outubro de 2025"; Borda(R(lnVer, 11, lnVer, 12));

                // Datas à direita
                R(7, 16, 7, 16).Value2 = "14/10/2025";
                R(7, 17, 7, 17).Value2 = "15/10/2025";
                Centro(R(7, 16, 7, 17)); Negrito(R(7, 16, 7, 17));

                // Contactos (2 linhas)
                int lnCont = 8;
                ws.Cells[lnCont, 4] = "Pessoa(s) p/ contacto:"; ws.Cells[lnCont, 6] = "Telf:"; ws.Cells[lnCont, 8] = "E-mail:"; ws.Cells[lnCont, 10] = "Função:"; lnCont++;
                ws.Cells[lnCont, 6] = "Telf:"; ws.Cells[lnCont, 8] = "E-mail:"; ws.Cells[lnCont, 10] = "Função:";

                // Nota legal
                lnCont++;
                ws.Cells[lnCont, 1] = "Nota: Documento ao abrigo do Artigo 21. do D.L. 273/2003, de 29/10";
                lnCont += 2;

                // ===== TABELA EMPRESA =====
                int startTableRow = lnCont;

                // EMPRESA (col 1–4)
                R(startTableRow, 1, startTableRow, 4).Merge();
                ws.Cells[startTableRow, 1] = "EMPRESA";
                var rEmp = R(startTableRow, 1, startTableRow, 4);
                Negrito(rEmp); Centro(rEmp); Borda(rEmp); rEmp.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                rEmp.RowHeight = 22;

                // Numeração 1..23 (col 5..23)
                int lastCol = 27;
                int numRow = startTableRow;

                for (int c = 5, n = 1; c <= lastCol && n <= 23; n++)
                {
                    Microsoft.Office.Interop.Excel.Range r;

                    // Se o número precisa ocupar 2 colunas (2 ou 9)
                    if (n == 2 || n == 9)
                    {
                        r = ws.Range[ws.Cells[numRow, c], ws.Cells[numRow, c + 1]];
                        r.Merge();
                        ws.Cells[numRow, c] = n.ToString();
                        c += 2; // avança duas colunas
                    }
                    else
                    {
                        // Célula normal
                        r = ws.Range[ws.Cells[numRow, c], ws.Cells[numRow, c]];
                        ws.Cells[numRow, c] = n.ToString();
                        c++; // avança uma coluna
                    }

                    // Estilos
                    Negrito(r);
                    Centro(r);
                    Borda(r);
                    r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                }



                // Cabeçalhos principais + sub-cabeçalhos
                int headerRow = startTableRow + 1;
                int subHeaderRow = startTableRow + 2;

                // Colunas 1–4 ocupam as duas linhas
                string[] col1a4 = { "N.º", "Designação Social", "Sede", "Atividade desenvolvida em Obra" };
                for (int i = 0; i < col1a4.Length; i++)
                {
                    var rr = R(headerRow, i + 1, subHeaderRow, i + 1);
                    rr.Merge(); ws.Cells[headerRow, i + 1] = col1a4[i];
                    Negrito(rr); Centro(rr); Wrap(rr); Borda(rr); rr.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                }

                // Mapa de cabeçalhos (5..23)
                var headers = new (int colStart, int colEnd, string title)[]
                {
            (5, 5,  "Contribuinte"),
            (6, 7,  "Alvará / Certificado"),
            (8, 8,  "Anexo D"),
            (9, 9,  "Cert. ND Finanças"),
            (10,10, "Decl. ND Seg. Social"),
            (11,11, "Folha Pag. Seg. Social"),
            (12,12, "Recibo de Pag. Seg. Social"),
            (13,13, "Apólice AT"),
            (14,15, "Modalidade do Seguro"),
            (16,16, "Recibo Apólice AT"),
            (17,17, "Apólice RC"),
            (18,18, "Recibo RC"),
            (19,19, "Registo(s) Criminal(ais)"),
            (20,20, "Horário de Trabalho"),
            (21,21, "Dec.  Trab. Imigr."),
            (22,22, "Dec. Resp. Estaleiro"),
            (23,23, "Dec. Ades. PSS"),
            (24,24, "Contrato Subempreitada"),
            (25,25, "Entrada em Obra"),
            (26,26, "Saída de Obra"),
            (27,27, "Autorização de Entrada"),
                    // Se quiseres ainda “Contrato Subempreitada”, “Entrada/Saída/Autoriz.” empurra estas colunas para a frente
                };

                foreach (var h in headers)
                {
                    var r = R(headerRow, h.colStart, headerRow, h.colEnd);
                    r.Merge(); ws.Cells[headerRow, h.colStart] = h.title;
                    Negrito(r); Centro(r); Wrap(r); Borda(r); r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                }

                // Sub-cabeçalhos (5..23) — exactamente como no modelo
                var subs = new (int col, string text)[]
                {
            (5,"N.º"),
            (6,"N.º"),
            (7,"PUB / PRIV"),
            (8,"C ; N/C ; N/A"),
            (9,"Validade"),
            (10,"Validade"),
            (11,"Validade"),
            (12,"C ; N/C ; N/A"),
            (13,"N.º"),
            (14,"Fixo?"),
            (15,"Prémio Variável?"),
            (16,"Validade"),
            (17,"N.º"),
            (18,"Validade"),
            (19,"C ; N/C ; N/A"),
            (20,"C ; N/C ; N/A"),
            (21,"C ; N/C ; N/A"),
            (22,"C ; N/C ; N/A"),
            (23,"C ; N/C ; N/A"),
            (24,"Com:"),
            (25,"Data"),
            (26,"Data"),
            (27,"Sim / Não"),
                    // Se estenderes com mais colunas:
                    // (24,"Com:"), (25,"Data"), (26,"Data"), (27,"Sim / Não")
                };

                foreach (var s in subs)
                {
                    var r = R(subHeaderRow, s.col, subHeaderRow, s.col);
                    ws.Cells[subHeaderRow, s.col] = s.text;
                    Negrito(r); Centro(r); Wrap(r); Borda(r); r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                }

                // Primeira linha de dados
                int firstDataRow = subHeaderRow + 1;

                // ===== Linha de exemplo com dados fictícios =====
                int row = firstDataRow;

                //mostrar os ids no message box
              
                

                var numIds = idsSelecionados.Count;

                for (int i = 0; i < numIds; i++)
                {
                 
                    var queryDadosEmpresa = $@"SELECT * FROM Geral_Entidade AS GE
                                            INNER JOIN TDU_AD_Autorizacoes AS A ON GE.ID = A.ID_Entidade
                                            WHEre GE.ID = '{idsSelecionados[i]}'";

                    var dadosEmpresa = BSO.Consulta(queryDadosEmpresa);
                    var numlinhas = dadosEmpresa.NumLinhas();

                    if (numlinhas < 1)
                    {
                        ws.Cells[row, 1] = i + 1; // N.º
                        ws.Cells[row, 2] = "JOAQUIM PEIXOTO AZEVEDO & FILHOS, LDA"; // Designação Social
                        ws.Cells[row, 3] = "RUA DE LONGRAS Nº 44"; // Sede
                        ws.Cells[row, 4] = "Empreiteiro Geral";// Atividade desenvolvida em Obra
                        ws.Cells[row, 5] = "502244585"; // Contribuinte
                        ws.Cells[row, 6] = ""; // N.º Alvará
                        ws.Cells[row, 7] = "PAR"; // PUB/PRIV
                        ws.Cells[row, 8] = "C";
                        ws.Cells[row, 9] = ""; // Cert. ND Finanças
                        ws.Cells[row, 10] = ""; // Decl. ND Seg. Social
                        ws.Cells[row, 11] = "";// Folha Pag. Seg. Social
                        ws.Cells[row, 12] = "C"; // sem valor
                        ws.Cells[row, 13] = ""; // Apólice AT
                        ws.Cells[row, 14] = ""; // Fixo?
                        ws.Cells[row, 15] = ""; // Prémio Variável?
                        ws.Cells[row, 16] = ""; // Recibo Apólice AT
                        ws.Cells[row, 17] = ""; // Apólice RC
                        ws.Cells[row, 18] = ""; 
                        ws.Cells[row, 19] = "C";
                        ws.Cells[row, 20] = "C";
                        ws.Cells[row, 21] = "C";
                        ws.Cells[row, 22] = "C";
                        ws.Cells[row, 23] = "C";
                        ws.Cells[row, 24] = "";
                        ws.Cells[row, 27] = "Sim";

                        // 🔹 Cria um range que abrange todas as células dessa linha
                        Excel.Range linhaRange = ws.Range[ws.Cells[row, 1], ws.Cells[row, 27]];

                        // 🔹 Aplica centralização e borda
                        Centro(linhaRange);
                        Borda(linhaRange);

                        row++;
                        continue;
                    }

                    // Colunas 1–4: dados da empresa
                    ws.Cells[row, 1] = i + 1; // N.º
                    ws.Cells[row, 2] = dadosEmpresa.DaValor<string>("Nome"); // Designação Social
                    ws.Cells[row, 3] = "Lisboa"; // Sede
                    ws.Cells[row, 4] = "Trabalhos de construção civil"; // Atividade desenvolvida em Obra

                    // Colunas 5–27: documentos e campos
                    ws.Cells[row, 5] = dadosEmpresa.DaValor<string>("NIPC"); // Contribuinte
                    ws.Cells[row, 6] = dadosEmpresa.DaValor<string>("AlvaraNumero"); // N.º Alvará
                    ws.Cells[row, 7] = "PÚB";       // PUB/PRIV

                    var valor = dadosEmpresa.DaValor<string>("CDU_AnexoAnexoD");
                    ws.Cells[row, 8] = !string.IsNullOrEmpty(valor) ? "C" : "N/C";
                    // Cert. ND Finanças
                    if (DateTime.TryParse(dadosEmpresa.DaValor<string>("CDU_ValidadeFinancas"), out DateTime data))
                        ws.Cells[row, 9] = data.ToString("dd/MM/yyyy");
                    else
                        ws.Cells[row, 9] = "";

                    // Decl. ND Seg. Social
                    if (DateTime.TryParse(dadosEmpresa.DaValor<string>("CDU_ValidadeSegSocial"), out data))
                        ws.Cells[row, 10] = data.ToString("dd/MM/yyyy");
                    else
                        ws.Cells[row, 10] = "";

                    // Folha Pag. Seg. Social
                    if (DateTime.TryParse(dadosEmpresa.DaValor<string>("CDU_ValidadeFolhaPag"), out data))
                        ws.Cells[row, 11] = data.ToString("dd/MM/yyyy");
                    else
                        ws.Cells[row, 11] = "";



                    string valFolha = dadosEmpresa.DaValor<string>("CDU_ValidadeComprovativoPagamento");
                    DateTime validade;

                    if (string.IsNullOrWhiteSpace(valFolha))
                    {
                        ws.Cells[row, 12] = "N/A"; // sem valor
                    }
                    else if (DateTime.TryParse(valFolha, out validade))
                    {
                        ws.Cells[row, 12] = validade < DateTime.Today ? "N/C" : "C";
                    }
                    else
                    {
                        // se a string existir mas não for uma data válida
                        ws.Cells[row, 12] = "N/A";
                    } // Recibo Pag. Seg. Social

                    ws.Cells[row, 13] = "";         // Apólice AT
                    ws.Cells[row, 14] = "";       // Fixo?
                    ws.Cells[row, 15] = "";       // Prémio Variável?
                  
                    if (DateTime.TryParse(dadosEmpresa.DaValor<string>("CDU_ValidadeReciboSeguroAT"), out data))
                        ws.Cells[row, 16] = data.ToString("dd/MM/yyyy");
                    else
                        ws.Cells[row, 16] = "";

                    ws.Cells[row, 17] = "";  // Apólice RC
                    if (DateTime.TryParse(dadosEmpresa.DaValor<string>("CDU_ValidadeSeguroRC"), out data))
                        ws.Cells[row, 18] = data.ToString("dd/MM/yyyy");
                    else
                        ws.Cells[row, 18] = "";
                    ws.Cells[row, 19] = "C";         // Registo(s) Criminal(ais) 

                    string valor01 = dadosEmpresa.DaValor<string>("caminho2");
                    string resultado = VerificaValidade(valor01);
                    ws.Cells[row, 20] = resultado;

                    string valor2 = dadosEmpresa.DaValor<string>("caminho5");
                    string resultado2 = VerificaValidade(valor2);
                    ws.Cells[row, 21] = resultado2;       // Dec. Trab. Imigr.

                    string valor3 = dadosEmpresa.DaValor<string>("caminho4");
                    string resultado3 = VerificaValidade(valor3);
                    ws.Cells[row, 22] = resultado3;         // Dec. Resp. Estaleiro

                    string valor4 = dadosEmpresa.DaValor<string>("caminho3");
                    string resultado4 = VerificaValidade(valor4);
                    ws.Cells[row, 23] = resultado4;         // Dec. Ades. PSS

                    ws.Cells[row, 24] = ""; // Contrato Subempreitada (Com)


                    string dataEntradaStr = dadosEmpresa.DaValor<string>("Data_Entrada");
                    string dataSaidaStr = dadosEmpresa.DaValor<string>("Data_Saida");

                    // Tenta converter as strings em DateTime
                    DateTime dataEntrada, dataSaida;
                    bool temEntrada = DateTime.TryParse(dataEntradaStr, out dataEntrada);
                    bool temSaida = DateTime.TryParse(dataSaidaStr, out dataSaida);

                    // Se a data de saída for 01/01/1900, considera como vazia
                    if (temSaida &&
                    (dataSaida == new DateTime(1900, 1, 1) || dataSaida == new DateTime(1753, 1, 1)))
                    {
                        temSaida = false;
                    }

                    string autorizacao;

                    if (!temEntrada)
                    {
                        // Sem data de entrada = sem autorização
                        autorizacao = "Não";
                    }
                    else if (!temSaida)
                    {
                        // Tem entrada e sem saída (ou data 1900) = sempre autorizado
                        autorizacao = "Sim";
                    }
                    else
                    {
                        // Tem ambas: verifica se hoje está entre as datas (inclusive)
                        if (DateTime.Today >= dataEntrada.Date && DateTime.Today <= dataSaida.Date)
                            autorizacao = "Sim";
                        else
                            autorizacao = "Não";
                    }

                    // Grava na célula
                    ws.Cells[row, 27] = "Sim";

                    // Se quiser gravar apenas a parte da data (sem horas) nas colunas 25 e 26:
                    ws.Cells[row, 25] = temEntrada ? dataEntrada.ToString("dd/MM/yyyy") : "";
                    ws.Cells[row, 26] = temSaida ? dataSaida.ToString("dd/MM/yyyy") : "";

                    // Formatação básica da linha
                    var linhaExemplo = R(row, 1, row, 27);
                    Centro(linhaExemplo);
                    Borda(linhaExemplo);
                    linhaExemplo.RowHeight = 22;
                    row++;
                    // ===== Fim da linha de exemplo =====
                }



                // Larguras (1..23) — ajusta se precisares
                double[] widths =
                {
            6, 28, 22, 28,   // 1..4
            10, 10, 11, 12,  // 5..8
            12, 12, 14, 16,  // 9..12
            12, 8, 16, 12,   // 13..16
            12, 12, 14, 14,  // 17..20
            14, 14, 14       // 21..23
        };
                for (int c = 1; c <= widths.Length; c++)
                    ((Excel.Range)ws.Columns[c]).ColumnWidth = widths[c - 1];

                // Alturas cabeçalhos
                ws.Rows[headerRow].RowHeight = 28;
                ws.Rows[subHeaderRow].RowHeight = 30;

                // Congelar painéis
                ws.Activate();
                excelApp.ActiveWindow.SplitRow = subHeaderRow;
                excelApp.ActiveWindow.FreezePanes = true;

                // Page setup
                ws.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                ws.PageSetup.LeftMargin = excelApp.InchesToPoints(0.4);
                ws.PageSetup.RightMargin = excelApp.InchesToPoints(0.3);
                ws.PageSetup.TopMargin = excelApp.InchesToPoints(0.5);
                ws.PageSetup.BottomMargin = excelApp.InchesToPoints(0.5);
                ws.PageSetup.Zoom = false;
                ws.PageSetup.FitToPagesWide = 1;
                ws.PageSetup.FitToPagesTall = false;



                // === Criar nova folha ===

                Criar2Pagina(workbook, excelApp, codigoObra, idsSelecionados, dadosObra, dadosDonoObra, donoObra);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao criar ficheiro Excel: " + ex.Message, "Erro",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private void Criar2Pagina(Excel.Workbook workbook, Excel.Application excelApp, string codigoObra, List<string> idsSelecionados, StdBELista dadosObra, StdBELista dadosDonoObra, string DonoObra)
        {
            Excel.Worksheet ws2 = null;

            try
            {
                ws2 = (Excel.Worksheet)workbook.Worksheets.Add();
                ws2.Name = "1 - JPA";

                // Helpers
                int ToOle(System.Drawing.Color c) => ColorTranslator.ToOle(c);
                void Borda(Excel.Range r) => r.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                void Negrito(Excel.Range r, bool v = true) => r.Font.Bold = v;
                void Centro(Excel.Range r) { r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; r.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; }
                void Wrap(Excel.Range r, bool v = true) => r.WrapText = v;
                Excel.Range R(int l1, int c1, int l2, int c2) => ws2.Range[ws2.Cells[l1, c1], ws2.Cells[l2, c2]];

                // Linha 1: Marca + Título
                ws2.Cells[1, 1] = "JPA";
                var rLogo = R(1, 1, 1, 1); Negrito(rLogo); rLogo.Font.Size = 14;

                ws2.Cells[1, 4] = "Controlo de Documentos de Empresas, Trabalhadores e Máquinas/Equipamentos";
                var rTitulo = R(1, 4, 1, 20); rTitulo.Merge(); Negrito(rTitulo); Centro(rTitulo);

                ws2.Cells[3, 3] = "Identificação dos Intervenientes";
                var rIdentificacao = ws2.Range[ws2.Cells[3, 3], ws2.Cells[3, 8]];
                rIdentificacao.Merge();
                rIdentificacao.Font.Bold = true;
                rIdentificacao.Font.Size = 12;
                rIdentificacao.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rIdentificacao.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                // Fundo cinzento
                rIdentificacao.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                // Bordas
                rIdentificacao.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rIdentificacao.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Linha 4
                ws2.Cells[4, 3] = "Designação da Empreitada:";
                var rDesignacao = ws2.Range[ws2.Cells[4, 4], ws2.Cells[4, 8]];
                rDesignacao.Merge();
                rDesignacao.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rDesignacao.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                ws2.Cells[4, 4] = codigoObra;

                // Linha 5
                ws2.Cells[5, 3] = "Dono de Obra:";
                var rDonoObra = ws2.Range[ws2.Cells[5, 4], ws2.Cells[5, 8]];
                rDonoObra.Merge();
                rDonoObra.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rDonoObra.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                ws2.Cells[5, 4] = DonoObra;

                // Linha 6
                ws2.Cells[6, 3] = "Entidade Executante:";
                var rEntidade = ws2.Range[ws2.Cells[6, 4], ws2.Cells[6, 8]];
                rEntidade.Merge();
                rEntidade.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rEntidade.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                ws2.Cells[6, 4] = DonoObra;

                ws2.Range[ws2.Cells[4, 3], ws2.Cells[6, 3]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                ws2.Range[ws2.Cells[4, 3], ws2.Cells[6, 3]].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                //---------------------
                ws2.Cells[3, 12] = "Elaborado por (Téc. QAS):";
                var rElaborado = ws2.Range[ws2.Cells[3, 12], ws2.Cells[3, 17]];
                rElaborado.Merge();
                rElaborado.Font.Bold = true;
                rElaborado.Font.Size = 12;
                rElaborado.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rElaborado.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                // Fundo cinzento
                rElaborado.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                // Bordas
                rElaborado.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rElaborado.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Linha 4
                ws2.Cells[4, 12] = "Nome:";
                var rNome = ws2.Range[ws2.Cells[4, 13], ws2.Cells[4, 17]];
                rNome.Merge();
                rNome.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rNome.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Linha 5
                ws2.Cells[5, 12] = "Assinatura:";
                var rAssinaura = ws2.Range[ws2.Cells[5, 13], ws2.Cells[5, 17]];
                rAssinaura.Merge();
                rAssinaura.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rAssinaura.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Linha 6
                ws2.Cells[6, 12] = "Data:";
                var rData1 = ws2.Range[ws2.Cells[6, 13], ws2.Cells[6, 17]];
                rData1.Merge();
                rData1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rData1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                ws2.Cells[6, 13] = DateTime.Now.ToString("d 'de' MMMM 'de' yyyy", new System.Globalization.CultureInfo("pt-PT"));


                ws2.Range[ws2.Cells[4, 12], ws2.Cells[6, 12]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                ws2.Range[ws2.Cells[4, 12], ws2.Cells[6, 12]].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                //---------------------------------
                ws2.Cells[3, 20] = "Verificado por (Dir. Técnico / Dir. Obra):";
                var rVerificado = ws2.Range[ws2.Cells[3, 20], ws2.Cells[3, 23]];
                rVerificado.Merge();
                rVerificado.Font.Bold = true;
                rVerificado.Font.Size = 12;
                rVerificado.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rVerificado.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                // Fundo cinzento
                rVerificado.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                // Bordas
                rVerificado.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rVerificado.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Linha 4
                ws2.Cells[4, 20] = "Nome:";
                var rNome2 = ws2.Range[ws2.Cells[4, 21], ws2.Cells[4, 23]];
                rNome2.Merge();
                rNome2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rNome2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Linha 5
                ws2.Cells[5, 20] = "Assinatura:";
                var rAssinaura2 = ws2.Range[ws2.Cells[5, 21], ws2.Cells[5, 23]];
                rAssinaura2.Merge();
                rAssinaura2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rAssinaura2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Linha 6
                ws2.Cells[6, 20] = "Data:";
                var rData2 = ws2.Range[ws2.Cells[6, 21], ws2.Cells[6, 23]];
                rData2.Merge();
                rData2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rData2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                ws2.Cells[6, 21] = DateTime.Now.ToString("d 'de' MMMM 'de' yyyy", new System.Globalization.CultureInfo("pt-PT"));


                ws2.Range[ws2.Cells[4, 20], ws2.Cells[6, 20]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                ws2.Range[ws2.Cells[4, 20], ws2.Cells[6, 20]].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                //----------------------------------------------------------------

                // Linha 8
                // Cabeçalho
                ws2.Cells[8, 2] = "Atividade desenvolvida no Estaleiro";
                var rAtividade = ws2.Range[ws2.Cells[8, 2], ws2.Cells[8, 3]];
                rAtividade.Merge();
                rAtividade.Font.Bold = true;
                rAtividade.Font.Size = 12;
                rAtividade.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                rAtividade.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rAtividade.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rAtividade.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rAtividade.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Linha 9
                ws2.Cells[9, 2] = "Empreiteiro Geral";
                var rEmpreiteiro = ws2.Range[ws2.Cells[9, 2], ws2.Cells[9, 3]];
                rEmpreiteiro.Merge();
                rEmpreiteiro.Font.Bold = true;
                rEmpreiteiro.Font.Size = 12;
                rEmpreiteiro.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rEmpreiteiro.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                Centro(rEmpreiteiro);


                // Linha 10

                ws2.Cells[10, 6] = "Pessoa(s) p/ contacto:";
                var rContacto1 = ws2.Range[ws2.Cells[10, 7], ws2.Cells[10, 9]];
                rContacto1.Merge();

                // Apenas a borda inferior
                rContacto1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rContacto1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                    Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                ws2.Cells[10, 10] = "Telf:";

                var rTelf1 = ws2.Range[ws2.Cells[10, 11], ws2.Cells[10, 13]];
                rTelf1.Merge();

                // Apenas a borda inferior
                rTelf1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rTelf1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                    Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                ws2.Cells[10, 14] = "E-mail:";

                var rEmail1 = ws2.Range[ws2.Cells[10, 15], ws2.Cells[10, 17]];
                rEmail1.Merge();

                // Apenas a borda inferior
                rEmail1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rEmail1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                    Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                ws2.Cells[10, 19] = "Função:";

                var rFuncao1 = ws2.Range[ws2.Cells[10, 20], ws2.Cells[10, 22]];
                rFuncao1.Merge();

                // Apenas a borda inferior
                rFuncao1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rFuncao1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                    Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;


                // Linha 11

                ws2.Cells[11, 6] = "Pessoa(s) p/ contacto:";
                var rContacto2 = ws2.Range[ws2.Cells[11, 7], ws2.Cells[11, 9]];
                rContacto2.Merge();

                // Apenas a borda inferior
                rContacto2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rContacto2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                    Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                ws2.Cells[11, 10] = "Telf:";

                var rTelf2 = ws2.Range[ws2.Cells[11, 11], ws2.Cells[11, 13]];
                rTelf2.Merge();

                // Apenas a borda inferior
                rTelf2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rTelf2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                    Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                ws2.Cells[11, 14] = "E-mail:";

                var rEmail2 = ws2.Range[ws2.Cells[11, 15], ws2.Cells[11, 17]];
                rEmail2.Merge();

                // Apenas a borda inferior
                rEmail2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rEmail2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                    Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                ws2.Cells[11, 19] = "Função:";

                var rFuncao2 = ws2.Range[ws2.Cells[11, 20], ws2.Cells[11, 22]];
                rFuncao2.Merge();

                // Apenas a borda inferior
                rFuncao2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rFuncao2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                    Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;



                // Linha 13 e 14
                ws2.Cells[13, 1] = "EMPRESA";
                // Mescla linha 13 e 14, colunas 1 a 5
                var rEmpresa = ws2.Range[ws2.Cells[13, 1], ws2.Cells[14, 5]];
                rEmpresa.Merge();
                rEmpresa.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rEmpresa.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rEmpresa.Font.Bold = true;
                rEmpresa.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rEmpresa.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                rEmpresa.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                //Linha 13

                // Numeração 1..23 (col 5..23)
                int lastCol = 28;
                int numRow = 13;

                for (int c = 6, n = 1; c <= lastCol && n <= 23; n++)
                {
                    Microsoft.Office.Interop.Excel.Range r;

                    // Se o número precisa ocupar 2 colunas (2 ou 9)
                    if (n == 2 || n == 9)
                    {
                        r = ws2.Range[ws2.Cells[numRow, c], ws2.Cells[numRow, c + 1]];
                        r.Merge();
                        ws2.Cells[numRow, c] = n.ToString();
                        c += 2; // avança duas colunas
                    }
                    else
                    {
                        // Célula normal
                        r = ws2.Range[ws2.Cells[numRow, c], ws2.Cells[numRow, c]];
                        ws2.Cells[numRow, c] = n.ToString();
                        c++; // avança uma coluna
                    }

                    // Estilos
                    Negrito(r);
                    Centro(r);
                    Borda(r);
                    r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                }

                // Linha 14

                ws2.Cells[14, 6] = "Contribuinte";
                var rContribuinte = ws2.Cells[14, 6]; // já é um Range
                Negrito(rContribuinte);
                Borda(rContribuinte);
                Centro(rContribuinte);

                ws2.Cells[14, 7] = "Alvará / Certificado";
                var rAlvara = ws2.Range[ws2.Cells[14, 7], ws2.Cells[14, 8]];
                rAlvara.Merge();
                Negrito(rAlvara);
                Borda(rAlvara);
                Centro(rAlvara);

                ws2.Cells[14, 9] = "Anexo D";
                var rAnexo = ws2.Cells[14, 9];
                Negrito(rAnexo);
                Borda(rAnexo);
                Centro(rAnexo);

                ws2.Cells[14, 10] = "Cert. ND Finanças";
                var rNDFinanças = ws2.Cells[14, 10] ;
                Negrito(rNDFinanças);
                Borda(rNDFinanças);
                Centro(rNDFinanças);

                ws2.Cells[14, 11] = "Decl. ND Seg. Social";
                var rDeclSegSocial = ws2.Cells[14, 11];
                Negrito(rDeclSegSocial);
                Borda(rDeclSegSocial);
                Centro(rDeclSegSocial);

                ws2.Cells[14, 12] = "Folha Pag. Seg. Social";
                var rFolhaPagSegSocial = ws2.Cells[14, 12];
                Negrito(rFolhaPagSegSocial);
                Borda(rFolhaPagSegSocial);
                Centro(rFolhaPagSegSocial);

                ws2.Cells[14, 13] = "Recibo de Pag. Seg. Social";
                var rRecibodePagSegSocial = ws2.Cells[14, 13];
                Negrito(rRecibodePagSegSocial);
                Borda(rRecibodePagSegSocial);
                Centro(rRecibodePagSegSocial);

                ws2.Cells[14, 14] = "Apólice AT";
                var rApoliceAT = ws2.Cells[14, 14];
                Negrito(rApoliceAT);
                Borda(rApoliceAT);
                Centro(rApoliceAT);

                ws2.Cells[14, 15] = "Modalidade do Seguro";
                var rModalidadeSeguro = ws2.Range[ws2.Cells[14, 15], ws2.Cells[14, 16]];
                rModalidadeSeguro.Merge();
                Negrito(rModalidadeSeguro);
                Borda(rModalidadeSeguro);
                Centro(rModalidadeSeguro);

                ws2.Cells[14, 17] = "Recibo Apólice AT";
                var rReciboApoliceAT = ws2.Cells[14, 17];
                Negrito(rReciboApoliceAT);
                Borda(rReciboApoliceAT);
                Centro(rReciboApoliceAT);

                ws2.Cells[14, 18] = "Apólice RC";
                var rApoliceRC = ws2.Cells[14, 18];
                Negrito(rApoliceRC);
                Borda(rApoliceRC);
                Centro(rApoliceRC);

                ws2.Cells[14, 19] = "Recibo RC";
                var rReciboRC = ws2.Cells[14, 19];
                Negrito(rReciboRC);
                Borda(rReciboRC);
                Centro(rReciboRC);

                ws2.Cells[14, 20] = "Registo(s) Criminal(ais)";
                var rRegistoCriminal = ws2.Cells[14, 20];
                Negrito(rRegistoCriminal);
                Borda(rRegistoCriminal);
                Centro(rRegistoCriminal);

                ws2.Cells[14, 21] = "Horário de Trabalho";
                var rHorarioTrabalho = ws2.Cells[14, 21];
                Negrito(rHorarioTrabalho);
                Borda(rHorarioTrabalho);
                Centro(rHorarioTrabalho);

                ws2.Cells[14, 22] = "Dec.  Trab. Imigr.";
                var rDecTrabImigr = ws2.Cells[14, 22];
                Negrito(rDecTrabImigr);
                Borda(rDecTrabImigr);
                Centro(rDecTrabImigr);

                ws2.Cells[14, 23] = "Dec. Resp. Estaleiro";
                var rDecRespEstaleiro = ws2.Cells[14, 23];
                Negrito(rDecRespEstaleiro);
                Borda(rDecRespEstaleiro);
                Centro(rDecRespEstaleiro);

                ws2.Cells[14, 24] = "Dec. Ades. PSS";
                var rDecAdesPSS = ws2.Cells[14, 24];
                Negrito(rDecAdesPSS);
                Borda(rDecAdesPSS);
                Centro(rDecAdesPSS);

                ws2.Cells[14, 25] = "Contrato Subempreitada";
                var rContratoSubempreitada = ws2.Cells[14, 25];
                Negrito(rContratoSubempreitada);
                Borda(rContratoSubempreitada);
                Centro(rContratoSubempreitada);

                ws2.Cells[14, 26] = "Entrada em Obra";
                var rEntradaObra = ws2.Cells[14, 26];
                Negrito(rEntradaObra);
                Borda(rEntradaObra);
                Centro(rEntradaObra);

                ws2.Cells[14, 27] = "Saída de Obra";
                var rSaidaObra = ws2.Cells[14, 27];
                Negrito(rSaidaObra);
                Borda(rSaidaObra);
                Centro(rSaidaObra);

                ws2.Cells[14, 28] = "Autorização de Entrada";
                var rAutorizacaoEntrada = ws2.Cells[14, 28];
                Negrito(rAutorizacaoEntrada);
                Borda(rAutorizacaoEntrada);
                Centro(rAutorizacaoEntrada);

                // Linha 15

                ws2.Cells[15, 1] = "N.º";
                var rNum = ws2.Cells[15, 1];
                Centro(rNum);
                Borda(rNum);

                ws2.Cells[15, 2] = "Designação Social";
                var rDesignacaoSocial = ws2.Cells[15, 2];
                Centro(rDesignacaoSocial);
                Borda(rDesignacaoSocial);

                ws2.Cells[15, 3] = "Sede";
                var rSede = ws2.Range[ws2.Cells[15, 3], ws2.Cells[15, 5]];
                rSede.Merge();
                Centro(rSede);
                Borda(rSede);

                ws2.Cells[15, 6] = "N.º";
                var rNum2 = ws2.Cells[15, 6];
                Centro(rNum2);
                Borda(rNum2);

                ws2.Cells[15, 7] = "N.º";
                var rNum3 = ws2.Cells[15, 7];
                Centro(rNum3);
                Borda(rNum3);

                ws2.Cells[15, 8] = "PUB / PAR";
                var rPubPar = ws2.Cells[15, 8];
                Centro(rPubPar);
                Borda(rPubPar);

                ws2.Cells[15, 9] = "C ; N/C ; N/A";
                var rCNCNA = ws2.Cells[15, 9];
                Centro(rCNCNA);
                Borda(rCNCNA);

                ws2.Cells[15, 10] = "Validade";
                var rValidade = ws2.Cells[15, 10];
                Centro(rValidade);
                Borda(rValidade);

                ws2.Cells[15, 11] = "Validade";
                var rValidade2 = ws2.Cells[15, 11];
                Centro(rValidade2);
                Borda(rValidade2);

                ws2.Cells[15, 12] = "Validade";
                var rValidade3 = ws2.Cells[15, 12];
                Centro(rValidade3);
                Borda(rValidade3);

                ws2.Cells[15, 13] = "C ; N/C ; N/A";
                var rCNCNA2 = ws2.Cells[15, 13];
                Centro(rCNCNA2);
                Borda(rCNCNA2);

                ws2.Cells[15, 14] = "N.º";
                var rNum4 = ws2.Cells[15, 14];
                Centro(rNum4);
                Borda(rNum4);

                ws2.Cells[15, 15] = "Fixo?";
                var rFixo = ws2.Cells[15, 15];
                Centro(rFixo);
                Borda(rFixo);

                ws2.Cells[15, 16] = "Prémio Variável?";
                var rPremioVariavel = ws2.Cells[15, 16];
                Centro(rPremioVariavel);
                Borda(rPremioVariavel);

                ws2.Cells[15, 17] = "Validade";
                var rValidade4 = ws2.Cells[15, 17];
                Centro(rValidade4);
                Borda(rValidade4);

                ws2.Cells[15, 18] = "N.º";
                var rNum5 = ws2.Cells[15, 18];
                Centro(rNum5);
                Borda(rNum5);

                ws2.Cells[15, 19] = "Validade";
                var rValidade5 = ws2.Cells[15, 19];
                Centro(rValidade5);
                Borda(rValidade5);

                ws2.Cells[15, 20] = "C ; N/C ; N/A";
                var rCNCNA3 = ws2.Cells[15, 20];
                Centro(rCNCNA3);
                Borda(rCNCNA3);

                ws2.Cells[15, 21] = "C ; N/C ; N/A";
                var rCNCNA4 = ws2.Cells[15, 21];
                Centro(rCNCNA4);
                Borda(rCNCNA4);

                ws2.Cells[15, 22] = "C ; N/C ; N/A";
                var rCNCNA5 = ws2.Cells[15, 22];
                Centro(rCNCNA5);
                Borda(rCNCNA5);

                ws2.Cells[15, 23] = "C ; N/C ; N/A";
                var rCNCNA6 = ws2.Cells[15, 23];
                Centro(rCNCNA6);
                Borda(rCNCNA6);

                ws2.Cells[15, 24] = "C ; N/C ; N/A";
                var rCNCNA7 = ws2.Cells[15, 24];
                Centro(rCNCNA7);
                Borda(rCNCNA7);

                ws2.Cells[15, 25] = "Com:";
                var rCom = ws2.Cells[15, 25];
                Centro(rCom);
                Borda(rCom);

                ws2.Cells[15, 26] = "Data";
                var rData3 = ws2.Cells[15, 26];
                Centro(rData3);
                Borda(rData3);

                ws2.Cells[15, 27] = "Data";
                var rData4 = ws2.Cells[15, 27];
                Centro(rData4);
                Borda(rData4);

                ws2.Cells[15, 28] = "Sim / Não";
                var rSimNao = ws2.Cells[15, 28];
                Centro(rSimNao);
                Borda(rSimNao);

                // Linha 16 Dados Jpa

                ws2.Cells[16, 1] = "1";
                var rNumJPA = ws2.Cells[16, 1];
                Centro(rNumJPA);

                ws2.Cells[16, 2] = "JOAQUIM PEIXOTO AZEVEDO & FILHOS, LDA";

                ws2.Cells[16, 3] = "RUA DE LONGRAS Nº 44";
                var rSedeJPA = ws2.Range[ws2.Cells[16, 3], ws2.Cells[16, 5]];
                rSedeJPA.Merge();

                ws2.Cells[16, 6] = "502244585";

                ws2.Cells[16, 7] = "";
                ws2.Cells[16, 8] = "PAR";
                ws2.Cells[16, 9] = "C";
                ws2.Cells[16, 10] = "";
                ws2.Cells[16, 11] = "";
                ws2.Cells[16, 12] = "";
                ws2.Cells[16, 13] = "C";
                ws2.Cells[16, 14] = "";
                ws2.Cells[16, 15] = "";
                ws2.Cells[16, 16] = "";
                ws2.Cells[16, 17] = "";
                ws2.Cells[16, 18] = "";
                ws2.Cells[16, 19] = "";
                ws2.Cells[16, 20] = "C";
                ws2.Cells[16, 21] = "C";
                ws2.Cells[16, 22] = "C";
                ws2.Cells[16, 23] = "C";
                ws2.Cells[16, 24] = "C";
                ws2.Cells[16, 25] = "Sim";
                ws2.Cells[16, 26] = ""; 
                ws2.Cells[16, 27] = "";
                ws2.Cells[16, 28] = "Sim";

                //*************************************
                // Linha 18 a 20
                ws2.Cells[18, 1] = "TRABALHADORES";
                var rTRABALHADORES = ws2.Range[ws2.Cells[18, 1], ws2.Cells[20, 4]];
                rTRABALHADORES.Merge();
                rTRABALHADORES.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                rTRABALHADORES.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rTRABALHADORES.Font.Bold = true;
                rTRABALHADORES.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rTRABALHADORES.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                rTRABALHADORES.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                // Linha 18

                // Numeração 1..23 (col 5..23)
                int lastCol2 = 28;
                int numRow2 = 18;

                for (int c = 5, n = 1; c <= lastCol2 && n <= 23; n++)
                {
                    Microsoft.Office.Interop.Excel.Range r;
                  
                    if (n == 5   || n == 12 || n == 6 || n == 13 || n == 9)
                    {
                        r = ws2.Range[ws2.Cells[numRow2, c], ws2.Cells[numRow2, c + 1]];
                        r.Merge();
                        ws2.Cells[numRow2, c] = n.ToString();
                        c += 2; // avança duas colunas
                    }else if(n == 10)// n== 10 ocupa 3 colunas
                    {
                        r = ws2.Range[ws2.Cells[numRow2, c], ws2.Cells[numRow2, c + 2]];
                        r.Merge();
                        ws2.Cells[numRow2, c] = n.ToString();
                        c += 3; // avança três colunas
                    }
                    else if (n == 11)// n== 11 ocupa 4 colunas 
                    {
                        r = ws2.Range[ws2.Cells[numRow2, c], ws2.Cells[numRow2, c + 3]];
                        r.Merge();
                        ws2.Cells[numRow2, c] = n.ToString();
                        c += 4; // avança quatro colunas
                    }
                    else
                    {
                        // Célula normal
                        r = ws2.Range[ws2.Cells[numRow2, c], ws2.Cells[numRow2, c]];
                        ws2.Cells[numRow2, c] = n.ToString();
                        c++; // avança uma coluna
                    }

                    // Estilos
                    Negrito(r);
                    Centro(r);
                    Borda(r);
                    r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                }

                // Linha 19

                ws2.Cells[19, 5] = "Categoria / Função";
                var rCategoriaFuncao = ws2.Range[ws2.Cells[19, 5], ws2.Cells[21, 5]]; // ocupa 3 linhas
                rCategoriaFuncao.Merge();
                Negrito(rCategoriaFuncao);
                Borda(rCategoriaFuncao);
                Centro(rCategoriaFuncao);


                ws2.Cells[19, 6] = "CAP (se aplicável)";
                var rCAP = ws2.Range[ws2.Cells[19, 6], ws2.Cells[20, 6]];
                rCAP.Merge();
                Negrito(rCAP);
                Borda(rCAP);
                Centro(rCAP);
               
                ws2.Cells[19, 7] = "Contribuinte";
                var rContribuinte2 = ws2.Range[ws2.Cells[19, 7], ws2.Cells[20, 7]];
                rContribuinte2.Merge();
                Negrito(rContribuinte2);
                Borda(rContribuinte2);
                Centro(rContribuinte2);

                ws2.Cells[19, 8] = "Segurança Social";
                var rSegurancaSocial = ws2.Range[ws2.Cells[19, 8], ws2.Cells[20, 8]];
                rSegurancaSocial.Merge();
                Negrito(rSegurancaSocial);
                Borda(rSegurancaSocial);
                Centro(rSegurancaSocial);

                ws2.Cells[19, 9] = "Cartão de Cidadão";
                var rFichaAptidaoMedica = ws2.Range[ws2.Cells[19, 9], ws2.Cells[20,10]];
                rFichaAptidaoMedica.Merge();
                Negrito(rFichaAptidaoMedica);
                Borda(rFichaAptidaoMedica);
                Centro(rFichaAptidaoMedica);

                ws2.Cells[19, 11] = "Ficha de Aptidão Médica";
                var rFichaDistribuicaoEPI = ws2.Range[ws2.Cells[19, 11], ws2.Cells[19, 12]];
                rFichaDistribuicaoEPI.Merge();
                Negrito(rFichaDistribuicaoEPI);
                Borda(rFichaDistribuicaoEPI);
                Centro(rFichaDistribuicaoEPI);

                ws2.Cells[19, 13] = "Ficha de Distribuição de EPI";
                var rConstaMapaSS = ws2.Range[ws2.Cells[19, 13], ws2.Cells[20,13]];
                rConstaMapaSS.Merge();
                Negrito(rConstaMapaSS);
                Borda(rConstaMapaSS);
                Centro(rConstaMapaSS);

                ws2.Cells[19, 14] = "Consta no Mapa  SS / Inscrito?";
                var rValidade6 = ws2.Range[ws2.Cells[19,14], ws2.Cells[20,14]];
                rValidade6.Merge();
                Negrito(rValidade6);
                Borda(rValidade6);
                Centro(rValidade6);

                ws2.Cells[19, 15] = "Admissão na SS (caso não conste no Mapa da SS)";
                var rTrabalhadorEstrangeiro = ws2.Range[ws2.Cells[19, 15], ws2.Cells[20, 16]];
                rTrabalhadorEstrangeiro.Merge();
                Negrito(rTrabalhadorEstrangeiro);
                Borda(rTrabalhadorEstrangeiro);
                Centro(rTrabalhadorEstrangeiro);

                ws2.Cells[19, 17] = "Trabalhador Estrangeiro";
                var rTrabalhadorEstrangeiro2 = ws2.Range[ws2.Cells[19, 17], ws2.Cells[19, 19]];
                rTrabalhadorEstrangeiro2.Merge();
                Negrito(rTrabalhadorEstrangeiro2);
                Borda(rTrabalhadorEstrangeiro2);
                Centro(rTrabalhadorEstrangeiro2);

                ws2.Cells[19, 20] = "Formação / Informação";
                var rFormacaoInformacao = ws2.Range[ws2.Cells[19, 20], ws2.Cells[19, 23]];
                rFormacaoInformacao.Merge();
                Negrito(rFormacaoInformacao);
                Borda(rFormacaoInformacao);
                Centro(rFormacaoInformacao);

                ws2.Cells[19, 24] = "Cadastro";
                var rCadastro = ws2.Range[ws2.Cells[19, 24], ws2.Cells[19, 25]];
                rCadastro.Merge();
                Negrito(rCadastro);
                Borda(rCadastro);
                Centro(rCadastro);

                ws2.Cells[19, 26] = "Entrada Obra";
                var rEntradaObra2 = ws2.Range[ws2.Cells[19, 26], ws2.Cells[20, 26]];
                rEntradaObra2.Merge();
                Negrito(rEntradaObra2);
                Borda(rEntradaObra2);
                Centro(rEntradaObra2);

                ws2.Cells[19, 27] = "Saída Obra";
                var rSaidaObra2 = ws2.Range[ws2.Cells[19, 27], ws2.Cells[20, 27]];
                rSaidaObra2.Merge();
                Negrito(rSaidaObra2);
                Borda(rSaidaObra2);
                Centro(rSaidaObra2);

                ws2.Cells[19, 28] = "Autorização de Entrada em Obra";
                var rAutorizacaoEntrada2 = ws2.Range[ws2.Cells[19, 28], ws2.Cells[20, 28]];
                rAutorizacaoEntrada2.Merge();
                Negrito(rAutorizacaoEntrada2);
                Borda(rAutorizacaoEntrada2);
                Centro(rAutorizacaoEntrada2);


                // LInha 20
                ws2.Cells[20, 11] = "Conforme Cat. Prof.?";
                var rConformeCatProf = ws2.Cells[20, 11];
                Negrito(rConformeCatProf);
                Borda(rConformeCatProf);
                Centro(rConformeCatProf);

                ws2.Cells[20, 12] = "Validade";
                var rValidade7 = ws2.Range[ws2.Cells[20, 12], ws2.Cells[21, 12]];
                rValidade7.Merge();
                Negrito(rValidade7);
                Borda(rValidade7);
                Centro(rValidade7);


                ws2.Cells[20, 17] = "Passaporte c/visto / Titulo de Residência";
                var rPassaporteVisto = ws2.Range[ws2.Cells[20, 17], ws2.Cells[20, 19]];
                rPassaporteVisto.Merge();
                Negrito(rPassaporteVisto);
                Borda(rPassaporteVisto);
                Centro(rPassaporteVisto);

                ws2.Cells[20, 20] = "Acolhimento";
                var rAcolhimento = ws2.Cells[20, 20];
                Negrito(rAcolhimento);
                Borda(rAcolhimento);
                Centro(rAcolhimento);

                ws2.Cells[20, 21] = "Específica 1";
                var rEspecifica1 = ws2.Cells[20, 21];
                Negrito(rEspecifica1);
                Borda(rEspecifica1);
                Centro(rEspecifica1);

                ws2.Cells[20, 22] = "Específica 2";
                var rEspecifica2 = ws2.Cells[20, 22];
                Negrito(rEspecifica2);
                Borda(rEspecifica2);
                Centro(rEspecifica2);

                ws2.Cells[20, 23] = "Específica 3";
                var rEspecifica3 = ws2.Cells[20, 23];
                Negrito(rEspecifica3);
                Borda(rEspecifica3);
                Centro(rEspecifica3);

                ws2.Cells[20, 24] = "1.º Aviso";
                var rPrimeiroAviso = ws2.Cells[20, 24];
                Negrito(rPrimeiroAviso);
                Borda(rPrimeiroAviso);
                Centro(rPrimeiroAviso);

                ws2.Cells[20, 25] = "2.º Aviso";
                var rSegundoAviso = ws2.Cells[20, 25];
                Negrito(rSegundoAviso);
                Borda(rSegundoAviso);
                Centro(rSegundoAviso);

                // Linha 21

                ws2.Cells[21, 1] = "N.º";
                var rNumTrabalhador = ws2.Cells[21, 1];
                Centro(rNumTrabalhador);
                Borda(rNumTrabalhador);

                ws2.Cells[21, 2] = "Nome Completo";
                var rNomeCompleto = ws2.Cells[21, 2];
                Centro(rNomeCompleto);
                Borda(rNomeCompleto);

                ws2.Cells[21, 3] = "Residência Habitual";
                var rResidenciaHabitual = ws2.Cells[21, 3];
                Centro(rResidenciaHabitual);
                Borda(rResidenciaHabitual);

                ws2.Cells[21, 4] = "Nacionalidade";
                var rNacionalidade = ws2.Cells[21, 4];
                Centro(rNacionalidade);
                Borda(rNacionalidade);

                ws2.Cells[21, 6] = "N.º";
                var rNumCAP = ws2.Cells[21, 6];
                Centro(rNumCAP);
                Borda(rNumCAP);

                ws2.Cells[21, 7] = "N.º";
                var rNumContribuinte = ws2.Cells[21, 7];
                Centro(rNumContribuinte);
                Borda(rNumContribuinte);

                ws2.Cells[21, 8] = "N.º";
                var rNumSegSocial = ws2.Cells[21, 8];
                Centro(rNumSegSocial);
                Borda(rNumSegSocial);

                ws2.Cells[21, 9] = "N.º";
                var rNumCC = ws2.Cells[21, 9];
                Centro(rNumCC);
                Borda(rNumCC);

                ws2.Cells[21, 10] = "Validade";
                var rValidade8 = ws2.Cells[21, 10];
                Centro(rValidade8);
                Borda(rValidade8);

                ws2.Cells[21, 11] = "C ; N/C ; N/A";
                var rCNCNA8 = ws2.Cells[21, 11];
                Centro(rCNCNA8);
                Borda(rCNCNA8);

                ws2.Cells[21, 13] = "C ; N/C";
                var rCNCNA9 = ws2.Cells[21, 13];
                Centro(rCNCNA9);
                Borda(rCNCNA9);

                ws2.Cells[21, 14] = "C ; N/C ; N/A";
                var rCNCNA10 = ws2.Cells[21, 14];
                Centro(rCNCNA10);
                Borda(rCNCNA10);

                ws2.Cells[21, 15] = "Data";
                var rData5 =  ws2.Range[ws2.Cells[21, 15], ws2.Cells[21, 16]];
                rData5.Merge();
                Centro(rData5);
                Borda(rData5);

                ws2.Cells[21, 17] = "Tipo de documento";
                var rTipoDocumento = ws2.Cells[21, 17];
                Centro(rTipoDocumento);
                Borda(rTipoDocumento);

                ws2.Cells[21, 18] = "Número";
                var rNumeroDoc = ws2.Cells[21, 18];
                Centro(rNumeroDoc);
                Borda(rNumeroDoc);

                ws2.Cells[21, 19] = "Validade";
                var rValidade9 = ws2.Cells[21, 19];
                Centro(rValidade9);
                Borda(rValidade9);

                ws2.Cells[21, 20] = "Data";
                var rData6 = ws2.Cells[21, 20];
                Centro(rData6);
                Borda(rData6);

                ws2.Cells[21, 21] = "Data";
                var rData7 = ws2.Cells[21, 21];
                Centro(rData7);
                Borda(rData7);

                ws2.Cells[21, 22] = "Data";
                var rData8 = ws2.Cells[21, 22];
                Centro(rData8);
                Borda(rData8);

                ws2.Cells[21, 23] = "Data";
                var rData9 = ws2.Cells[21, 23];
                Centro(rData9);
                Borda(rData9);

                ws2.Cells[21, 24] = "Data";
                var rData10 = ws2.Cells[21, 24];
                Centro(rData10);
                Borda(rData10);

                ws2.Cells[21, 25] = "Data";
                var rData11 = ws2.Cells[21, 25];
                Centro(rData11);
                Borda(rData11);

                ws2.Cells[21, 26] = "Data";
                var rData12 = ws2.Cells[21, 26];
                Centro(rData12);
                Borda(rData12);

                ws2.Cells[21, 27] = "Data";
                var rData13 = ws2.Cells[21, 27];
                Centro(rData13);
                Borda(rData13);

                ws2.Cells[21, 28] = "Sim / Não";
                var rSimNao2 = ws2.Cells[21, 28];
                Centro(rSimNao2);
                Borda(rSimNao2);

                //DADOS NAS LINHAS //********************************************************************** TRABALHADORES JPA
         
                var queryObraPAITrabalhadores = $@"SELECT 
    f.*, 
    p.Descricao AS ProfissaoDescricao
FROM COP_Obras o
JOIN COP_Obras_Pessoal op ON op.obraId = o.id
JOIN GPR_Operadores g ON g.idOperador = op.colaboradorID
JOIN Funcionarios f ON f.codigo = g.Operador
LEFT JOIN Profissoes p ON f.Profissao = p.Profissao
WHERE o.Codigo = '{codigoObra}';
;";
                var dadosTrabalhadores = BSO.Consulta(queryObraPAITrabalhadores);

                var numregistos = dadosTrabalhadores.NumLinhas();

                int linhaAtual = 22;
                dadosTrabalhadores.Inicio();
                for (int i = 0; i < numregistos; i++)
                {
                  
                    ws2.Cells[linhaAtual, 1] = (i + 1).ToString(); // N.º
                    ws2.Cells[linhaAtual, 2] = dadosTrabalhadores.DaValor<string>("Nome"); // Nome Completo
                    ws2.Cells[linhaAtual, 3] = dadosTrabalhadores.DaValor<string>("Morada"); // Residência Habitual
                    ws2.Cells[linhaAtual, 4] = dadosTrabalhadores.DaValor<string>("Nacionalidade"); // Nacionalidade
                    ws2.Cells[linhaAtual, 5] = dadosTrabalhadores.DaValor<string>("ProfissaoDescricao"); // Categoria / Função
                    //ws2.Cells[linhaAtual, 6] = dadosTrabalhadores.DaValor<string>("CAP", i); // CAP (se aplicável)
                    ws2.Cells[linhaAtual, 7] = dadosTrabalhadores.DaValor<string>("NumContr"); // Contribuinte
                    ws2.Cells[linhaAtual, 8] = dadosTrabalhadores.DaValor<string>("NumBeneficiario"); // Segurança Social
                    ws2.Cells[linhaAtual, 9] = dadosTrabalhadores.DaValor<string>("NumBI"); // Cartão de Cidadão
                    if (DateTime.TryParse(dadosTrabalhadores.DaValor<string>("DataValidadeBI"), out DateTime data))
                        ws2.Cells[linhaAtual, 10] = data.ToString("dd/MM/yyyy");
                    else
                        ws2.Cells[linhaAtual, 10] = "";

                    ws2.Cells[linhaAtual, 11] = "C"; // Conforme Cat. Prof.?
                    ws2.Cells[linhaAtual, 12] = ""; // Validade Ficha de Aptidão Médica
                    ws2.Cells[linhaAtual, 13] = "C"; // Ficha de Distribuição de EPI
                    ws2.Cells[linhaAtual, 14] = "C"; // Consta no Mapa  SS / Inscrito?
                    ws2.Cells[linhaAtual, 15] = ""; // Admissão na SS (caso não conste no Mapa da SS)
                    ws2.Cells[linhaAtual, 16] = ""; // Admissão na SS (caso não conste no Mapa da SS)
                    ws2.Cells[linhaAtual, 17] = ""; // Passaporte c/visto / Titulo de Residência
                    ws2.Cells[linhaAtual, 18] = ""; // Validade
                    ws2.Cells[linhaAtual, 19] = ""; // Data
                    ws2.Cells[linhaAtual, 20] = ""; // Acolhimento
                    ws2.Cells[linhaAtual, 21] = ""; // Específica 1
                    ws2.Cells[linhaAtual, 22] = ""; // Específica 2
                    ws2.Cells[linhaAtual, 23] = ""; // Específica 3
                    ws2.Cells[linhaAtual, 24] = ""; // 1.º Aviso
                    ws2.Cells[linhaAtual, 25] = ""; // 2.º Aviso
                    ws2.Cells[linhaAtual, 26] = ""; // Entrada Obra
                    ws2.Cells[linhaAtual, 27] = ""; // Saída Obra
                    ws2.Cells[linhaAtual, 28] = "Sim"; // Autorização de Entrada em Obra

                    linhaAtual++;    // Começa na linha 22
                    dadosTrabalhadores.Seguinte();
                }
                //MÁQUINAS E EQUIPAMENTOS
                linhaAtual = linhaAtual + 1;
                ws2.Cells[linhaAtual, 1] = "MÁQUINAS E EQUIPAMENTOS";
                var rMaquinasEquipamentos = ws2.Range[ws2.Cells[linhaAtual, 1], ws2.Cells[linhaAtual + 2, 7]];
                rMaquinasEquipamentos.Merge();
                Negrito(rMaquinasEquipamentos);
                Centro(rMaquinasEquipamentos);
                Borda(rMaquinasEquipamentos);
                rMaquinasEquipamentos.Interior.Color = ToOle(System.Drawing.Color.LightGray);

                //cilco for do numero 1 ao 11 mas o numero 4 vai ocupar 2 colunas , numero 5 vai ocupar 3 colunas, numero 6 vai ocupar 5 colunas, numero 7 vai ocupar 3 colunas, e o numero 8 vai ocupar 2 colunas
                int lastCol3 = 28;
                int numRow3 = linhaAtual;
                for (int c = 8, n = 1; c <= lastCol3 && n <= 11; n++)
                {
                    Microsoft.Office.Interop.Excel.Range r;

                    if (n == 4) // n== 4 ocupa 2 colunas
                    {
                        r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c + 1]];
                        r.Merge();
                        ws2.Cells[numRow3, c] = n.ToString();
                        c += 2; // avança duas colunas
                    }
                    else if (n == 5) // n== 5 ocupa 3 colunas
                    {
                        r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c + 3]];
                        r.Merge();
                        ws2.Cells[numRow3, c] = n.ToString();
                        c += 4; // avança três colunas
                    }
                    else if (n == 6) // n== 6 ocupa 5 colunas 
                    {
                        r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c + 4]];
                        r.Merge();
                        ws2.Cells[numRow3, c] = n.ToString();
                        c += 5; // avança cinco colunas
                    }
                    else if (n == 7) // n== 7 ocupa 3 colunas 
                    {
                        r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c + 2]];
                        r.Merge();
                        ws2.Cells[numRow3, c] = n.ToString();
                        c += 3; // avança três colunas
                    }
                    else
                    {
                        // Célula normal
                        r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c]];
                        ws2.Cells[numRow3
                            , c] = n.ToString();
                        c++; // avança uma coluna
                    }
                    // Estilos
                    Negrito(r);
                    Centro(r);
                    Borda(r);
                    r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                }

                // Linha abaixo dos números (Cabeçalhos)
                linhaAtual = linhaAtual + 1;
                ws2.Cells[linhaAtual, 8] = "Manual de Instruções em língua PT";
                var rManualInstrucoes = ws2.Range[ws2.Cells[linhaAtual, 8], ws2.Cells[linhaAtual + 1, 8]];
                rManualInstrucoes.Merge();
                Negrito(rManualInstrucoes);
                Borda(rManualInstrucoes);
                Centro(rManualInstrucoes);

                ws2.Cells[linhaAtual, 9] = "Declaração  Conformidade CE";
                var rDeclaracaoConformidadeCE = ws2.Range[ws2.Cells[linhaAtual, 9], ws2.Cells[linhaAtual + 1, 9]];
                rDeclaracaoConformidadeCE.Merge();
                Negrito(rDeclaracaoConformidadeCE);
                Borda(rDeclaracaoConformidadeCE);
                Centro(rDeclaracaoConformidadeCE);

                ws2.Cells[linhaAtual, 10] = "Plano de Manutenção";
                var rPlanoManutencao = ws2.Range[ws2.Cells[linhaAtual, 10], ws2.Cells[linhaAtual + 1, 10]];
                rPlanoManutencao.Merge();
                Negrito(rPlanoManutencao);
                Borda(rPlanoManutencao);
                Centro(rPlanoManutencao);

                ws2.Cells[linhaAtual, 11] = "Relatório de Verificação de Segurança (DL 50/2005)";
                var rRelatorioVerificacaoSeguranca = ws2.Range[ws2.Cells[linhaAtual, 11], ws2.Cells[linhaAtual, 12]];
                rRelatorioVerificacaoSeguranca.Merge();
                Negrito(rRelatorioVerificacaoSeguranca);
                Borda(rRelatorioVerificacaoSeguranca);
                Centro(rRelatorioVerificacaoSeguranca);

                ws2.Cells[linhaAtual, 13] = "Registo de Manutenção (Último)";
                var rRegistoManutencao = ws2.Range[ws2.Cells[linhaAtual, 13], ws2.Cells[linhaAtual + 1, 16]];
                rRegistoManutencao.Merge();
                Negrito(rRegistoManutencao);
                Borda(rRegistoManutencao);
                Centro(rRegistoManutencao);

                ws2.Cells[linhaAtual, 17] = "Seguro Casco (carta verde)";
                var rSeguroCasco = ws2.Range[ws2.Cells[linhaAtual, 17], ws2.Cells[linhaAtual + 1, 21]];
                rSeguroCasco.Merge();
                Negrito(rSeguroCasco);
                Borda(rSeguroCasco);
                Centro(rSeguroCasco);

                ws2.Cells[linhaAtual, 22] = "Manobrador (Se Aplicável)";
                var rManobrador = ws2.Range[ws2.Cells[linhaAtual, 22], ws2.Cells[linhaAtual , 24]];
                rManobrador.Merge();
                Negrito(rManobrador);
                Borda(rManobrador);
                Centro(rManobrador);

                ws2.Cells[linhaAtual, 25] = "Observações";
                var rObservacoes = ws2.Range[ws2.Cells[linhaAtual, 25], ws2.Cells[linhaAtual + 2, 25]];
                rObservacoes.Merge();
                Negrito(rObservacoes);
                Borda(rObservacoes);
                Centro(rObservacoes);

                ws2.Cells[linhaAtual, 26] = "Entrada em Obra";
                var rEntradaObra3 = ws2.Range[ws2.Cells[linhaAtual, 26], ws2.Cells[linhaAtual + 1, 26]];
                rEntradaObra3.Merge();
                Negrito(rEntradaObra3);
                Borda(rEntradaObra3);
                Centro(rEntradaObra3);

                ws2.Cells[linhaAtual, 27] = "Saída de Obra";
                var rSaidaObra3 = ws2.Range[ws2.Cells[linhaAtual, 27], ws2.Cells[linhaAtual + 1, 27]];
                rSaidaObra3.Merge();
                Negrito(rSaidaObra3);
                Borda(rSaidaObra3);
                Centro(rSaidaObra3);

                ws2.Cells[linhaAtual, 28] = "Autorização de Entrada em Obra";
                var rAutorizacaoEntrada3 = ws2.Range[ws2.Cells[linhaAtual, 28], ws2.Cells[linhaAtual + 1, 28]];
                rAutorizacaoEntrada3.Merge();
                Negrito(rAutorizacaoEntrada3);
                Borda(rAutorizacaoEntrada3);
                Centro(rAutorizacaoEntrada3);


                //
                linhaAtual = linhaAtual + 1;
                //coluna 11
                
                ws2.Cells[linhaAtual, 11] = "Possui?";
                var rPossui = ws2.Cells[linhaAtual, 11];
                Negrito(rPossui);
                Borda(rPossui);
                Centro(rPossui);

                ws2.Cells[linhaAtual, 12] = "Validade";
                var rValidade10 = ws2.Cells[linhaAtual, 12];
                Negrito(rValidade10);
                Borda(rValidade10);
                Centro(rValidade10);

                ws2.Cells[linhaAtual, 22] = "Nome";
                var rNomeManobrador = ws2.Range[ws2.Cells[linhaAtual, 22], ws2.Cells[linhaAtual +1, 22]];
                rNomeManobrador.Merge();
                Negrito(rNomeManobrador);
                Borda(rNomeManobrador);
                Centro(rNomeManobrador);

                ws2.Cells[linhaAtual, 23] = "Habilitações";
                var rHabilitacoes = ws2.Range[ws2.Cells[linhaAtual, 23], ws2.Cells[linhaAtual, 24]];
                rHabilitacoes.Merge();
                Negrito(rHabilitacoes);
                Borda(rHabilitacoes);
                Centro(rHabilitacoes);


               //
                linhaAtual = linhaAtual + 1;
                ws2.Cells[linhaAtual, 1] = "N.º";
                var rNumEquipamento = ws2.Cells[linhaAtual, 1];
                Centro(rNumEquipamento);
                Borda(rNumEquipamento);

                ws2.Cells[linhaAtual, 2] = "Marca/ Modelo";
                var rMarcaModelo = ws2.Cells[linhaAtual, 2];
                Centro(rMarcaModelo);
                Borda(rMarcaModelo);

                ws2.Cells[linhaAtual, 3] = "Tipo de Máquina"; // ocupa 2 colunas
                var rTipoMaquina = ws2.Range[ws2.Cells[linhaAtual, 3], ws2.Cells[linhaAtual, 4]];
                rTipoMaquina.Merge();
                Centro(rTipoMaquina);
                Borda(rTipoMaquina);

                ws2.Cells[linhaAtual, 5] = "Número de Série"; // ocupa 3 colunas
                var rNumSerie = ws2.Range[ws2.Cells[linhaAtual, 5], ws2.Cells[linhaAtual, 7]];
                rNumSerie.Merge();
                Centro(rNumSerie);
                Borda(rNumSerie);

                ws2.Cells[linhaAtual, 8] = "C ; N/C ; N/A";
                var rCNCNA11 = ws2.Cells[linhaAtual, 8];
                Centro(rCNCNA11);
                Borda(rCNCNA11);

                ws2.Cells[linhaAtual, 9] = "C ; N/C ; N/A";
                var rCNCNA12 = ws2.Cells[linhaAtual, 9];
                Centro(rCNCNA12);
                Borda(rCNCNA12);

                ws2.Cells[linhaAtual, 10] = "C ; N/C ; N/A";
                var rCNCNA13 = ws2.Cells[linhaAtual, 10];
                Centro(rCNCNA13);
                Borda(rCNCNA13);

                ws2.Cells[linhaAtual, 11] = "C ; N/C ; N/A";
                var rCNCNA14 = ws2.Cells[linhaAtual, 11];
                Centro(rCNCNA14);
                Borda(rCNCNA14);

                ws2.Cells[linhaAtual, 12] = "Data";
                var rData14 = ws2.Cells[linhaAtual, 12];
                Centro(rData14);
                Borda(rData14);

                ws2.Cells[linhaAtual, 13] = "C ; N/C ; N/A";
                var rCNCNA15 = ws2.Cells[linhaAtual, 13];
                Centro(rCNCNA15);
                Borda(rCNCNA15);

                ws2.Cells[linhaAtual, 14] = "N.º Horas (à entrada em Obra)";
                var rNumHoras = ws2.Cells[linhaAtual, 14];
                Centro(rNumHoras);
                Borda(rNumHoras);

                ws2.Cells[linhaAtual, 15] = "Validade"; // ocupa 2 colunas
                var rValidade11 = ws2.Range[ws2.Cells[linhaAtual, 15], ws2.Cells[linhaAtual, 16]];
                rValidade11.Merge();
                Centro(rValidade11);
                Borda(rValidade11);

                ws2.Cells[linhaAtual, 17] = "Seguradora"; // ocupa 2 colunas
                var rSeguradora = ws2.Range[ws2.Cells[linhaAtual, 17], ws2.Cells[linhaAtual, 18]];
                rSeguradora.Merge();
                Centro(rSeguradora);
                Borda(rSeguradora);

                ws2.Cells[linhaAtual, 19] = "N.º Apólice";
                var rNumApolice = ws2.Cells[linhaAtual, 19];
                Centro(rNumApolice);
                Borda(rNumApolice);

                ws2.Cells[linhaAtual, 20] = "C ; N/C ; N/A";
                var rCNCNA16 = ws2.Cells[linhaAtual, 20];
                Centro(rCNCNA16);
                Borda(rCNCNA16);

                ws2.Cells[linhaAtual, 21] = "Validade";
                var rValidade12 = ws2.Cells[linhaAtual, 21];
                Centro(rValidade12);
                Borda(rValidade12);

                ws2.Cells[linhaAtual, 23] = "Tipo";
                var rTipoManobrador = ws2.Cells[linhaAtual, 23];
                Centro(rTipoManobrador);
                Borda(rTipoManobrador);

                ws2.Cells[linhaAtual, 24] = "C ; N/C ; N/A";
                var rCNCNA17 = ws2.Cells[linhaAtual, 24];
                Centro(rCNCNA17);
                Borda(rCNCNA17);

                ws2.Cells[linhaAtual, 26] = "Data";
                var rData15 = ws2.Cells[linhaAtual, 26];
                Centro(rData15);
                Borda(rData15);

                ws2.Cells[linhaAtual, 27] = "Data";
                var rData16 = ws2.Cells[linhaAtual, 27];
                Centro(rData16);
                Borda(rData16);

                ws2.Cells[linhaAtual, 28] = "Sim / Não";
                var rSimNao3 = ws2.Cells[linhaAtual, 28];
                Centro(rSimNao3);
                Borda(rSimNao3);

                //Linhas dos equipamentos da JPA TODO
                var queryEquipamentos = $@"
                SELECT DISTINCT
                    fei.ClasseID,
                    gc.Descricao AS ClasseDescricao,
                    pc.*  
                FROM COP_Obras o
                JOIN COP_FichasEquipamento fe ON fe.ObraID = o.id
                JOIN COP_FichasEquipamentoItems fei ON fei.FichasEquipamentoID = fe.id
                JOIN Precos_Componente pc ON pc.ComponenteID = fei.ComponenteID
                LEFT JOIN Geral_Classe gc ON gc.ClasseID = fei.ClasseID
                WHERE o.Codigo = '{codigoObra}';
                ";
                var dadosEquipamentos = BSO.Consulta(queryEquipamentos);
                var numRegistosEquipamentos = dadosEquipamentos.NumLinhas();
                dadosEquipamentos.Inicio();
                int equipamentoLinhaAtual = linhaAtual + 1;
                for (int i = 0; i < numRegistosEquipamentos; i++)
                {
                    ws2.Cells[equipamentoLinhaAtual, 1] = (i + 1).ToString(); // N.º
                    ws2.Cells[equipamentoLinhaAtual, 2] = dadosEquipamentos.DaValor<string>("Desig"); // Marca/ Modelo
                    ws2.Cells[equipamentoLinhaAtual, 3] = dadosEquipamentos.DaValor<string>("ClasseDescricao"); // Tipo de Máquina
                    ws2.Cells[equipamentoLinhaAtual, 5] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 6] = "";
                    ws2.Cells[equipamentoLinhaAtual, 7] = "C";
                    ws2.Cells[equipamentoLinhaAtual, 8] = "C";
                    ws2.Cells[equipamentoLinhaAtual, 9] = "C"; 
                    ws2.Cells[equipamentoLinhaAtual, 10] = "C"; 
                    ws2.Cells[equipamentoLinhaAtual, 11] = "";
                    ws2.Cells[equipamentoLinhaAtual, 12] = "C";
                    ws2.Cells[equipamentoLinhaAtual, 13] = "";
                    ws2.Cells[equipamentoLinhaAtual, 14] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 15] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 16] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 17] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 18] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 19] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 20] = "";
                    ws2.Cells[equipamentoLinhaAtual, 21] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 22] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 23] = "";
                    ws2.Cells[equipamentoLinhaAtual, 24] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 25] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 26] = "";
                    ws2.Cells[equipamentoLinhaAtual, 27] = ""; 
                    ws2.Cells[equipamentoLinhaAtual, 28] = ""; 

                    // As restantes células ficam vazias para serem preenchidas manualmente

                    equipamentoLinhaAtual++;
                    dadosEquipamentos.Seguinte();
                }




                // Page setup
                ws2.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                ws2.PageSetup.LeftMargin = excelApp.InchesToPoints(0.4);
                ws2.PageSetup.RightMargin = excelApp.InchesToPoints(0.3);
                ws2.PageSetup.TopMargin = excelApp.InchesToPoints(0.5);
                ws2.PageSetup.BottomMargin = excelApp.InchesToPoints(0.5);
                ws2.PageSetup.Zoom = false;
                ws2.PageSetup.FitToPagesWide = 1;
                ws2.PageSetup.FitToPagesTall = false;

                CriarPaginasPorIds(workbook, excelApp, codigoObra, idsSelecionados, dadosObra, dadosDonoObra, DonoObra);

                }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao criar segunda folha Excel: " + ex.Message, "Erro",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (ws2 != null) Marshal.ReleaseComObject(ws2);
            }


        }

        private void CriarPaginasPorIds(Excel.Workbook workbook, Excel.Application excelApp, string codigoObra, List<string> idsSelecionados, StdBELista dadosObra, StdBELista dadosDonoObra, string DonoObra)
        {
            var index = 2;
            foreach (var id in idsSelecionados.Skip(1))
            {
                Excel.Worksheet ws2 = null;

                try
                {
                    var queryNomeEntidade = $"SELECT Nome FROM  Geral_Entidade WHERE ID = '{id}'";
                    var nomeEnti = BSO.Consulta(queryNomeEntidade).DaValor<string>("Nome");
                    ws2 = (Excel.Worksheet)workbook.Worksheets.Add();
                    ws2.Name = index.ToString();
                    index++;

                    // Helpers
                    int ToOle(System.Drawing.Color c) => ColorTranslator.ToOle(c);
                    void Borda(Excel.Range r) => r.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    void Negrito(Excel.Range r, bool v = true) => r.Font.Bold = v;
                    void Centro(Excel.Range r) { r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; r.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; }
                    void Wrap(Excel.Range r, bool v = true) => r.WrapText = v;
                    Excel.Range R(int l1, int c1, int l2, int c2) => ws2.Range[ws2.Cells[l1, c1], ws2.Cells[l2, c2]];

                    // Linha 1: Marca + Título
                    ws2.Cells[1, 1] = "JPA";
                    var rLogo = R(1, 1, 1, 1); Negrito(rLogo); rLogo.Font.Size = 14;

                    ws2.Cells[1, 4] = "Controlo de Documentos de Empresas, Trabalhadores e Máquinas/Equipamentos";
                    var rTitulo = R(1, 4, 1, 20); rTitulo.Merge(); Negrito(rTitulo); Centro(rTitulo);

                    ws2.Cells[3, 3] = "Identificação dos Intervenientes";
                    var rIdentificacao = ws2.Range[ws2.Cells[3, 3], ws2.Cells[3, 8]];
                    rIdentificacao.Merge();
                    rIdentificacao.Font.Bold = true;
                    rIdentificacao.Font.Size = 12;
                    rIdentificacao.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rIdentificacao.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    // Fundo cinzento
                    rIdentificacao.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    // Bordas
                    rIdentificacao.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rIdentificacao.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Linha 4
                    ws2.Cells[4, 3] = "Designação da Empreitada:";
                    var rDesignacao = ws2.Range[ws2.Cells[4, 4], ws2.Cells[4, 8]];
                    rDesignacao.Merge();
                    rDesignacao.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rDesignacao.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    ws2.Cells[4, 4] = $"{codigoObra}";
                    // Linha 5
                    ws2.Cells[5, 3] = "Dono de Obra:";
                    var rDonoObra = ws2.Range[ws2.Cells[5, 4], ws2.Cells[5, 8]];
                    rDonoObra.Merge();
                    rDonoObra.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rDonoObra.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    ws2.Cells[5, 4] = DonoObra;

                    // Linha 6
                    ws2.Cells[6, 3] = "Entidade Executante:";
                    var rEntidade = ws2.Range[ws2.Cells[6, 4], ws2.Cells[6, 8]];
                    rEntidade.Merge();
                    rEntidade.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rEntidade.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    ws2.Cells[6, 4] = DonoObra;

                    ws2.Range[ws2.Cells[4, 3], ws2.Cells[6, 3]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    ws2.Range[ws2.Cells[4, 3], ws2.Cells[6, 3]].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    //---------------------
                    ws2.Cells[3, 12] = "Elaborado por (Téc. QAS):";
                    var rElaborado = ws2.Range[ws2.Cells[3, 12], ws2.Cells[3, 17]];
                    rElaborado.Merge();
                    rElaborado.Font.Bold = true;
                    rElaborado.Font.Size = 12;
                    rElaborado.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rElaborado.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    // Fundo cinzento
                    rElaborado.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    // Bordas
                    rElaborado.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rElaborado.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Linha 4
                    ws2.Cells[4, 12] = "Nome:";
                    var rNome = ws2.Range[ws2.Cells[4, 13], ws2.Cells[4, 17]];
                    rNome.Merge();
                    rNome.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rNome.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Linha 5
                    ws2.Cells[5, 12] = "Assinatura:";
                    var rAssinaura = ws2.Range[ws2.Cells[5, 13], ws2.Cells[5, 17]];
                    rAssinaura.Merge();
                    rAssinaura.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rAssinaura.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Linha 6
                    ws2.Cells[6, 12] = "Data:";
                    var rData1 = ws2.Range[ws2.Cells[6, 13], ws2.Cells[6, 17]];
                    rData1.Merge();
                    rData1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rData1.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    ws2.Cells[6, 13] = DateTime.Now.ToString("d 'de' MMMM 'de' yyyy", new System.Globalization.CultureInfo("pt-PT"));


                    ws2.Range[ws2.Cells[4, 12], ws2.Cells[6, 12]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    ws2.Range[ws2.Cells[4, 12], ws2.Cells[6, 12]].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    //---------------------------------
                    ws2.Cells[3, 20] = "Verificado por (Dir. Técnico / Dir. Obra):";
                    var rVerificado = ws2.Range[ws2.Cells[3, 20], ws2.Cells[3, 23]];
                    rVerificado.Merge();
                    rVerificado.Font.Bold = true;
                    rVerificado.Font.Size = 12;
                    rVerificado.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rVerificado.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    // Fundo cinzento
                    rVerificado.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    // Bordas
                    rVerificado.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rVerificado.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Linha 4
                    ws2.Cells[4, 20] = "Nome:";
                    var rNome2 = ws2.Range[ws2.Cells[4, 21], ws2.Cells[4, 23]];
                    rNome2.Merge();
                    rNome2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rNome2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Linha 5
                    ws2.Cells[5, 20] = "Assinatura:";
                    var rAssinaura2 = ws2.Range[ws2.Cells[5, 21], ws2.Cells[5, 23]];
                    rAssinaura2.Merge();
                    rAssinaura2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rAssinaura2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Linha 6
                    ws2.Cells[6, 20] = "Data:";
                    var rData2 = ws2.Range[ws2.Cells[6, 21], ws2.Cells[6, 23]];
                    rData2.Merge();
                    rData2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rData2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    ws2.Cells[6, 21] = DateTime.Now.ToString("d 'de' MMMM 'de' yyyy", new System.Globalization.CultureInfo("pt-PT"));


                    ws2.Range[ws2.Cells[4, 20], ws2.Cells[6, 20]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    ws2.Range[ws2.Cells[4, 20], ws2.Cells[6, 20]].Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    //----------------------------------------------------------------

                    // Linha 8
                    // Cabeçalho
                    ws2.Cells[8, 2] = "Atividade desenvolvida no Estaleiro";
                    var rAtividade = ws2.Range[ws2.Cells[8, 2], ws2.Cells[8, 3]];
                    rAtividade.Merge();
                    rAtividade.Font.Bold = true;
                    rAtividade.Font.Size = 12;
                    rAtividade.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    rAtividade.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rAtividade.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rAtividade.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rAtividade.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    // Linha 9
                    ws2.Cells[9, 2] = "";
                    var rEmpreiteiro = ws2.Range[ws2.Cells[9, 2], ws2.Cells[9, 3]];
                    rEmpreiteiro.Merge();
                    rEmpreiteiro.Font.Bold = true;
                    rEmpreiteiro.Font.Size = 12;
                    rEmpreiteiro.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rEmpreiteiro.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    Centro(rEmpreiteiro);


                    // Linha 10

                    ws2.Cells[10, 6] = "Pessoa(s) p/ contacto:";
                    var rContacto1 = ws2.Range[ws2.Cells[10, 7], ws2.Cells[10, 9]];
                    rContacto1.Merge();

                    // Apenas a borda inferior
                    rContacto1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rContacto1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    ws2.Cells[10, 10] = "Telf:";

                    var rTelf1 = ws2.Range[ws2.Cells[10, 11], ws2.Cells[10, 13]];
                    rTelf1.Merge();

                    // Apenas a borda inferior
                    rTelf1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rTelf1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    ws2.Cells[10, 14] = "E-mail:";

                    var rEmail1 = ws2.Range[ws2.Cells[10, 15], ws2.Cells[10, 17]];
                    rEmail1.Merge();

                    // Apenas a borda inferior
                    rEmail1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rEmail1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    ws2.Cells[10, 19] = "Função:";

                    var rFuncao1 = ws2.Range[ws2.Cells[10, 20], ws2.Cells[10, 22]];
                    rFuncao1.Merge();

                    // Apenas a borda inferior
                    rFuncao1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rFuncao1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;


                    // Linha 11

                    ws2.Cells[11, 6] = "Pessoa(s) p/ contacto:";
                    var rContacto2 = ws2.Range[ws2.Cells[11, 7], ws2.Cells[11, 9]];
                    rContacto2.Merge();

                    // Apenas a borda inferior
                    rContacto2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rContacto2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    ws2.Cells[11, 10] = "Telf:";

                    var rTelf2 = ws2.Range[ws2.Cells[11, 11], ws2.Cells[11, 13]];
                    rTelf2.Merge();

                    // Apenas a borda inferior
                    rTelf2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rTelf2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    ws2.Cells[11, 14] = "E-mail:";

                    var rEmail2 = ws2.Range[ws2.Cells[11, 15], ws2.Cells[11, 17]];
                    rEmail2.Merge();

                    // Apenas a borda inferior
                    rEmail2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rEmail2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    ws2.Cells[11, 19] = "Função:";

                    var rFuncao2 = ws2.Range[ws2.Cells[11, 20], ws2.Cells[11, 22]];
                    rFuncao2.Merge();

                    // Apenas a borda inferior
                    rFuncao2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rFuncao2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight =
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;



                    // Linha 13 e 14
                    ws2.Cells[13, 1] = "EMPRESA";
                    // Mescla linha 13 e 14, colunas 1 a 5
                    var rEmpresa = ws2.Range[ws2.Cells[13, 1], ws2.Cells[14, 5]];
                    rEmpresa.Merge();
                    rEmpresa.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rEmpresa.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rEmpresa.Font.Bold = true;
                    rEmpresa.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rEmpresa.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    rEmpresa.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    //Linha 13

                    // Numeração 1..23 (col 5..23)
                    int lastCol = 28;
                    int numRow = 13;

                    for (int c = 6, n = 1; c <= lastCol && n <= 23; n++)
                    {
                        Microsoft.Office.Interop.Excel.Range r;

                        // Se o número precisa ocupar 2 colunas (2 ou 9)
                        if (n == 2 || n == 9)
                        {
                            r = ws2.Range[ws2.Cells[numRow, c], ws2.Cells[numRow, c + 1]];
                            r.Merge();
                            ws2.Cells[numRow, c] = n.ToString();
                            c += 2; // avança duas colunas
                        }
                        else
                        {
                            // Célula normal
                            r = ws2.Range[ws2.Cells[numRow, c], ws2.Cells[numRow, c]];
                            ws2.Cells[numRow, c] = n.ToString();
                            c++; // avança uma coluna
                        }

                        // Estilos
                        Negrito(r);
                        Centro(r);
                        Borda(r);
                        r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                    }

                    // Linha 14

                    ws2.Cells[14, 6] = "Contribuinte";
                    var rContribuinte = ws2.Cells[14, 6]; // já é um Range
                    Negrito(rContribuinte);
                    Borda(rContribuinte);
                    Centro(rContribuinte);

                    ws2.Cells[14, 7] = "Alvará / Certificado";
                    var rAlvara = ws2.Range[ws2.Cells[14, 7], ws2.Cells[14, 8]];
                    rAlvara.Merge();
                    Negrito(rAlvara);
                    Borda(rAlvara);
                    Centro(rAlvara);

                    ws2.Cells[14, 9] = "Anexo D";
                    var rAnexo = ws2.Cells[14, 9];
                    Negrito(rAnexo);
                    Borda(rAnexo);
                    Centro(rAnexo);

                    ws2.Cells[14, 10] = "Cert. ND Finanças";
                    var rNDFinanças = ws2.Cells[14, 10];
                    Negrito(rNDFinanças);
                    Borda(rNDFinanças);
                    Centro(rNDFinanças);

                    ws2.Cells[14, 11] = "Decl. ND Seg. Social";
                    var rDeclSegSocial = ws2.Cells[14, 11];
                    Negrito(rDeclSegSocial);
                    Borda(rDeclSegSocial);
                    Centro(rDeclSegSocial);

                    ws2.Cells[14, 12] = "Folha Pag. Seg. Social";
                    var rFolhaPagSegSocial = ws2.Cells[14, 12];
                    Negrito(rFolhaPagSegSocial);
                    Borda(rFolhaPagSegSocial);
                    Centro(rFolhaPagSegSocial);

                    ws2.Cells[14, 13] = "Recibo de Pag. Seg. Social";
                    var rRecibodePagSegSocial = ws2.Cells[14, 13];
                    Negrito(rRecibodePagSegSocial);
                    Borda(rRecibodePagSegSocial);
                    Centro(rRecibodePagSegSocial);

                    ws2.Cells[14, 14] = "Apólice AT";
                    var rApoliceAT = ws2.Cells[14, 14];
                    Negrito(rApoliceAT);
                    Borda(rApoliceAT);
                    Centro(rApoliceAT);

                    ws2.Cells[14, 15] = "Modalidade do Seguro";
                    var rModalidadeSeguro = ws2.Range[ws2.Cells[14, 15], ws2.Cells[14, 16]];
                    rModalidadeSeguro.Merge();
                    Negrito(rModalidadeSeguro);
                    Borda(rModalidadeSeguro);
                    Centro(rModalidadeSeguro);

                    ws2.Cells[14, 17] = "Recibo Apólice AT";
                    var rReciboApoliceAT = ws2.Cells[14, 17];
                    Negrito(rReciboApoliceAT);
                    Borda(rReciboApoliceAT);
                    Centro(rReciboApoliceAT);

                    ws2.Cells[14, 18] = "Apólice RC";
                    var rApoliceRC = ws2.Cells[14, 18];
                    Negrito(rApoliceRC);
                    Borda(rApoliceRC);
                    Centro(rApoliceRC);

                    ws2.Cells[14, 19] = "Recibo RC";
                    var rReciboRC = ws2.Cells[14, 19];
                    Negrito(rReciboRC);
                    Borda(rReciboRC);
                    Centro(rReciboRC);

                    ws2.Cells[14, 20] = "Registo(s) Criminal(ais)";
                    var rRegistoCriminal = ws2.Cells[14, 20];
                    Negrito(rRegistoCriminal);
                    Borda(rRegistoCriminal);
                    Centro(rRegistoCriminal);

                    ws2.Cells[14, 21] = "Horário de Trabalho";
                    var rHorarioTrabalho = ws2.Cells[14, 21];
                    Negrito(rHorarioTrabalho);
                    Borda(rHorarioTrabalho);
                    Centro(rHorarioTrabalho);

                    ws2.Cells[14, 22] = "Dec.  Trab. Imigr.";
                    var rDecTrabImigr = ws2.Cells[14, 22];
                    Negrito(rDecTrabImigr);
                    Borda(rDecTrabImigr);
                    Centro(rDecTrabImigr);

                    ws2.Cells[14, 23] = "Dec. Resp. Estaleiro";
                    var rDecRespEstaleiro = ws2.Cells[14, 23];
                    Negrito(rDecRespEstaleiro);
                    Borda(rDecRespEstaleiro);
                    Centro(rDecRespEstaleiro);

                    ws2.Cells[14, 24] = "Dec. Ades. PSS";
                    var rDecAdesPSS = ws2.Cells[14, 24];
                    Negrito(rDecAdesPSS);
                    Borda(rDecAdesPSS);
                    Centro(rDecAdesPSS);

                    ws2.Cells[14, 25] = "Contrato Subempreitada";
                    var rContratoSubempreitada = ws2.Cells[14, 25];
                    Negrito(rContratoSubempreitada);
                    Borda(rContratoSubempreitada);
                    Centro(rContratoSubempreitada);

                    ws2.Cells[14, 26] = "Entrada em Obra";
                    var rEntradaObra = ws2.Cells[14, 26];
                    Negrito(rEntradaObra);
                    Borda(rEntradaObra);
                    Centro(rEntradaObra);

                    ws2.Cells[14, 27] = "Saída de Obra";
                    var rSaidaObra = ws2.Cells[14, 27];
                    Negrito(rSaidaObra);
                    Borda(rSaidaObra);
                    Centro(rSaidaObra);

                    ws2.Cells[14, 28] = "Autorização de Entrada";
                    var rAutorizacaoEntrada = ws2.Cells[14, 28];
                    Negrito(rAutorizacaoEntrada);
                    Borda(rAutorizacaoEntrada);
                    Centro(rAutorizacaoEntrada);

                    // Linha 15

                    ws2.Cells[15, 1] = "N.º";
                    var rNum = ws2.Cells[15, 1];
                    Centro(rNum);
                    Borda(rNum);

                    ws2.Cells[15, 2] = "Designação Social";
                    var rDesignacaoSocial = ws2.Cells[15, 2];
                    Centro(rDesignacaoSocial);
                    Borda(rDesignacaoSocial);

                    ws2.Cells[15, 3] = "Sede";
                    var rSede = ws2.Range[ws2.Cells[15, 3], ws2.Cells[15, 5]];
                    rSede.Merge();
                    Centro(rSede);
                    Borda(rSede);

                    ws2.Cells[15, 6] = "N.º";
                    var rNum2 = ws2.Cells[15, 6];
                    Centro(rNum2);
                    Borda(rNum2);

                    ws2.Cells[15, 7] = "N.º";
                    var rNum3 = ws2.Cells[15, 7];
                    Centro(rNum3);
                    Borda(rNum3);

                    ws2.Cells[15, 8] = "PUB / PAR";
                    var rPubPar = ws2.Cells[15, 8];
                    Centro(rPubPar);
                    Borda(rPubPar);

                    ws2.Cells[15, 9] = "C ; N/C ; N/A";
                    var rCNCNA = ws2.Cells[15, 9];
                    Centro(rCNCNA);
                    Borda(rCNCNA);

                    ws2.Cells[15, 10] = "Validade";
                    var rValidade = ws2.Cells[15, 10];
                    Centro(rValidade);
                    Borda(rValidade);

                    ws2.Cells[15, 11] = "Validade";
                    var rValidade2 = ws2.Cells[15, 11];
                    Centro(rValidade2);
                    Borda(rValidade2);

                    ws2.Cells[15, 12] = "Validade";
                    var rValidade3 = ws2.Cells[15, 12];
                    Centro(rValidade3);
                    Borda(rValidade3);

                    ws2.Cells[15, 13] = "C ; N/C ; N/A";
                    var rCNCNA2 = ws2.Cells[15, 13];
                    Centro(rCNCNA2);
                    Borda(rCNCNA2);

                    ws2.Cells[15, 14] = "N.º";
                    var rNum4 = ws2.Cells[15, 14];
                    Centro(rNum4);
                    Borda(rNum4);

                    ws2.Cells[15, 15] = "Fixo?";
                    var rFixo = ws2.Cells[15, 15];
                    Centro(rFixo);
                    Borda(rFixo);

                    ws2.Cells[15, 16] = "Prémio Variável?";
                    var rPremioVariavel = ws2.Cells[15, 16];
                    Centro(rPremioVariavel);
                    Borda(rPremioVariavel);

                    ws2.Cells[15, 17] = "Validade";
                    var rValidade4 = ws2.Cells[15, 17];
                    Centro(rValidade4);
                    Borda(rValidade4);

                    ws2.Cells[15, 18] = "N.º";
                    var rNum5 = ws2.Cells[15, 18];
                    Centro(rNum5);
                    Borda(rNum5);

                    ws2.Cells[15, 19] = "Validade";
                    var rValidade5 = ws2.Cells[15, 19];
                    Centro(rValidade5);
                    Borda(rValidade5);

                    ws2.Cells[15, 20] = "C ; N/C ; N/A";
                    var rCNCNA3 = ws2.Cells[15, 20];
                    Centro(rCNCNA3);
                    Borda(rCNCNA3);

                    ws2.Cells[15, 21] = "C ; N/C ; N/A";
                    var rCNCNA4 = ws2.Cells[15, 21];
                    Centro(rCNCNA4);
                    Borda(rCNCNA4);

                    ws2.Cells[15, 22] = "C ; N/C ; N/A";
                    var rCNCNA5 = ws2.Cells[15, 22];
                    Centro(rCNCNA5);
                    Borda(rCNCNA5);

                    ws2.Cells[15, 23] = "C ; N/C ; N/A";
                    var rCNCNA6 = ws2.Cells[15, 23];
                    Centro(rCNCNA6);
                    Borda(rCNCNA6);

                    ws2.Cells[15, 24] = "C ; N/C ; N/A";
                    var rCNCNA7 = ws2.Cells[15, 24];
                    Centro(rCNCNA7);
                    Borda(rCNCNA7);

                    ws2.Cells[15, 25] = "Com:";
                    var rCom = ws2.Cells[15, 25];
                    Centro(rCom);
                    Borda(rCom);

                    ws2.Cells[15, 26] = "Data";
                    var rData3 = ws2.Cells[15, 26];
                    Centro(rData3);
                    Borda(rData3);

                    ws2.Cells[15, 27] = "Data";
                    var rData4 = ws2.Cells[15, 27];
                    Centro(rData4);
                    Borda(rData4);

                    ws2.Cells[15, 28] = "Sim / Não";
                    var rSimNao = ws2.Cells[15, 28];
                    Centro(rSimNao);
                    Borda(rSimNao);

                    // Linha 16 Dados 

                    var querydadosEntidade = $"SELECT * FROM  Geral_Entidade WHERE ID = '{id}' ";
                    var dadosEntidade = BSO.Consulta(querydadosEntidade);



                    ws2.Cells[16, 1] = "1";
                    var rNumJPA = ws2.Cells[16, 1];
                    Centro(rNumJPA);

                    ws2.Cells[16, 2] = dadosEntidade.DaValor<string>("Nome");

                    ws2.Cells[16, 3] = dadosEntidade.DaValor<string>("Morada");
                    var rSedeJPA = ws2.Range[ws2.Cells[16, 3], ws2.Cells[16, 5]];
                    rSedeJPA.Merge();

                    ws2.Cells[16, 6] = dadosEntidade.DaValor<string>("NIPC");
                    ws2.Cells[16, 7] = dadosEntidade.DaValor<string>("AlvaraNumero");
                    ws2.Cells[16, 8] = "PAR";


                    var valor = dadosEntidade.DaValor<string>("CDU_AnexoAnexoD");
                    ws2.Cells[16, 9] = !string.IsNullOrEmpty(valor) ? "C" : "N/C";
                 
                    // Lista de colunas e campos correspondentes
                    var colunas = new int[] { 10, 11, 12 };
                    var campos = new string[] { "CDU_validadeFinancas", "CDU_ValidadeSegSocial", "CDU_ValidadeFolhaPag" };

                    for (int i = 0; i < colunas.Length; i++)
                    {
                        var valorStr = dadosEntidade.DaValor<string>(campos[i]);

                        if (DateTime.TryParse(valorStr, out DateTime data))
                        {
                            ws2.Cells[16, colunas[i]] = data.ToString("dd-MM-yyyy");

                            // Se a data for anterior a hoje, pinta de vermelho
                            if (data < DateTime.Today)
                            {
                                ws2.Cells[16, colunas[i]].Interior.Color = System.Drawing.Color.Red;
                            }
                        }
                        else
                        {
                            ws2.Cells[16, colunas[i]] = "";
                        }
                    }
                    var valor2 = dadosEntidade.DaValor<string>("CDU_ValidadeComprovativoPagamento");
                    ws2.Cells[16, 13] = !string.IsNullOrEmpty(valor2) ? "C" : "N/C"; //dadosEntidade.DaValor<string>("CDU_ValidadeComprovativoPagamento");  se tiver preenchido colo

                    ws2.Cells[16, 14] = "";
                    ws2.Cells[16, 15] = "";
                    ws2.Cells[16, 16] = "";
                    // Lista de novas colunas e campos correspondentes
                    var novasColunas = new int[] { 17, 19 };
                    var novosCampos = new string[] { "CDU_ValidadeReciboSeguroAT", "CDU_ValidadeSeguroRC" };

                    for (int i = 0; i < novasColunas.Length; i++)
                    {
                        var valorStr = dadosEntidade.DaValor<string>(novosCampos[i]);

                        if (DateTime.TryParse(valorStr, out DateTime data))
                        {
                            ws2.Cells[16, novasColunas[i]] = data.ToString("dd-MM-yyyy");

                            // Se a data for anterior a hoje, pinta de vermelho
                            if (data < DateTime.Today)
                            {
                                ws2.Cells[16, novasColunas[i]].Interior.Color = System.Drawing.Color.Red;
                            }
                        }
                        else
                        {
                            ws2.Cells[16, novasColunas[i]] = "";
                        }
                    }

                    ws2.Cells[16, 18] = "TODO RC";
                    ws2.Cells[16, 20] = "TODO";
                    var querydadosAutorizacoesEntidades = $"SELECT * fROM TDU_AD_Autorizacoes WHERE ID_Entidade = '{id}'";
                    var dadosAutorizacoesEntidades = BSO.Consulta(querydadosAutorizacoesEntidades);
                    var validadeCaminho2 = dadosAutorizacoesEntidades.DaValor<string>("caminho2");
                    var resultado = !string.IsNullOrEmpty(validadeCaminho2) ? "C" : "NC";
                    ws2.Cells[16, 21] = resultado;
                    var validadeCaminho5 = dadosAutorizacoesEntidades.DaValor<string>("caminho5");
                    var resultado2 = !string.IsNullOrEmpty(validadeCaminho5) ? "C" : "NC";
                    ws2.Cells[16, 22] = resultado2;
                    var validadeCaminho4 = dadosAutorizacoesEntidades.DaValor<string>("caminho4");
                    var resultado3 = !string.IsNullOrEmpty(validadeCaminho4) ? "C" : "NC";
                    ws2.Cells[16, 23] = resultado3;
                    var validadeCaminho3 = dadosAutorizacoesEntidades.DaValor<string>("caminho3");
                    var resultado4 = !string.IsNullOrEmpty(validadeCaminho3) ? "C" : "NC";
                    ws2.Cells[16, 24] = resultado4;

                    ws2.Cells[16, 25] = "";
                    var dataEntradaStr = dadosAutorizacoesEntidades.DaValor<string>("Data_Entrada");

                    if (DateTime.TryParse(dataEntradaStr, out DateTime dataEntrada))
                    {
                        ws2.Cells[16, 26] = dataEntrada.ToString("dd-MM-yyyy");
                    }
                    else
                    {
                        ws2.Cells[16, 26] = ""; // Ou algum valor padrão se a data for inválida
                    }
                    var dataSaidaStr = dadosAutorizacoesEntidades.DaValor<string>("Data_Saida");

                    if (DateTime.TryParse(dataSaidaStr, out DateTime dataSaida) && dataSaida != new DateTime(1900, 1, 1))
                    {
                        ws2.Cells[16, 27] = dataSaida.ToString("dd-MM-yyyy");
                    }
                    else
                    {
                        ws2.Cells[16, 27] = ""; // Vazio se for null, inválido ou 1900-01-01
                    }

                    ws2.Cells[16, 28] = "Sim";


                    // Linha 18 a 20
                    ws2.Cells[18, 1] = "TRABALHADORES";
                    var rTRABALHADORES = ws2.Range[ws2.Cells[18, 1], ws2.Cells[20, 4]];
                    rTRABALHADORES.Merge();
                    rTRABALHADORES.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rTRABALHADORES.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    rTRABALHADORES.Font.Bold = true;
                    rTRABALHADORES.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rTRABALHADORES.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    rTRABALHADORES.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    // Linha 18

                    // Numeração 1..23 (col 5..23)
                    int lastCol2 = 28;
                    int numRow2 = 18;

                    for (int c = 5, n = 1; c <= lastCol2 && n <= 23; n++)
                    {
                        Microsoft.Office.Interop.Excel.Range r;

                        if (n == 5 || n == 12 || n == 6 || n == 13 || n == 9)
                        {
                            r = ws2.Range[ws2.Cells[numRow2, c], ws2.Cells[numRow2, c + 1]];
                            r.Merge();
                            ws2.Cells[numRow2, c] = n.ToString();
                            c += 2; // avança duas colunas
                        }
                        else if (n == 10)// n== 10 ocupa 3 colunas
                        {
                            r = ws2.Range[ws2.Cells[numRow2, c], ws2.Cells[numRow2, c + 2]];
                            r.Merge();
                            ws2.Cells[numRow2, c] = n.ToString();
                            c += 3; // avança três colunas
                        }
                        else if (n == 11)// n== 11 ocupa 4 colunas 
                        {
                            r = ws2.Range[ws2.Cells[numRow2, c], ws2.Cells[numRow2, c + 3]];
                            r.Merge();
                            ws2.Cells[numRow2, c] = n.ToString();
                            c += 4; // avança quatro colunas
                        }
                        else
                        {
                            // Célula normal
                            r = ws2.Range[ws2.Cells[numRow2, c], ws2.Cells[numRow2, c]];
                            ws2.Cells[numRow2, c] = n.ToString();
                            c++; // avança uma coluna
                        }

                        // Estilos
                        Negrito(r);
                        Centro(r);
                        Borda(r);
                        r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                    }

                    // Linha 19

                    ws2.Cells[19, 5] = "Categoria / Função";
                    var rCategoriaFuncao = ws2.Range[ws2.Cells[19, 5], ws2.Cells[21, 5]]; // ocupa 3 linhas
                    rCategoriaFuncao.Merge();
                    Negrito(rCategoriaFuncao);
                    Borda(rCategoriaFuncao);
                    Centro(rCategoriaFuncao);


                    ws2.Cells[19, 6] = "CAP (se aplicável)";
                    var rCAP = ws2.Range[ws2.Cells[19, 6], ws2.Cells[20, 6]];
                    rCAP.Merge();
                    Negrito(rCAP);
                    Borda(rCAP);
                    Centro(rCAP);

                    ws2.Cells[19, 7] = "Contribuinte";
                    var rContribuinte2 = ws2.Range[ws2.Cells[19, 7], ws2.Cells[20, 7]];
                    rContribuinte2.Merge();
                    Negrito(rContribuinte2);
                    Borda(rContribuinte2);
                    Centro(rContribuinte2);

                    ws2.Cells[19, 8] = "Segurança Social";
                    var rSegurancaSocial = ws2.Range[ws2.Cells[19, 8], ws2.Cells[20, 8]];
                    rSegurancaSocial.Merge();
                    Negrito(rSegurancaSocial);
                    Borda(rSegurancaSocial);
                    Centro(rSegurancaSocial);

                    ws2.Cells[19, 9] = "Cartão de Cidadão";
                    var rFichaAptidaoMedica = ws2.Range[ws2.Cells[19, 9], ws2.Cells[20, 10]];
                    rFichaAptidaoMedica.Merge();
                    Negrito(rFichaAptidaoMedica);
                    Borda(rFichaAptidaoMedica);
                    Centro(rFichaAptidaoMedica);

                    ws2.Cells[19, 11] = "Ficha de Aptidão Médica";
                    var rFichaDistribuicaoEPI = ws2.Range[ws2.Cells[19, 11], ws2.Cells[19, 12]];
                    rFichaDistribuicaoEPI.Merge();
                    Negrito(rFichaDistribuicaoEPI);
                    Borda(rFichaDistribuicaoEPI);
                    Centro(rFichaDistribuicaoEPI);

                    ws2.Cells[19, 13] = "Ficha de Distribuição de EPI";
                    var rConstaMapaSS = ws2.Range[ws2.Cells[19, 13], ws2.Cells[20, 13]];
                    rConstaMapaSS.Merge();
                    Negrito(rConstaMapaSS);
                    Borda(rConstaMapaSS);
                    Centro(rConstaMapaSS);

                    ws2.Cells[19, 14] = "Consta no Mapa  SS / Inscrito?";
                    var rValidade6 = ws2.Range[ws2.Cells[19, 14], ws2.Cells[20, 14]];
                    rValidade6.Merge();
                    Negrito(rValidade6);
                    Borda(rValidade6);
                    Centro(rValidade6);

                    ws2.Cells[19, 15] = "Admissão na SS (caso não conste no Mapa da SS)";
                    var rTrabalhadorEstrangeiro = ws2.Range[ws2.Cells[19, 15], ws2.Cells[20, 16]];
                    rTrabalhadorEstrangeiro.Merge();
                    Negrito(rTrabalhadorEstrangeiro);
                    Borda(rTrabalhadorEstrangeiro);
                    Centro(rTrabalhadorEstrangeiro);

                    ws2.Cells[19, 17] = "Trabalhador Estrangeiro";
                    var rTrabalhadorEstrangeiro2 = ws2.Range[ws2.Cells[19, 17], ws2.Cells[19, 19]];
                    rTrabalhadorEstrangeiro2.Merge();
                    Negrito(rTrabalhadorEstrangeiro2);
                    Borda(rTrabalhadorEstrangeiro2);
                    Centro(rTrabalhadorEstrangeiro2);

                    ws2.Cells[19, 20] = "Formação / Informação";
                    var rFormacaoInformacao = ws2.Range[ws2.Cells[19, 20], ws2.Cells[19, 23]];
                    rFormacaoInformacao.Merge();
                    Negrito(rFormacaoInformacao);
                    Borda(rFormacaoInformacao);
                    Centro(rFormacaoInformacao);

                    ws2.Cells[19, 24] = "Cadastro";
                    var rCadastro = ws2.Range[ws2.Cells[19, 24], ws2.Cells[19, 25]];
                    rCadastro.Merge();
                    Negrito(rCadastro);
                    Borda(rCadastro);
                    Centro(rCadastro);

                    ws2.Cells[19, 26] = "Entrada Obra";
                    var rEntradaObra2 = ws2.Range[ws2.Cells[19, 26], ws2.Cells[20, 26]];
                    rEntradaObra2.Merge();
                    Negrito(rEntradaObra2);
                    Borda(rEntradaObra2);
                    Centro(rEntradaObra2);

                    ws2.Cells[19, 27] = "Saída Obra";
                    var rSaidaObra2 = ws2.Range[ws2.Cells[19, 27], ws2.Cells[20, 27]];
                    rSaidaObra2.Merge();
                    Negrito(rSaidaObra2);
                    Borda(rSaidaObra2);
                    Centro(rSaidaObra2);

                    ws2.Cells[19, 28] = "Autorização de Entrada em Obra";
                    var rAutorizacaoEntrada2 = ws2.Range[ws2.Cells[19, 28], ws2.Cells[20, 28]];
                    rAutorizacaoEntrada2.Merge();
                    Negrito(rAutorizacaoEntrada2);
                    Borda(rAutorizacaoEntrada2);
                    Centro(rAutorizacaoEntrada2);


                    // LInha 20
                    ws2.Cells[20, 11] = "Conforme Cat. Prof.?";
                    var rConformeCatProf = ws2.Cells[20, 11];
                    Negrito(rConformeCatProf);
                    Borda(rConformeCatProf);
                    Centro(rConformeCatProf);

                    ws2.Cells[20, 12] = "Validade";
                    var rValidade7 = ws2.Range[ws2.Cells[20, 12], ws2.Cells[21, 12]];
                    rValidade7.Merge();
                    Negrito(rValidade7);
                    Borda(rValidade7);
                    Centro(rValidade7);


                    ws2.Cells[20, 17] = "Passaporte c/visto / Titulo de Residência";
                    var rPassaporteVisto = ws2.Range[ws2.Cells[20, 17], ws2.Cells[20, 19]];
                    rPassaporteVisto.Merge();
                    Negrito(rPassaporteVisto);
                    Borda(rPassaporteVisto);
                    Centro(rPassaporteVisto);

                    ws2.Cells[20, 20] = "Acolhimento";
                    var rAcolhimento = ws2.Cells[20, 20];
                    Negrito(rAcolhimento);
                    Borda(rAcolhimento);
                    Centro(rAcolhimento);

                    ws2.Cells[20, 21] = "Específica 1";
                    var rEspecifica1 = ws2.Cells[20, 21];
                    Negrito(rEspecifica1);
                    Borda(rEspecifica1);
                    Centro(rEspecifica1);

                    ws2.Cells[20, 22] = "Específica 2";
                    var rEspecifica2 = ws2.Cells[20, 22];
                    Negrito(rEspecifica2);
                    Borda(rEspecifica2);
                    Centro(rEspecifica2);

                    ws2.Cells[20, 23] = "Específica 3";
                    var rEspecifica3 = ws2.Cells[20, 23];
                    Negrito(rEspecifica3);
                    Borda(rEspecifica3);
                    Centro(rEspecifica3);

                    ws2.Cells[20, 24] = "1.º Aviso";
                    var rPrimeiroAviso = ws2.Cells[20, 24];
                    Negrito(rPrimeiroAviso);
                    Borda(rPrimeiroAviso);
                    Centro(rPrimeiroAviso);

                    ws2.Cells[20, 25] = "2.º Aviso";
                    var rSegundoAviso = ws2.Cells[20, 25];
                    Negrito(rSegundoAviso);
                    Borda(rSegundoAviso);
                    Centro(rSegundoAviso);

                    // Linha 21

                    ws2.Cells[21, 1] = "N.º";
                    var rNumTrabalhador = ws2.Cells[21, 1];
                    Centro(rNumTrabalhador);
                    Borda(rNumTrabalhador);

                    ws2.Cells[21, 2] = "Nome Completo";
                    var rNomeCompleto = ws2.Cells[21, 2];
                    Centro(rNomeCompleto);
                    Borda(rNomeCompleto);

                    ws2.Cells[21, 3] = "Residência Habitual";
                    var rResidenciaHabitual = ws2.Cells[21, 3];
                    Centro(rResidenciaHabitual);
                    Borda(rResidenciaHabitual);

                    ws2.Cells[21, 4] = "Nacionalidade";
                    var rNacionalidade = ws2.Cells[21, 4];
                    Centro(rNacionalidade);
                    Borda(rNacionalidade);

                    ws2.Cells[21, 6] = "N.º";
                    var rNumCAP = ws2.Cells[21, 6];
                    Centro(rNumCAP);
                    Borda(rNumCAP);

                    ws2.Cells[21, 7] = "N.º";
                    var rNumContribuinte = ws2.Cells[21, 7];
                    Centro(rNumContribuinte);
                    Borda(rNumContribuinte);

                    ws2.Cells[21, 8] = "N.º";
                    var rNumSegSocial = ws2.Cells[21, 8];
                    Centro(rNumSegSocial);
                    Borda(rNumSegSocial);

                    ws2.Cells[21, 9] = "N.º";
                    var rNumCC = ws2.Cells[21, 9];
                    Centro(rNumCC);
                    Borda(rNumCC);

                    ws2.Cells[21, 10] = "Validade";
                    var rValidade8 = ws2.Cells[21, 10];
                    Centro(rValidade8);
                    Borda(rValidade8);

                    ws2.Cells[21, 11] = "C ; N/C ; N/A";
                    var rCNCNA8 = ws2.Cells[21, 11];
                    Centro(rCNCNA8);
                    Borda(rCNCNA8);

                    ws2.Cells[21, 13] = "C ; N/C";
                    var rCNCNA9 = ws2.Cells[21, 13];
                    Centro(rCNCNA9);
                    Borda(rCNCNA9);

                    ws2.Cells[21, 14] = "C ; N/C ; N/A";
                    var rCNCNA10 = ws2.Cells[21, 14];
                    Centro(rCNCNA10);
                    Borda(rCNCNA10);

                    ws2.Cells[21, 15] = "Data";
                    var rData5 = ws2.Range[ws2.Cells[21, 15], ws2.Cells[21, 16]];
                    rData5.Merge();
                    Centro(rData5);
                    Borda(rData5);

                    ws2.Cells[21, 17] = "Tipo de documento";
                    var rTipoDocumento = ws2.Cells[21, 17];
                    Centro(rTipoDocumento);
                    Borda(rTipoDocumento);

                    ws2.Cells[21, 18] = "Número";
                    var rNumeroDoc = ws2.Cells[21, 18];
                    Centro(rNumeroDoc);
                    Borda(rNumeroDoc);

                    ws2.Cells[21, 19] = "Validade";
                    var rValidade9 = ws2.Cells[21, 19];
                    Centro(rValidade9);
                    Borda(rValidade9);

                    ws2.Cells[21, 20] = "Data";
                    var rData6 = ws2.Cells[21, 20];
                    Centro(rData6);
                    Borda(rData6);

                    ws2.Cells[21, 21] = "Data";
                    var rData7 = ws2.Cells[21, 21];
                    Centro(rData7);
                    Borda(rData7);

                    ws2.Cells[21, 22] = "Data";
                    var rData8 = ws2.Cells[21, 22];
                    Centro(rData8);
                    Borda(rData8);

                    ws2.Cells[21, 23] = "Data";
                    var rData9 = ws2.Cells[21, 23];
                    Centro(rData9);
                    Borda(rData9);

                    ws2.Cells[21, 24] = "Data";
                    var rData10 = ws2.Cells[21, 24];
                    Centro(rData10);
                    Borda(rData10);

                    ws2.Cells[21, 25] = "Data";
                    var rData11 = ws2.Cells[21, 25];
                    Centro(rData11);
                    Borda(rData11);

                    ws2.Cells[21, 26] = "Data";
                    var rData12 = ws2.Cells[21, 26];
                    Centro(rData12);
                    Borda(rData12);

                    ws2.Cells[21, 27] = "Data";
                    var rData13 = ws2.Cells[21, 27];
                    Centro(rData13);
                    Borda(rData13);

                    ws2.Cells[21, 28] = "Sim / Não";
                    var rSimNao2 = ws2.Cells[21, 28];
                    Centro(rSimNao2);
                    Borda(rSimNao2);

                    //DADOS NAS LINHAS //********************************************************************** TRABALHADORES JPA

                    var queryTrabalhadoresEntidade = $@"SELECT * fROM TDU_AD_Trabalhadores WHERE id_empresa = '{id}';";
                    var dadosTrabalhadoresEntidades = BSO.Consulta(queryTrabalhadoresEntidade);

                    var numregistos = dadosTrabalhadoresEntidades.NumLinhas();

                    int linhaAtual = 22;
                    dadosTrabalhadoresEntidades.Inicio();
                    for (int i = 0; i < numregistos; i++)
                    {
                        
                        ws2.Cells[linhaAtual, 1] = (i + 1).ToString(); // N.º
                        ws2.Cells[linhaAtual, 2] = dadosTrabalhadoresEntidades.DaValor<string>("nome"); // Nome Completo
                        ws2.Cells[linhaAtual, 3] = "Todo"; // Residência Habitual
                        ws2.Cells[linhaAtual, 4] = "TODO"; // Nacionalidade
                                                          
                        ws2.Cells[linhaAtual, 6] = "TODO";                                                                              
                        ws2.Cells[linhaAtual, 7] = dadosTrabalhadoresEntidades.DaValor<string>("contribuinte"); // Contribuinte
                        ws2.Cells[linhaAtual, 8] = dadosTrabalhadoresEntidades.DaValor<string>("seguranca_social"); // Segurança Social
                        ws2.Cells[linhaAtual, 9] = "TODO";
                        string caminho1 = dadosTrabalhadoresEntidades.DaValor<string>("caminho1");

                        // Regex para pegar a data no formato dd/MM/yyyy
                        Match match = Regex.Match(caminho1, @"\d{2}/\d{2}/\d{4}");

                        if (match.Success)
                        {
                            ws2.Cells[linhaAtual, 10] = match.Value;
                        }
                        else
                        {
                            ws2.Cells[linhaAtual, 10] = "";
                        }
                        ws2.Cells[linhaAtual, 11] = "TODO"; 
                        ws2.Cells[linhaAtual, 12] = "TODO";

                        string valorCaminho5 = dadosTrabalhadoresEntidades.DaValor<string>("caminho5");

                        ws2.Cells[linhaAtual, 13] = string.IsNullOrWhiteSpace(valorCaminho5) ? "N/C" : "C";


                        ws2.Cells[linhaAtual, 14] = "TODO"; // Consta no Mapa  SS / Inscrito?
                        ws2.Cells[linhaAtual, 15] = "TODO"; // Admissão na SS (caso não conste no Mapa da SS)
                        ws2.Cells[linhaAtual, 16] = ""; // Admissão na SS (caso não conste no Mapa da SS)
                        ws2.Cells[linhaAtual, 17] = ""; // Passaporte c/visto / Titulo de Residência
                        ws2.Cells[linhaAtual, 18] = ""; // Validade
                        ws2.Cells[linhaAtual, 19] = ""; // Data
                        ws2.Cells[linhaAtual, 20] = ""; // Acolhimento
                        ws2.Cells[linhaAtual, 21] = ""; // Específica 1
                        ws2.Cells[linhaAtual, 22] = ""; // Específica 2
                        ws2.Cells[linhaAtual, 23] = ""; // Específica 3
                        ws2.Cells[linhaAtual, 24] = ""; // 1.º Aviso
                        ws2.Cells[linhaAtual, 25] = ""; // 2.º Aviso


                        var dataEntradaStr2 = dadosAutorizacoesEntidades.DaValor<string>("Data_Entrada");

                        if (DateTime.TryParse(dataEntradaStr2, out DateTime dataEntrada2))
                        {
                            ws2.Cells[linhaAtual, 26] = dataEntrada2.ToString("dd-MM-yyyy");
                        }
                        else
                        {
                            ws2.Cells[linhaAtual, 26] = ""; // Ou algum valor padrão se a data for inválida
                        }
                        var dataSaidaStr2 = dadosAutorizacoesEntidades.DaValor<string>("Data_Saida");

                        if (DateTime.TryParse(dataSaidaStr2, out DateTime dataSaida2) && dataSaida2 != new DateTime(1900, 1, 1))
                        {
                            ws2.Cells[linhaAtual, 27] = dataSaida2.ToString("dd-MM-yyyy");
                        }
                        else
                        {
                            ws2.Cells[linhaAtual, 27] = ""; // Vazio se for null, inválido ou 1900-01-01
                        }

                        //ws2.Cells[linhaAtual, 26] = ""; // Entrada Obra
                       //ws2.Cells[linhaAtual, 27] = ""; // Saída Obra
                        ws2.Cells[linhaAtual, 28] = "Sim"; // Autorização de Entrada em Obra

                        linhaAtual++;    // Começa na linha 22
                        dadosTrabalhadoresEntidades.Seguinte();
                    }
                    //MÁQUINAS E EQUIPAMENTOS
                    linhaAtual = linhaAtual + 2;
                    ws2.Cells[linhaAtual, 1] = "MÁQUINAS E EQUIPAMENTOS";
                    var rMaquinasEquipamentos = ws2.Range[ws2.Cells[linhaAtual, 1], ws2.Cells[linhaAtual + 2, 7]];
                    rMaquinasEquipamentos.Merge();
                    Negrito(rMaquinasEquipamentos);
                    Centro(rMaquinasEquipamentos);
                    Borda(rMaquinasEquipamentos);
                    rMaquinasEquipamentos.Interior.Color = ToOle(System.Drawing.Color.LightGray);

                    //cilco for do numero 1 ao 11 mas o numero 4 vai ocupar 2 colunas , numero 5 vai ocupar 3 colunas, numero 6 vai ocupar 5 colunas, numero 7 vai ocupar 3 colunas, e o numero 8 vai ocupar 2 colunas
                    int lastCol3 = 28;
                    int numRow3 = linhaAtual;
                    for (int c = 8, n = 1; c <= lastCol3 && n <= 11; n++)
                    {
                        Microsoft.Office.Interop.Excel.Range r;

                        if (n == 4) // n== 4 ocupa 2 colunas
                        {
                            r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c + 1]];
                            r.Merge();
                            ws2.Cells[numRow3, c] = n.ToString();
                            c += 2; // avança duas colunas
                        }
                        else if (n == 5) // n== 5 ocupa 3 colunas
                        {
                            r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c + 3]];
                            r.Merge();
                            ws2.Cells[numRow3, c] = n.ToString();
                            c += 4; // avança três colunas
                        }
                        else if (n == 6) // n== 6 ocupa 5 colunas 
                        {
                            r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c + 4]];
                            r.Merge();
                            ws2.Cells[numRow3, c] = n.ToString();
                            c += 5; // avança cinco colunas
                        }
                        else if (n == 7) // n== 7 ocupa 3 colunas 
                        {
                            r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c + 2]];
                            r.Merge();
                            ws2.Cells[numRow3, c] = n.ToString();
                            c += 3; // avança três colunas
                        }
                        else
                        {
                            // Célula normal
                            r = ws2.Range[ws2.Cells[numRow3, c], ws2.Cells[numRow3, c]];
                            ws2.Cells[numRow3
                                , c] = n.ToString();
                            c++; // avança uma coluna
                        }
                        // Estilos
                        Negrito(r);
                        Centro(r);
                        Borda(r);
                        r.Interior.Color = ToOle(System.Drawing.Color.LightGray);
                    }

                    // Linha abaixo dos números (Cabeçalhos)
                    linhaAtual = linhaAtual + 1;
                    ws2.Cells[linhaAtual, 8] = "Manual de Instruções em língua PT";
                    var rManualInstrucoes = ws2.Range[ws2.Cells[linhaAtual, 8], ws2.Cells[linhaAtual + 1, 8]];
                    rManualInstrucoes.Merge();
                    Negrito(rManualInstrucoes);
                    Borda(rManualInstrucoes);
                    Centro(rManualInstrucoes);

                    ws2.Cells[linhaAtual, 9] = "Declaração  Conformidade CE";
                    var rDeclaracaoConformidadeCE = ws2.Range[ws2.Cells[linhaAtual, 9], ws2.Cells[linhaAtual + 1, 9]];
                    rDeclaracaoConformidadeCE.Merge();
                    Negrito(rDeclaracaoConformidadeCE);
                    Borda(rDeclaracaoConformidadeCE);
                    Centro(rDeclaracaoConformidadeCE);

                    ws2.Cells[linhaAtual, 10] = "Plano de Manutenção";
                    var rPlanoManutencao = ws2.Range[ws2.Cells[linhaAtual, 10], ws2.Cells[linhaAtual + 1, 10]];
                    rPlanoManutencao.Merge();
                    Negrito(rPlanoManutencao);
                    Borda(rPlanoManutencao);
                    Centro(rPlanoManutencao);

                    ws2.Cells[linhaAtual, 11] = "Relatório de Verificação de Segurança (DL 50/2005)";
                    var rRelatorioVerificacaoSeguranca = ws2.Range[ws2.Cells[linhaAtual, 11], ws2.Cells[linhaAtual, 12]];
                    rRelatorioVerificacaoSeguranca.Merge();
                    Negrito(rRelatorioVerificacaoSeguranca);
                    Borda(rRelatorioVerificacaoSeguranca);
                    Centro(rRelatorioVerificacaoSeguranca);

                    ws2.Cells[linhaAtual, 13] = "Registo de Manutenção (Último)";
                    var rRegistoManutencao = ws2.Range[ws2.Cells[linhaAtual, 13], ws2.Cells[linhaAtual + 1, 16]];
                    rRegistoManutencao.Merge();
                    Negrito(rRegistoManutencao);
                    Borda(rRegistoManutencao);
                    Centro(rRegistoManutencao);

                    ws2.Cells[linhaAtual, 17] = "Seguro Casco (carta verde)";
                    var rSeguroCasco = ws2.Range[ws2.Cells[linhaAtual, 17], ws2.Cells[linhaAtual + 1, 21]];
                    rSeguroCasco.Merge();
                    Negrito(rSeguroCasco);
                    Borda(rSeguroCasco);
                    Centro(rSeguroCasco);

                    ws2.Cells[linhaAtual, 22] = "Manobrador (Se Aplicável)";
                    var rManobrador = ws2.Range[ws2.Cells[linhaAtual, 22], ws2.Cells[linhaAtual, 24]];
                    rManobrador.Merge();
                    Negrito(rManobrador);
                    Borda(rManobrador);
                    Centro(rManobrador);

                    ws2.Cells[linhaAtual, 25] = "Observações";
                    var rObservacoes = ws2.Range[ws2.Cells[linhaAtual, 25], ws2.Cells[linhaAtual + 2, 25]];
                    rObservacoes.Merge();
                    Negrito(rObservacoes);
                    Borda(rObservacoes);
                    Centro(rObservacoes);

                    ws2.Cells[linhaAtual, 26] = "Entrada em Obra";
                    var rEntradaObra3 = ws2.Range[ws2.Cells[linhaAtual, 26], ws2.Cells[linhaAtual + 1, 26]];
                    rEntradaObra3.Merge();
                    Negrito(rEntradaObra3);
                    Borda(rEntradaObra3);
                    Centro(rEntradaObra3);

                    ws2.Cells[linhaAtual, 27] = "Saída de Obra";
                    var rSaidaObra3 = ws2.Range[ws2.Cells[linhaAtual, 27], ws2.Cells[linhaAtual + 1, 27]];
                    rSaidaObra3.Merge();
                    Negrito(rSaidaObra3);
                    Borda(rSaidaObra3);
                    Centro(rSaidaObra3);

                    ws2.Cells[linhaAtual, 28] = "Autorização de Entrada em Obra";
                    var rAutorizacaoEntrada3 = ws2.Range[ws2.Cells[linhaAtual, 28], ws2.Cells[linhaAtual + 1, 28]];
                    rAutorizacaoEntrada3.Merge();
                    Negrito(rAutorizacaoEntrada3);
                    Borda(rAutorizacaoEntrada3);
                    Centro(rAutorizacaoEntrada3);


                    //
                    linhaAtual = linhaAtual + 1;
                    //coluna 11

                    ws2.Cells[linhaAtual, 11] = "Possui?";
                    var rPossui = ws2.Cells[linhaAtual, 11];
                    Negrito(rPossui);
                    Borda(rPossui);
                    Centro(rPossui);

                    ws2.Cells[linhaAtual, 12] = "Validade";
                    var rValidade10 = ws2.Cells[linhaAtual, 12];
                    Negrito(rValidade10);
                    Borda(rValidade10);
                    Centro(rValidade10);

                    ws2.Cells[linhaAtual, 22] = "Nome";
                    var rNomeManobrador = ws2.Range[ws2.Cells[linhaAtual, 22], ws2.Cells[linhaAtual + 1, 22]];
                    rNomeManobrador.Merge();
                    Negrito(rNomeManobrador);
                    Borda(rNomeManobrador);
                    Centro(rNomeManobrador);

                    ws2.Cells[linhaAtual, 23] = "Habilitações";
                    var rHabilitacoes = ws2.Range[ws2.Cells[linhaAtual, 23], ws2.Cells[linhaAtual, 24]];
                    rHabilitacoes.Merge();
                    Negrito(rHabilitacoes);
                    Borda(rHabilitacoes);
                    Centro(rHabilitacoes);


                    //
                    linhaAtual = linhaAtual + 1;
                    ws2.Cells[linhaAtual, 1] = "N.º";
                    var rNumEquipamento = ws2.Cells[linhaAtual, 1];
                    Centro(rNumEquipamento);
                    Borda(rNumEquipamento);

                    ws2.Cells[linhaAtual, 2] = "Marca/ Modelo";
                    var rMarcaModelo = ws2.Cells[linhaAtual, 2];
                    Centro(rMarcaModelo);
                    Borda(rMarcaModelo);

                    ws2.Cells[linhaAtual, 3] = "Tipo de Máquina"; // ocupa 2 colunas
                    var rTipoMaquina = ws2.Range[ws2.Cells[linhaAtual, 3], ws2.Cells[linhaAtual, 4]];
                    rTipoMaquina.Merge();
                    Centro(rTipoMaquina);
                    Borda(rTipoMaquina);

                    ws2.Cells[linhaAtual, 5] = "Número de Série"; // ocupa 3 colunas
                    var rNumSerie = ws2.Range[ws2.Cells[linhaAtual, 5], ws2.Cells[linhaAtual, 7]];
                    rNumSerie.Merge();
                    Centro(rNumSerie);
                    Borda(rNumSerie);

                    ws2.Cells[linhaAtual, 8] = "C ; N/C ; N/A";
                    var rCNCNA11 = ws2.Cells[linhaAtual, 8];
                    Centro(rCNCNA11);
                    Borda(rCNCNA11);

                    ws2.Cells[linhaAtual, 9] = "C ; N/C ; N/A";
                    var rCNCNA12 = ws2.Cells[linhaAtual, 9];
                    Centro(rCNCNA12);
                    Borda(rCNCNA12);

                    ws2.Cells[linhaAtual, 10] = "C ; N/C ; N/A";
                    var rCNCNA13 = ws2.Cells[linhaAtual, 10];
                    Centro(rCNCNA13);
                    Borda(rCNCNA13);

                    ws2.Cells[linhaAtual, 11] = "C ; N/C ; N/A";
                    var rCNCNA14 = ws2.Cells[linhaAtual, 11];
                    Centro(rCNCNA14);
                    Borda(rCNCNA14);

                    ws2.Cells[linhaAtual, 12] = "Data";
                    var rData14 = ws2.Cells[linhaAtual, 12];
                    Centro(rData14);
                    Borda(rData14);

                    ws2.Cells[linhaAtual, 13] = "C ; N/C ; N/A";
                    var rCNCNA15 = ws2.Cells[linhaAtual, 13];
                    Centro(rCNCNA15);
                    Borda(rCNCNA15);

                    ws2.Cells[linhaAtual, 14] = "N.º Horas (à entrada em Obra)";
                    var rNumHoras = ws2.Cells[linhaAtual, 14];
                    Centro(rNumHoras);
                    Borda(rNumHoras);

                    ws2.Cells[linhaAtual, 15] = "Validade"; // ocupa 2 colunas
                    var rValidade11 = ws2.Range[ws2.Cells[linhaAtual, 15], ws2.Cells[linhaAtual, 16]];
                    rValidade11.Merge();
                    Centro(rValidade11);
                    Borda(rValidade11);

                    ws2.Cells[linhaAtual, 17] = "Seguradora"; // ocupa 2 colunas
                    var rSeguradora = ws2.Range[ws2.Cells[linhaAtual, 17], ws2.Cells[linhaAtual, 18]];
                    rSeguradora.Merge();
                    Centro(rSeguradora);
                    Borda(rSeguradora);

                    ws2.Cells[linhaAtual, 19] = "N.º Apólice";
                    var rNumApolice = ws2.Cells[linhaAtual, 19];
                    Centro(rNumApolice);
                    Borda(rNumApolice);

                    ws2.Cells[linhaAtual, 20] = "C ; N/C ; N/A";
                    var rCNCNA16 = ws2.Cells[linhaAtual, 20];
                    Centro(rCNCNA16);
                    Borda(rCNCNA16);

                    ws2.Cells[linhaAtual, 21] = "Validade";
                    var rValidade12 = ws2.Cells[linhaAtual, 21];
                    Centro(rValidade12);
                    Borda(rValidade12);

                    ws2.Cells[linhaAtual, 23] = "Tipo";
                    var rTipoManobrador = ws2.Cells[linhaAtual, 23];
                    Centro(rTipoManobrador);
                    Borda(rTipoManobrador);

                    ws2.Cells[linhaAtual, 24] = "C ; N/C ; N/A";
                    var rCNCNA17 = ws2.Cells[linhaAtual, 24];
                    Centro(rCNCNA17);
                    Borda(rCNCNA17);

                    ws2.Cells[linhaAtual, 26] = "Data";
                    var rData15 = ws2.Cells[linhaAtual, 26];
                    Centro(rData15);
                    Borda(rData15);

                    ws2.Cells[linhaAtual, 27] = "Data";
                    var rData16 = ws2.Cells[linhaAtual, 27];
                    Centro(rData16);
                    Borda(rData16);

                    ws2.Cells[linhaAtual, 28] = "Sim / Não";
                    var rSimNao3 = ws2.Cells[linhaAtual, 28];
                    Centro(rSimNao3);
                    Borda(rSimNao3);

                    //Linhas dos equipamentos  TODO

                    var queryEquipamentosEntidade = $@"SELECT * fROM TDU_AD_Equipamentos WHERE id_empresa = '{id}';";
                    var dadosEquipamentosEntidades = BSO.Consulta(queryEquipamentosEntidade);
                    var numregistosEquipamentos = dadosEquipamentosEntidades.NumLinhas();
                    dadosEquipamentosEntidades.Inicio();
                   
                    for (int i = 0; i < numregistosEquipamentos; i++)
                    {
                        linhaAtual++;
                        ws2.Cells[linhaAtual, 1] = (i + 1).ToString(); // N.º
                        ws2.Cells[linhaAtual, 2] = dadosEquipamentosEntidades.DaValor<string>("marca"); // Marca/ Modelo
                        ws2.Cells[linhaAtual, 3] = dadosEquipamentosEntidades.DaValor<string>("tipo"); // Tipo de Máquina
                        ws2.Cells[linhaAtual, 5] = dadosEquipamentosEntidades.DaValor<string>("serie"); // Número de Série
                        string valorAnexo4 = dadosEquipamentosEntidades.DaValor<string>("anexo4");

                        ws2.Cells[linhaAtual, 8] = valorAnexo4 == "True" ? "C" : "N/C";

                        string valorAnexo1 = dadosEquipamentosEntidades.DaValor<string>("anexo1");
                        ws2.Cells[linhaAtual, 9] = valorAnexo1 == "True" ? "C" : "N/C";


                        ws2.Cells[linhaAtual, 10] = "";

                        string valorAnexo2 = dadosEquipamentosEntidades.DaValor<string>("anexo2");
                        ws2.Cells[linhaAtual, 11] = valorAnexo2 == "True" ? "C" : "N/C";
                        ws2.Cells[linhaAtual, 12] = "TODO"; // Validade
                        string valorAnexo3 = dadosEquipamentosEntidades.DaValor<string>("anexo3");
                        ws2.Cells[linhaAtual, 13] = valorAnexo3 == "True" ? "C" : "N/C";
                        ws2.Cells[linhaAtual, 14] = "TODO"; // N.º Horas (à entrada em Obra)
                        ws2.Cells[linhaAtual, 15] = "TODO"; // Validade
                        ws2.Cells[linhaAtual, 17] = "TODO"; // Seguradora
                        ws2.Cells[linhaAtual, 19] = "TODO"; // N.º Apólice
                        string valorAnexo5 = dadosEquipamentosEntidades.DaValor<string>("anexo5");
                        ws2.Cells[linhaAtual, 20] = valorAnexo5 == "True" ? "C" : "N/C";
                        ws2.Cells[linhaAtual, 21] = "TODO"; // Validade
                        ws2.Cells[linhaAtual, 23] = "TODO"; // Tipo
                        ws2.Cells[linhaAtual, 24] = "TODO"; // C ; N/C ; N/A



                        var dataEntradaStr2 = dadosAutorizacoesEntidades.DaValor<string>("Data_Entrada");

                        if (DateTime.TryParse(dataEntradaStr2, out DateTime dataEntrada2))
                        {
                            ws2.Cells[linhaAtual, 26] = dataEntrada2.ToString("dd-MM-yyyy");
                        }
                        else
                        {
                            ws2.Cells[linhaAtual, 26] = ""; // Ou algum valor padrão se a data for inválida
                        }
                        var dataSaidaStr2 = dadosAutorizacoesEntidades.DaValor<string>("Data_Saida");

                        if (DateTime.TryParse(dataSaidaStr2, out DateTime dataSaida2) && dataSaida2 != new DateTime(1900, 1, 1))
                        {
                            ws2.Cells[linhaAtual, 27] = dataSaida2.ToString("dd-MM-yyyy");
                        }
                        else
                        {
                            ws2.Cells[linhaAtual, 27] = ""; // Vazio se for null, inválido ou 1900-01-01
                        }
                        ws2.Cells[linhaAtual, 28] = "Sim"; // Entrada em Obra
                    }


                    // Page setup
                    ws2.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    ws2.PageSetup.LeftMargin = excelApp.InchesToPoints(0.4);
                    ws2.PageSetup.RightMargin = excelApp.InchesToPoints(0.3);
                    ws2.PageSetup.TopMargin = excelApp.InchesToPoints(0.5);
                    ws2.PageSetup.BottomMargin = excelApp.InchesToPoints(0.5);
                    ws2.PageSetup.Zoom = false;
                    ws2.PageSetup.FitToPagesWide = 1;
                    ws2.PageSetup.FitToPagesTall = false;


                }



                catch (System.Exception ex)
                {
                    MessageBox.Show("Erro ao criar segunda folha Excel: " + ex.Message, "Erro",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (ws2 != null) Marshal.ReleaseComObject(ws2);
                }
            }
        }

        public static string VerificaValidade(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return "N/A";
            // Decodifica entidades HTML
            valor = WebUtility.HtmlDecode(valor);

            // Procura data no formato dd/MM/yyyy
            Match match = Regex.Match(valor, @"\b\d{2}/\d{2}/\d{4}\b");
       
            if (!match.Success)
                return "N/A";

            string dataStr = match.Value;
            DateTime dataValidade;

            if (DateTime.TryParseExact(dataStr, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out dataValidade))
            {
                if (dataValidade < DateTime.Today)
                    return "N/C"; // vencido
                else
                    return "C";   // válido
            }

            return "N/A"; // caso a conversão falhe
        }
        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
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

        private void f4TabelaSQL1_Load(object sender, EventArgs e)
        {
            PriSDKContext.Initialize(BSO, PSO);
            InitializeSDKControls();
        }

        private void QuadroControlo_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                //Ensure that resources released.
                f4TabelaSQL1.Termina();
                controlsInitialized = false;
            }
            catch { }
        }
        private void InitializeSDKControls()
        {
            //Initializes controls
            if (!controlsInitialized)
            {

                f4TabelaSQL1.CampoChave = "Codigo";
                f4TabelaSQL1.CampoDescricao = "Descricao";
                f4TabelaSQL1.SelectionFormula = "SELECT Codigo, Descricao FROM COP_Obras WHERE ObraPaiID is null AND Estado = 'CONS' ";// WHERe ObraPaiID is null order by Codigo desc";

                f4TabelaSQL1.Caption = "Codigo:";
                f4TabelaSQL1.MostraCaption = true;
                f4TabelaSQL1.WidthCaption = 500;
                f4TabelaSQL1.Caption = "Codigo:";
                f4TabelaSQL1.AliasCampoChave = "Codigo";

                f4TabelaSQL1.Change += F4TabelaSQL1_Change;
                f4TabelaSQL1.Inicializa(PriSDKContext.SdkContext);
                controlsInitialized = true;


            }
        }
        public string ObraCodigo { get; set; }
        private void F4TabelaSQL1_Change(object sender, F4TabelaSQL.ChangeEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(e.Value))
            {
                var queryValidaCodigo = $"SELECT COUNT(*) AS Total FROM COP_Obras WHERE Codigo = '{e.Value}'";
                var resultado = BSO.Consulta(queryValidaCodigo).DaValor<int>("Total");

                if (resultado > 0)
                {
                    ObraCodigo = e.Value; // Armazena o código da obra selecionada
                    DadosLista();
                    FiltrarPorObra(e.Value);
                }
                else
                {
                    // Ignora valores inválidos, evita o erro
                    // ou exibe uma mensagem temporária se quiser
                }
            }
            else
            {
                DadosLista();
            }
        }

        private void FiltrarPorObra(string value)
        {
            for (int i = dataGridView1.Rows.Count - 1; i >= 0; i--)
            {
                var row = dataGridView1.Rows[i];
                if (!row.IsNewRow)
                {
                    var cellValue = row.Cells[2].Value?.ToString() ?? "";
                    bool autorizado = VerificaSeTemObra(cellValue, value);

                    if (!autorizado)
                    {
                        dataGridView1.Rows.RemoveAt(i);
                    }
                }
            }
        }



        private bool VerificaSeTemObra(string cellValue, string value)
        {
            try
            {
                // 1. Obter o ID da obra pelo código informado

                var queryObra = $"SELECT ID FROM COP_Obras WHERE Codigo = '{value}'";
                var obraID = BSO.Consulta(queryObra).DaValor<string>("ID");

                if (string.IsNullOrEmpty(obraID))
                {
                    return false;
                }
                // 2. Obter o ID da entidade pelo nome
                var queryEntidade = $"SELECT ID FROM Geral_Entidade WHERE Nome = '{cellValue}'";
                var idEntidade = BSO.Consulta(queryEntidade).DaValor<string>("ID");

                if (string.IsNullOrEmpty(idEntidade))
                {
                    return false;
                }

                // 3. Verificar se a entidade tem alguma obra autorizada
                var queryAutorizacoes = $"SELECT Codigo_Obra FROM TDU_AD_Autorizacoes WHERE ID_Entidade = '{idEntidade}'";
                var resObrasAutorizadas = BSO.Consulta(queryAutorizacoes);
                var numitem = resObrasAutorizadas.NumLinhas();

                if (numitem == 0)
                {
                    return false;
                }

                bool temAutorizacaoNaObraPai = false;

                resObrasAutorizadas.Inicio();

                while (!resObrasAutorizadas.NoFim())
                {
                    var codigoObraFilha = resObrasAutorizadas.Valor("Codigo_Obra")?.ToString();

                    if (!string.IsNullOrEmpty(codigoObraFilha))
                    {
                        var queryObraFilha = $"SELECT ObraPaiID FROM COP_Obras WHERE Codigo = '{codigoObraFilha}'";
                        var obraPaiID = BSO.Consulta(queryObraFilha).DaValor<string>("ObraPaiID");

                        if (obraPaiID == obraID)
                        {
                            temAutorizacaoNaObraPai = true;
                            break;
                        }
                    }

                    resObrasAutorizadas.Seguinte();
                }

                if (temAutorizacaoNaObraPai)
                {
                    // MessageBox.Show($"A entidade {cellValue} TEM autorização numa obra cuja obra pai é {value}.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }

                return false;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Erro: {ex.Message}", "Exceção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;  // <-- Essencial para evitar erro de compilação
            }
        }

        private void QuadroControlo_Load_1(object sender, EventArgs e)
        {
            label1.ForeColor = System.Drawing.Color.FromArgb(59, 89, 152);
        }

        private void BT_ImprimirJPA_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> idsSelecionados = new List<string>();
                string obraComum = f4TabelaSQL1.Text;
                //verifica sem aturorizaçoes em aguma obra 1º
                /*  Dictionary<string, List<string>> autorizacoes;

                  var autorizado = VerificaAutorizacao(idsSelecionados, out autorizacoes, out obraComum);
                 */
                string idPadrao = "2A8C7ECD-309B-49F9-A337-203B45CED948";

                idsSelecionados.Add(idPadrao);

                ExportarParaExcel2(idsSelecionados, ObraCodigo);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Erro ao exportar para Excel: " + ex.Message);
            }
        }

        private async void BT_CriarTrabalhadores_Click(object sender, EventArgs e)
        {
            try
            {
                // Verificar se há linhas selecionadas
                List<string> idsSelecionados = new List<string>();

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[" "].Value != null && (bool)row.Cells[" "].Value)
                    {
                        string id = row.Cells["ID"].Value?.ToString();
                        if (!string.IsNullOrEmpty(id))
                            idsSelecionados.Add(id);
                    }
                }

                if (idsSelecionados.Count == 0)
                {
                    MessageBox.Show("Por favor, selecione pelo menos uma empresa.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (idsSelecionados.Count > 1)
                {
                    MessageBox.Show("Por favor, selecione apenas uma empresa de cada vez.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string idEmpresa = idsSelecionados[0];

                // Buscar código da empresa
                string queryEmpresa = $"SELECT ID, Nome FROM Geral_Entidade WHERE ID = '{idEmpresa}'";
                var dadosEmpresa = BSO.Consulta(queryEmpresa);

                if (dadosEmpresa.Vazia())
                {
                    MessageBox.Show("Empresa não encontrada.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                dadosEmpresa.Inicio();
                string codigoEmpresa = dadosEmpresa.Valor("ID")?.ToString() ?? "";
                string nomeEmpresa = dadosEmpresa.Valor("Nome")?.ToString() ?? "";

                // Buscar trabalhadores da empresa
                string queryTrabalhadores = $@"
                    SELECT nome 
                    FROM TDU_AD_Trabalhadores 
                    WHERE id_empresa = '{idEmpresa}'";

                var trabalhadores = BSO.Consulta(queryTrabalhadores);

                if (trabalhadores.Vazia())
                {
                    MessageBox.Show("Nenhum trabalhador encontrado para esta empresa.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int totalEnviados = 0;
                int totalErros = 0;
                trabalhadores.Inicio();

                while (!trabalhadores.NoFim())
                {
                    string nomeTrabalhador = trabalhadores.Valor("nome")?.ToString() ?? "";

                    if (!string.IsNullOrEmpty(nomeTrabalhador))
                    {
                        // Gerar QR Code único
                        string qrCode = GerarQRCode();

                        // Enviar para API - agora enviando o nome da empresa
                        bool sucesso = await EnviarTrabalhadorParaAPI(nomeTrabalhador, qrCode, nomeEmpresa);

                        if (sucesso)
                            totalEnviados++;
                        else
                            totalErros++;
                    }

                    trabalhadores.Seguinte();
                }

                string mensagem = $"Processo concluído!\n\n" +
                                 $"Trabalhadores enviados com sucesso: {totalEnviados}\n" +
                                 $"Erros: {totalErros}";

                MessageBox.Show(mensagem, "Resultado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Erro ao criar trabalhadores: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GerarQRCode()
        {
            // Gerar um código único usando timestamp e GUID
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string guid = Guid.NewGuid().ToString("N").Substring(0, 8).ToUpper();
            return $"QR{timestamp}{guid}";
        }

        private async Task<bool> EnviarTrabalhadorParaAPI(string nome, string qrCode, string empresa)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);

                    var dados = new
                    {
                        Nome = nome,
                        Qrcode = qrCode,
                        Empresa = empresa
                    };

                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(dados);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync("https://backend.advir.pt/api/externos-jpa", content);

                    return response.IsSuccessStatusCode;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Erro ao enviar trabalhador {nome}: {ex.Message}", "Erro API", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
    }
}