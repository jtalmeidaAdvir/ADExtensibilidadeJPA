using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    partial class Menu
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        // Control declarations - only declared once here
        private TextBox TXT_Nome;
        private Button BTF4;
        private TextBox TXT_Codigo;
        private TabControl tabControl1;
        private TabPage tabPage1;
        private Panel AlertaValidadeAlvara;
        private Panel pnlAutorizacaoObra;
        private Button btnAutorizarEntrada;
        private Label lblDataEntrada;
        private DateTimePicker dtpDataEntrada;
        private Label lblDataSaida;
        private DateTimePicker dtpDataSaida;
        private Label lblContratoSubempreitada;
        private TextBox txtContratoSubempreitada;
        private Label lblStatusEntrada;
        private ComboBox cmbStatusEntrada;
        private Panel pnlDadosObra;
        private ComboBox cb_ReciboPagSegSocial;
        private Label label22;
        private DateTimePicker TXT_AlvaraValidade;
        private DateTimePicker TXT_FolhaPagSegSocial;
        private DateTimePicker TXT_NaoDivSegSocial;
        private DateTimePicker TXT_NaoDivFinancas;
        private TextBox TXT_Alvara;
        private TextBox TXT_Contribuinte;
        private TextBox TXT_Sede;
        private Label label8;
        private Label label7;
        private Label label6;
        private Label label5;
        private Label label4;
        private Label label3;
        private Label label2;
        private TabPage tabPage2;
        private DataGridView dataGridView1;
        private ComboBox cb_Obras;
        private ToolStrip toolStrip1;
        private ToolStripButton BT_Salvar_Click;
        private DataGridViewTextBoxColumn EntradaObra_;
        private DataGridViewTextBoxColumn SaidaObra_;
        private DataGridViewTextBoxColumn ContratoSubempreitada;
        private DataGridViewTextBoxColumn StatusAutorizacao;
        private DataGridViewCheckBoxColumn AutorizacaoEntrada;
        private Panel panelDadosEmpresa;
        private GroupBox groupBoxInfoBasica;
        private GroupBox groupBoxSituacaoFiscal;
        private Panel panelObras;
        private GroupBox groupBoxObras;
        private Label label17;
        private Button btnGravarObra;
        private Label lblSelecionar;
        private Label labelCaminho;
        private TextBox txtCaminhoPasta;
        private Button btnSelecionarPasta;
        private Button btnAnexoFinancas;
        private Label lblAnexoFinancas;
        private Button btnAnexoSegSocial;
        private Label lblAnexoSegSocial;
        private Label lblFolhaPagSS;
        private Button btnAnexoFolhaPag;
        private Label lblAnexoApoliceAT;
        private Button btnAnexoApoliceAT;
        private Label lblAnexoApoliceRC;
        private Button btnAnexoApoliceRC;
        private Label lblAnexoHorarioTrabalho;
        private Button btnAnexoHorarioTrabalho;
        private Label lblAnexoD;
        private Button btnAnexoD;
        private Label lblDecTrabEmigr;
        private Button btnDecTrabEmigr;
        private Label lblInscricaoSS;
        private Button btnInscricaoSS;
        private Label lblHorarioTrabalhoTitle;
        private Label lblApoliceATTitle;
        private Label lblApoliceRCTitle;
        private Label lblAnexoDTitle;
        private Label lblDecTrabEmigrTitle;
        private Label lblInscricaoSSTitle;
        private Button btnAnexarDocumentoGeral;
        private Button btnVerificarDocumentosFaltantes;
        private Panel panelModalDocumentos;
        private ComboBox cmbTipoDocumento;
        private Button btnConfirmarAnexo;
        private Button btnCancelarAnexo;
        private Label lblTipoDocumento;
        private Label lblValidade;
        private DateTimePicker dtpValidade;
        private Button btnAbrirPastaAnexos;
        // Trabalhadores tab controls
        private TextBox txtNomeTrabalhador;
        private ComboBox cmbTipoDocumentoTrabalhador;
        private TextBox txtNumDocumento;
        private DateTimePicker dtpValidadeDocumento;
        private TextBox txtNIF;
        private TextBox txtNumSS;
        private CheckBox chkFichaAptidaoMedica;
        private CheckBox chkCredenciacao;
        private TextBox txtCredenciacao;
        private CheckBox chkFichaEPI;
        private DataGridView gridTrabalhadores;
        private ComboBox cmbObrasTrabalhador;
        private Label lblFichaAptidaoAnexo;
        private Label lblCredenciacaoAnexo;
        private Label lblFichaEPIAnexo;
        private Button btnAdicionarTrabalhador;
        private Button btnEditarTrabalhador;
        private Button btnExcluirTrabalhador;
        private Button btnSalvarTrabalhador;
        private Button btnAutorizarTrabalhador;
        private Button btnAnexarFichaAptidao;
        private Button btnAnexarCredenciacao;
        private Button btnAnexarFichaEPI;
        private Panel panelTrabalhadores;
        private GroupBox groupBoxInfoTrabalhador;
        private GroupBox groupBoxListaTrabalhadores;
        private Panel panelBotoesTrabalhador;


        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Menu));
            this.TXT_Nome = new System.Windows.Forms.TextBox();
            this.BTF4 = new System.Windows.Forms.Button();
            this.TXT_Codigo = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panelDadosEmpresa = new System.Windows.Forms.Panel();
            this.groupBoxInfoBasica = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TXT_Sede = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TXT_Contribuinte = new System.Windows.Forms.TextBox();
            this.groupBoxSituacaoFiscal = new System.Windows.Forms.GroupBox();
            this.btnAbrirPastaAnexos = new System.Windows.Forms.Button();
            this.panelModalDocumentos = new System.Windows.Forms.Panel();
            this.lblTipoDocumento = new System.Windows.Forms.Label();
            this.cmbTipoDocumento = new System.Windows.Forms.ComboBox();
            this.lblValidade = new System.Windows.Forms.Label();
            this.dtpValidade = new System.Windows.Forms.DateTimePicker();
            this.btnConfirmarAnexo = new System.Windows.Forms.Button();
            this.btnCancelarAnexo = new System.Windows.Forms.Button();
            this.lblFolhaPagSS = new System.Windows.Forms.Label();
            this.btnAnexoFolhaPag = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.TXT_NaoDivFinancas = new System.Windows.Forms.DateTimePicker();
            this.btnAnexoFinancas = new System.Windows.Forms.Button();
            this.lblAnexoFinancas = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.TXT_NaoDivSegSocial = new System.Windows.Forms.DateTimePicker();
            this.btnAnexoSegSocial = new System.Windows.Forms.Button();
            this.lblAnexoSegSocial = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.TXT_FolhaPagSegSocial = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.cb_ReciboPagSegSocial = new System.Windows.Forms.ComboBox();
            this.labelCaminho = new System.Windows.Forms.Label();
            this.txtCaminhoPasta = new System.Windows.Forms.TextBox();
            this.btnSelecionarPasta = new System.Windows.Forms.Button();
            this.lblAnexoHorarioTrabalho = new System.Windows.Forms.Label();
            this.btnAnexoHorarioTrabalho = new System.Windows.Forms.Button();
            this.lblAnexoApoliceAT = new System.Windows.Forms.Label();
            this.btnAnexoApoliceAT = new System.Windows.Forms.Button();
            this.lblAnexoApoliceRC = new System.Windows.Forms.Label();
            this.btnAnexoApoliceRC = new System.Windows.Forms.Button();
            this.lblAnexoD = new System.Windows.Forms.Label();
            this.btnAnexoD = new System.Windows.Forms.Button();
            this.lblDecTrabEmigr = new System.Windows.Forms.Label();
            this.btnDecTrabEmigr = new System.Windows.Forms.Button();
            this.lblInscricaoSS = new System.Windows.Forms.Label();
            this.btnInscricaoSS = new System.Windows.Forms.Button();
            this.lblHorarioTrabalhoTitle = new System.Windows.Forms.Label();
            this.lblApoliceATTitle = new System.Windows.Forms.Label();
            this.lblApoliceRCTitle = new System.Windows.Forms.Label();
            this.lblAnexoDTitle = new System.Windows.Forms.Label();
            this.lblDecTrabEmigrTitle = new System.Windows.Forms.Label();
            this.lblInscricaoSSTitle = new System.Windows.Forms.Label();
            this.btnAnexarDocumentoGeral = new System.Windows.Forms.Button();
            this.btnVerificarDocumentosFaltantes = new System.Windows.Forms.Button();
            this.groupBoxApolices = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.cb_ApoliceAT = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.TXT_ReciboApoliceAT = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.TXT_Alvara = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.cb_ApoliceRC = new System.Windows.Forms.ComboBox();
            this.label22 = new System.Windows.Forms.Label();
            this.TXT_AlvaraValidade = new System.Windows.Forms.DateTimePicker();
            this.label12 = new System.Windows.Forms.Label();
            this.AlertaValidadeAlvara = new System.Windows.Forms.Panel();
            this.TXT_ReciboRC = new System.Windows.Forms.TextBox();
            this.groupBoxDeclaracoes = new System.Windows.Forms.GroupBox();
            this.label13 = new System.Windows.Forms.Label();
            this.cb_HorarioTrabalho = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.cb_DecTrabIlegais = new System.Windows.Forms.ComboBox();
            this.label15 = new System.Windows.Forms.Label();
            this.cb_DecRespEstaleiro = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.cb_DecConhecimPSS = new System.Windows.Forms.ComboBox();
            this.panelObras = new System.Windows.Forms.Panel();
            this.groupBoxObras = new System.Windows.Forms.GroupBox();
            this.label17 = new System.Windows.Forms.Label();
            this.cb_Obras = new System.Windows.Forms.ComboBox();
            this.btnGravarObra = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.EntradaObra_ = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SaidaObra_ = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ContratoSubempreitada = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StatusAutorizacao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AutorizacaoEntrada = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.btnAutorizarEntrada = new System.Windows.Forms.Button();
            this.lblDataEntrada = new System.Windows.Forms.Label();
            this.dtpDataEntrada = new System.Windows.Forms.DateTimePicker();
            this.lblDataSaida = new System.Windows.Forms.Label();
            this.dtpDataSaida = new System.Windows.Forms.DateTimePicker();
            this.lblContratoSubempreitada = new System.Windows.Forms.Label();
            this.txtContratoSubempreitada = new System.Windows.Forms.TextBox();
            this.pnlDadosObra = new System.Windows.Forms.Panel();
            this.pnlAutorizacaoObra = new System.Windows.Forms.Panel();
            this.lblAutorizacao = new System.Windows.Forms.Label();
            this.cmbAutorizacaoStatus = new System.Windows.Forms.ComboBox();
            this.lblObservacao = new System.Windows.Forms.Label();
            this.txtObservacao = new System.Windows.Forms.TextBox();
            this.btnSalvarAutorizacao = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panelTrabalhadores = new System.Windows.Forms.Panel();
            this.groupBoxInfoTrabalhador = new System.Windows.Forms.GroupBox();
            this.txtNomeTrabalhador = new System.Windows.Forms.TextBox();
            this.cmbTipoDocumentoTrabalhador = new System.Windows.Forms.ComboBox();
            this.txtNumDocumento = new System.Windows.Forms.TextBox();
            this.dtpValidadeDocumento = new System.Windows.Forms.DateTimePicker();
            this.txtNIF = new System.Windows.Forms.TextBox();
            this.txtNumSS = new System.Windows.Forms.TextBox();
            this.chkFichaAptidaoMedica = new System.Windows.Forms.CheckBox();
            this.chkCredenciacao = new System.Windows.Forms.CheckBox();
            this.txtCredenciacao = new System.Windows.Forms.TextBox();
            this.chkFichaEPI = new System.Windows.Forms.CheckBox();
            this.gridTrabalhadores = new System.Windows.Forms.DataGridView();
            this.cmbObrasTrabalhador = new System.Windows.Forms.ComboBox();
            this.lblFichaAptidaoAnexo = new System.Windows.Forms.Label();
            this.lblCredenciacaoAnexo = new System.Windows.Forms.Label();
            this.lblFichaEPIAnexo = new System.Windows.Forms.Label();
            this.btnAdicionarTrabalhador = new System.Windows.Forms.Button();
            this.btnEditarTrabalhador = new System.Windows.Forms.Button();
            this.btnExcluirTrabalhador = new System.Windows.Forms.Button();
            this.btnSalvarTrabalhador = new System.Windows.Forms.Button();
            this.btnAutorizarTrabalhador = new System.Windows.Forms.Button();
            this.btnAnexarFichaAptidao = new System.Windows.Forms.Button();
            this.btnAnexarCredenciacao = new System.Windows.Forms.Button();
            this.btnAnexarFichaEPI = new System.Windows.Forms.Button();
            this.groupBoxListaTrabalhadores = new System.Windows.Forms.GroupBox();
            this.panelBotoesTrabalhador = new System.Windows.Forms.Panel();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.BT_Salvar_Click = new System.Windows.Forms.ToolStripButton();
            this.lblSelecionar = new System.Windows.Forms.Label();
            this.lblStatusEntrada = new System.Windows.Forms.Label();
            this.cmbStatusEntrada = new System.Windows.Forms.ComboBox();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panelDadosEmpresa.SuspendLayout();
            this.groupBoxInfoBasica.SuspendLayout();
            this.groupBoxSituacaoFiscal.SuspendLayout();
            this.panelModalDocumentos.SuspendLayout();
            this.groupBoxApolices.SuspendLayout();
            this.groupBoxDeclaracoes.SuspendLayout();
            this.panelObras.SuspendLayout();
            this.groupBoxObras.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.pnlAutorizacaoObra.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panelTrabalhadores.SuspendLayout();
            this.groupBoxInfoTrabalhador.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridTrabalhadores)).BeginInit();
            this.groupBoxListaTrabalhadores.SuspendLayout();
            this.panelBotoesTrabalhador.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TXT_Nome
            // 
            this.TXT_Nome.BackColor = System.Drawing.Color.White;
            this.TXT_Nome.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TXT_Nome.Enabled = false;
            this.TXT_Nome.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold);
            this.TXT_Nome.Location = new System.Drawing.Point(175, 35);
            this.TXT_Nome.Name = "TXT_Nome";
            this.TXT_Nome.Size = new System.Drawing.Size(466, 24);
            this.TXT_Nome.TabIndex = 0;
            // 
            // BTF4
            // 
            this.BTF4.BackColor = System.Drawing.Color.LightSteelBlue;
            this.BTF4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTF4.Location = new System.Drawing.Point(647, 35);
            this.BTF4.Name = "BTF4";
            this.BTF4.Size = new System.Drawing.Size(55, 24);
            this.BTF4.TabIndex = 1;
            this.BTF4.Text = "F4";
            this.BTF4.UseVisualStyleBackColor = false;
            this.BTF4.Click += new System.EventHandler(this.BTF4_Click);
            // 
            // TXT_Codigo
            // 
            this.TXT_Codigo.BackColor = System.Drawing.Color.White;
            this.TXT_Codigo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TXT_Codigo.Enabled = false;
            this.TXT_Codigo.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold);
            this.TXT_Codigo.Location = new System.Drawing.Point(70, 35);
            this.TXT_Codigo.Name = "TXT_Codigo";
            this.TXT_Codigo.Size = new System.Drawing.Size(100, 24);
            this.TXT_Codigo.TabIndex = 2;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("Calibri", 9.5F);
            this.tabControl1.Location = new System.Drawing.Point(15, 65);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(726, 731);
            this.tabControl1.TabIndex = 3;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.tabPage1.Controls.Add(this.panelDadosEmpresa);
            this.tabPage1.Controls.Add(this.panelObras);
            this.tabPage1.Location = new System.Drawing.Point(4, 24);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(718, 703);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Empresa";
            // 
            // panelDadosEmpresa
            // 
            this.panelDadosEmpresa.Controls.Add(this.groupBoxInfoBasica);
            this.panelDadosEmpresa.Controls.Add(this.panelModalDocumentos);
            this.panelDadosEmpresa.Controls.Add(this.groupBoxSituacaoFiscal);
            this.panelDadosEmpresa.Controls.Add(this.groupBoxApolices);
            this.panelDadosEmpresa.Controls.Add(this.groupBoxDeclaracoes);
            this.panelDadosEmpresa.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelDadosEmpresa.Location = new System.Drawing.Point(3, 3);
            this.panelDadosEmpresa.Name = "panelDadosEmpresa";
            this.panelDadosEmpresa.Size = new System.Drawing.Size(712, 361);
            this.panelDadosEmpresa.TabIndex = 92;
            // 
            // groupBoxInfoBasica
            // 
            this.groupBoxInfoBasica.Controls.Add(this.label1);
            this.groupBoxInfoBasica.Controls.Add(this.textBox1);
            this.groupBoxInfoBasica.Controls.Add(this.label2);
            this.groupBoxInfoBasica.Controls.Add(this.TXT_Sede);
            this.groupBoxInfoBasica.Controls.Add(this.label3);
            this.groupBoxInfoBasica.Controls.Add(this.TXT_Contribuinte);
            this.groupBoxInfoBasica.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxInfoBasica.Location = new System.Drawing.Point(8, 5);
            this.groupBoxInfoBasica.Name = "groupBoxInfoBasica";
            this.groupBoxInfoBasica.Size = new System.Drawing.Size(697, 132);
            this.groupBoxInfoBasica.TabIndex = 0;
            this.groupBoxInfoBasica.TabStop = false;
            this.groupBoxInfoBasica.Text = "Informações Básicas";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 9F);
            this.label1.Location = new System.Drawing.Point(29, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 14);
            this.label1.TabIndex = 73;
            this.label1.Text = "Nome";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Calibri", 9F);
            this.textBox1.Location = new System.Drawing.Point(69, 16);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(611, 22);
            this.textBox1.TabIndex = 74;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9F);
            this.label2.Location = new System.Drawing.Point(29, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 14);
            this.label2.TabIndex = 56;
            this.label2.Text = "Sede";
            // 
            // TXT_Sede
            // 
            this.TXT_Sede.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_Sede.Location = new System.Drawing.Point(69, 44);
            this.TXT_Sede.Name = "TXT_Sede";
            this.TXT_Sede.Size = new System.Drawing.Size(611, 22);
            this.TXT_Sede.TabIndex = 71;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 9F);
            this.label3.Location = new System.Drawing.Point(9, 75);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 14);
            this.label3.TabIndex = 57;
            this.label3.Text = "NIF/NIPC";
            // 
            // TXT_Contribuinte
            // 
            this.TXT_Contribuinte.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_Contribuinte.Location = new System.Drawing.Point(69, 72);
            this.TXT_Contribuinte.Name = "TXT_Contribuinte";
            this.TXT_Contribuinte.Size = new System.Drawing.Size(611, 22);
            this.TXT_Contribuinte.TabIndex = 72;
            // 
            // groupBoxSituacaoFiscal
            // 
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAbrirPastaAnexos);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblFolhaPagSS);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAnexoFolhaPag);
            this.groupBoxSituacaoFiscal.Controls.Add(this.label5);
            this.groupBoxSituacaoFiscal.Controls.Add(this.TXT_NaoDivFinancas);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAnexoFinancas);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblAnexoFinancas);
            this.groupBoxSituacaoFiscal.Controls.Add(this.label6);
            this.groupBoxSituacaoFiscal.Controls.Add(this.TXT_NaoDivSegSocial);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAnexoSegSocial);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblAnexoSegSocial);
            this.groupBoxSituacaoFiscal.Controls.Add(this.label7);
            this.groupBoxSituacaoFiscal.Controls.Add(this.TXT_FolhaPagSegSocial);
            this.groupBoxSituacaoFiscal.Controls.Add(this.label8);
            this.groupBoxSituacaoFiscal.Controls.Add(this.cb_ReciboPagSegSocial);
            this.groupBoxSituacaoFiscal.Controls.Add(this.labelCaminho);
            this.groupBoxSituacaoFiscal.Controls.Add(this.txtCaminhoPasta);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnSelecionarPasta);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblAnexoHorarioTrabalho);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAnexoHorarioTrabalho);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblAnexoApoliceAT);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAnexoApoliceAT);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblAnexoApoliceRC);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAnexoApoliceRC);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblAnexoD);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAnexoD);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblDecTrabEmigr);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnDecTrabEmigr);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblInscricaoSS);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnInscricaoSS);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblHorarioTrabalhoTitle);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblApoliceATTitle);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblApoliceRCTitle);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblAnexoDTitle);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblDecTrabEmigrTitle);
            this.groupBoxSituacaoFiscal.Controls.Add(this.lblInscricaoSSTitle);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnAnexarDocumentoGeral);
            this.groupBoxSituacaoFiscal.Controls.Add(this.btnVerificarDocumentosFaltantes);
            this.groupBoxSituacaoFiscal.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxSituacaoFiscal.Location = new System.Drawing.Point(8, 143);
            this.groupBoxSituacaoFiscal.Name = "groupBoxSituacaoFiscal";
            this.groupBoxSituacaoFiscal.Size = new System.Drawing.Size(352, 208);
            this.groupBoxSituacaoFiscal.TabIndex = 1;
            this.groupBoxSituacaoFiscal.TabStop = false;
            this.groupBoxSituacaoFiscal.Text = "Situação Fiscal";
            // 
            // btnAbrirPastaAnexos
            // 
            this.btnAbrirPastaAnexos.BackColor = System.Drawing.Color.LightSalmon;
            this.btnAbrirPastaAnexos.DialogResult = System.Windows.Forms.DialogResult.Retry;
            this.btnAbrirPastaAnexos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAbrirPastaAnexos.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnAbrirPastaAnexos.Location = new System.Drawing.Point(12, 48);
            this.btnAbrirPastaAnexos.Name = "btnAbrirPastaAnexos";
            this.btnAbrirPastaAnexos.Size = new System.Drawing.Size(130, 24);
            this.btnAbrirPastaAnexos.TabIndex = 126;
            this.btnAbrirPastaAnexos.Text = " Abrir Pasta Anexos";
            this.btnAbrirPastaAnexos.UseVisualStyleBackColor = false;
            this.btnAbrirPastaAnexos.Click += new System.EventHandler(this.btnAbrirPastaAnexos_Click_1);
            // 
            // panelModalDocumentos
            // 
            this.panelModalDocumentos.BackColor = System.Drawing.Color.White;
            this.panelModalDocumentos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelModalDocumentos.Controls.Add(this.lblTipoDocumento);
            this.panelModalDocumentos.Controls.Add(this.cmbTipoDocumento);
            this.panelModalDocumentos.Controls.Add(this.lblValidade);
            this.panelModalDocumentos.Controls.Add(this.dtpValidade);
            this.panelModalDocumentos.Controls.Add(this.btnConfirmarAnexo);
            this.panelModalDocumentos.Controls.Add(this.btnCancelarAnexo);
            this.panelModalDocumentos.Location = new System.Drawing.Point(366, 143);
            this.panelModalDocumentos.Name = "panelModalDocumentos";
            this.panelModalDocumentos.Size = new System.Drawing.Size(339, 208);
            this.panelModalDocumentos.TabIndex = 125;
            this.panelModalDocumentos.Visible = false;
            // 
            // lblTipoDocumento
            // 
            this.lblTipoDocumento.AutoSize = true;
            this.lblTipoDocumento.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold);
            this.lblTipoDocumento.Location = new System.Drawing.Point(78, 20);
            this.lblTipoDocumento.Name = "lblTipoDocumento";
            this.lblTipoDocumento.Size = new System.Drawing.Size(154, 18);
            this.lblTipoDocumento.TabIndex = 3;
            this.lblTipoDocumento.Text = "Selecione o documento";
            // 
            // cmbTipoDocumento
            // 
            this.cmbTipoDocumento.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTipoDocumento.Font = new System.Drawing.Font("Calibri", 10F);
            this.cmbTipoDocumento.FormattingEnabled = true;
            this.cmbTipoDocumento.Items.AddRange(new object[] {
            "Não Div. Financas",
            "Não Div. Seg. Social",
            "Folha Pag. S.S.",
            "Apólice AT",
            "Apólice RC",
            "Horário Trabalho",
            "Anexo D",
            "Dec. Trab. Emigrantes",
            "Inscrição SS",
            "Outro documento"});
            this.cmbTipoDocumento.Location = new System.Drawing.Point(30, 50);
            this.cmbTipoDocumento.Name = "cmbTipoDocumento";
            this.cmbTipoDocumento.Size = new System.Drawing.Size(291, 23);
            this.cmbTipoDocumento.TabIndex = 0;
            this.cmbTipoDocumento.SelectedIndexChanged += new System.EventHandler(this.cmbTipoDocumento_SelectedIndexChanged);
            // 
            // lblValidade
            // 
            this.lblValidade.AutoSize = true;
            this.lblValidade.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblValidade.Location = new System.Drawing.Point(30, 85);
            this.lblValidade.Name = "lblValidade";
            this.lblValidade.Size = new System.Drawing.Size(59, 14);
            this.lblValidade.TabIndex = 3;
            this.lblValidade.Text = "Validade:";
            // 
            // dtpValidade
            // 
            this.dtpValidade.Font = new System.Drawing.Font("Calibri", 9F);
            this.dtpValidade.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpValidade.Location = new System.Drawing.Point(30, 105);
            this.dtpValidade.Name = "dtpValidade";
            this.dtpValidade.ShowCheckBox = true;
            this.dtpValidade.Size = new System.Drawing.Size(291, 22);
            this.dtpValidade.TabIndex = 4;
            // 
            // btnConfirmarAnexo
            // 
            this.btnConfirmarAnexo.BackColor = System.Drawing.Color.LightGreen;
            this.btnConfirmarAnexo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnConfirmarAnexo.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnConfirmarAnexo.Location = new System.Drawing.Point(88, 150);
            this.btnConfirmarAnexo.Name = "btnConfirmarAnexo";
            this.btnConfirmarAnexo.Size = new System.Drawing.Size(80, 30);
            this.btnConfirmarAnexo.TabIndex = 1;
            this.btnConfirmarAnexo.Text = "Confirmar";
            this.btnConfirmarAnexo.UseVisualStyleBackColor = false;
            this.btnConfirmarAnexo.Click += new System.EventHandler(this.btnConfirmarAnexo_Click);
            // 
            // btnCancelarAnexo
            // 
            this.btnCancelarAnexo.BackColor = System.Drawing.Color.LightCoral;
            this.btnCancelarAnexo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelarAnexo.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnCancelarAnexo.Location = new System.Drawing.Point(188, 150);
            this.btnCancelarAnexo.Name = "btnCancelarAnexo";
            this.btnCancelarAnexo.Size = new System.Drawing.Size(80, 30);
            this.btnCancelarAnexo.TabIndex = 2;
            this.btnCancelarAnexo.Text = "Cancelar";
            this.btnCancelarAnexo.UseVisualStyleBackColor = false;
            this.btnCancelarAnexo.Click += new System.EventHandler(this.btnCancelarAnexo_Click);
            // 
            // lblFolhaPagSS
            // 
            this.lblFolhaPagSS.AutoSize = true;
            this.lblFolhaPagSS.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblFolhaPagSS.Location = new System.Drawing.Point(251, 414);
            this.lblFolhaPagSS.Name = "lblFolhaPagSS";
            this.lblFolhaPagSS.Size = new System.Drawing.Size(78, 13);
            this.lblFolhaPagSS.TabIndex = 105;
            this.lblFolhaPagSS.Text = "Nenhum anexo";
            this.lblFolhaPagSS.Visible = false;
            this.lblFolhaPagSS.Click += new System.EventHandler(this.visualizarFolhaPag_Click);
            // 
            // btnAnexoFolhaPag
            // 
            this.btnAnexoFolhaPag.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoFolhaPag.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoFolhaPag.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoFolhaPag.Location = new System.Drawing.Point(197, 410);
            this.btnAnexoFolhaPag.Name = "btnAnexoFolhaPag";
            this.btnAnexoFolhaPag.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoFolhaPag.TabIndex = 104;
            this.btnAnexoFolhaPag.Text = "...";
            this.btnAnexoFolhaPag.UseVisualStyleBackColor = false;
            this.btnAnexoFolhaPag.Visible = false;
            this.btnAnexoFolhaPag.Click += new System.EventHandler(this.btnAnexoFolhaPag_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri", 9F);
            this.label5.Location = new System.Drawing.Point(25, 358);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(74, 14);
            this.label5.TabIndex = 59;
            this.label5.Text = "Não Div. Fin.";
            this.label5.Visible = false;
            // 
            // TXT_NaoDivFinancas
            // 
            this.TXT_NaoDivFinancas.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_NaoDivFinancas.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.TXT_NaoDivFinancas.Location = new System.Drawing.Point(105, 352);
            this.TXT_NaoDivFinancas.Name = "TXT_NaoDivFinancas";
            this.TXT_NaoDivFinancas.Size = new System.Drawing.Size(120, 22);
            this.TXT_NaoDivFinancas.TabIndex = 74;
            this.TXT_NaoDivFinancas.Visible = false;
            // 
            // btnAnexoFinancas
            // 
            this.btnAnexoFinancas.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoFinancas.Location = new System.Drawing.Point(197, 354);
            this.btnAnexoFinancas.Name = "btnAnexoFinancas";
            this.btnAnexoFinancas.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoFinancas.TabIndex = 100;
            this.btnAnexoFinancas.Text = "...";
            this.btnAnexoFinancas.UseVisualStyleBackColor = true;
            this.btnAnexoFinancas.Visible = false;
            this.btnAnexoFinancas.Click += new System.EventHandler(this.btnAnexoFinancas_Click);
            // 
            // lblAnexoFinancas
            // 
            this.lblAnexoFinancas.AutoSize = true;
            this.lblAnexoFinancas.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoFinancas.Location = new System.Drawing.Point(479, 350);
            this.lblAnexoFinancas.Name = "lblAnexoFinancas";
            this.lblAnexoFinancas.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoFinancas.TabIndex = 101;
            this.lblAnexoFinancas.Text = "Nenhum anexo";
            this.lblAnexoFinancas.Visible = false;
            this.lblAnexoFinancas.Click += new System.EventHandler(this.visualizarAnexoFinancas_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Calibri", 9F);
            this.label6.Location = new System.Drawing.Point(27, 386);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 14);
            this.label6.TabIndex = 60;
            this.label6.Text = "Não Div. S.S.";
            this.label6.Visible = false;
            // 
            // TXT_NaoDivSegSocial
            // 
            this.TXT_NaoDivSegSocial.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_NaoDivSegSocial.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.TXT_NaoDivSegSocial.Location = new System.Drawing.Point(105, 380);
            this.TXT_NaoDivSegSocial.Name = "TXT_NaoDivSegSocial";
            this.TXT_NaoDivSegSocial.Size = new System.Drawing.Size(120, 22);
            this.TXT_NaoDivSegSocial.TabIndex = 75;
            this.TXT_NaoDivSegSocial.Visible = false;
            // 
            // btnAnexoSegSocial
            // 
            this.btnAnexoSegSocial.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoSegSocial.Location = new System.Drawing.Point(197, 384);
            this.btnAnexoSegSocial.Name = "btnAnexoSegSocial";
            this.btnAnexoSegSocial.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoSegSocial.TabIndex = 102;
            this.btnAnexoSegSocial.Text = "...";
            this.btnAnexoSegSocial.UseVisualStyleBackColor = true;
            this.btnAnexoSegSocial.Visible = false;
            this.btnAnexoSegSocial.Click += new System.EventHandler(this.btnAnexoSegSocial_Click);
            // 
            // lblAnexoSegSocial
            // 
            this.lblAnexoSegSocial.AutoSize = true;
            this.lblAnexoSegSocial.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoSegSocial.Location = new System.Drawing.Point(479, 379);
            this.lblAnexoSegSocial.Name = "lblAnexoSegSocial";
            this.lblAnexoSegSocial.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoSegSocial.TabIndex = 103;
            this.lblAnexoSegSocial.Text = "Nenhum anexo";
            this.lblAnexoSegSocial.Visible = false;
            this.lblAnexoSegSocial.Click += new System.EventHandler(this.visualizarAnexoSegSocial_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri", 9F);
            this.label7.Location = new System.Drawing.Point(25, 414);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(81, 14);
            this.label7.TabIndex = 61;
            this.label7.Text = "Folha Pag. S.S";
            this.label7.Visible = false;
            // 
            // TXT_FolhaPagSegSocial
            // 
            this.TXT_FolhaPagSegSocial.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_FolhaPagSegSocial.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.TXT_FolhaPagSegSocial.Location = new System.Drawing.Point(105, 408);
            this.TXT_FolhaPagSegSocial.Name = "TXT_FolhaPagSegSocial";
            this.TXT_FolhaPagSegSocial.Size = new System.Drawing.Size(120, 22);
            this.TXT_FolhaPagSegSocial.TabIndex = 76;
            this.TXT_FolhaPagSegSocial.Visible = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Calibri", 9F);
            this.label8.Location = new System.Drawing.Point(453, 435);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(47, 14);
            this.label8.TabIndex = 62;
            this.label8.Text = "Recibo:";
            this.label8.Visible = false;
            // 
            // cb_ReciboPagSegSocial
            // 
            this.cb_ReciboPagSegSocial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_ReciboPagSegSocial.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_ReciboPagSegSocial.FormattingEnabled = true;
            this.cb_ReciboPagSegSocial.Location = new System.Drawing.Point(429, 376);
            this.cb_ReciboPagSegSocial.Name = "cb_ReciboPagSegSocial";
            this.cb_ReciboPagSegSocial.Size = new System.Drawing.Size(44, 22);
            this.cb_ReciboPagSegSocial.TabIndex = 81;
            this.cb_ReciboPagSegSocial.Visible = false;
            // 
            // labelCaminho
            // 
            this.labelCaminho.AutoSize = true;
            this.labelCaminho.Font = new System.Drawing.Font("Calibri", 9F);
            this.labelCaminho.Location = new System.Drawing.Point(9, 22);
            this.labelCaminho.Name = "labelCaminho";
            this.labelCaminho.Size = new System.Drawing.Size(58, 14);
            this.labelCaminho.TabIndex = 82;
            this.labelCaminho.Text = "Caminho:";
            // 
            // txtCaminhoPasta
            // 
            this.txtCaminhoPasta.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtCaminhoPasta.Location = new System.Drawing.Point(92, 20);
            this.txtCaminhoPasta.Name = "txtCaminhoPasta";
            this.txtCaminhoPasta.ReadOnly = true;
            this.txtCaminhoPasta.Size = new System.Drawing.Size(193, 22);
            this.txtCaminhoPasta.TabIndex = 83;
            // 
            // btnSelecionarPasta
            // 
            this.btnSelecionarPasta.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnSelecionarPasta.Location = new System.Drawing.Point(291, 20);
            this.btnSelecionarPasta.Name = "btnSelecionarPasta";
            this.btnSelecionarPasta.Size = new System.Drawing.Size(34, 22);
            this.btnSelecionarPasta.TabIndex = 84;
            this.btnSelecionarPasta.Text = "...";
            this.btnSelecionarPasta.UseVisualStyleBackColor = true;
            this.btnSelecionarPasta.Click += new System.EventHandler(this.btnSelecionarPasta_Click);
            // 
            // lblAnexoHorarioTrabalho
            // 
            this.lblAnexoHorarioTrabalho.AutoSize = true;
            this.lblAnexoHorarioTrabalho.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoHorarioTrabalho.Location = new System.Drawing.Point(566, 397);
            this.lblAnexoHorarioTrabalho.Name = "lblAnexoHorarioTrabalho";
            this.lblAnexoHorarioTrabalho.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoHorarioTrabalho.TabIndex = 110;
            this.lblAnexoHorarioTrabalho.Text = "Nenhum anexo";
            this.lblAnexoHorarioTrabalho.Visible = false;
            this.lblAnexoHorarioTrabalho.Click += new System.EventHandler(this.visualizarHorarioTrabalho_Click);
            // 
            // btnAnexoHorarioTrabalho
            // 
            this.btnAnexoHorarioTrabalho.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoHorarioTrabalho.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoHorarioTrabalho.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoHorarioTrabalho.Location = new System.Drawing.Point(513, 394);
            this.btnAnexoHorarioTrabalho.Name = "btnAnexoHorarioTrabalho";
            this.btnAnexoHorarioTrabalho.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoHorarioTrabalho.TabIndex = 111;
            this.btnAnexoHorarioTrabalho.Text = "...";
            this.btnAnexoHorarioTrabalho.UseVisualStyleBackColor = false;
            this.btnAnexoHorarioTrabalho.Visible = false;
            this.btnAnexoHorarioTrabalho.Click += new System.EventHandler(this.btnAnexoHorarioTrabalho_Click);
            // 
            // lblAnexoApoliceAT
            // 
            this.lblAnexoApoliceAT.AutoSize = true;
            this.lblAnexoApoliceAT.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoApoliceAT.Location = new System.Drawing.Point(566, 423);
            this.lblAnexoApoliceAT.Name = "lblAnexoApoliceAT";
            this.lblAnexoApoliceAT.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoApoliceAT.TabIndex = 106;
            this.lblAnexoApoliceAT.Text = "Nenhum anexo";
            this.lblAnexoApoliceAT.Visible = false;
            this.lblAnexoApoliceAT.Click += new System.EventHandler(this.visualizarApoliceAT_Click);
            // 
            // btnAnexoApoliceAT
            // 
            this.btnAnexoApoliceAT.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoApoliceAT.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoApoliceAT.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoApoliceAT.Location = new System.Drawing.Point(513, 420);
            this.btnAnexoApoliceAT.Name = "btnAnexoApoliceAT";
            this.btnAnexoApoliceAT.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoApoliceAT.TabIndex = 107;
            this.btnAnexoApoliceAT.Text = "...";
            this.btnAnexoApoliceAT.UseVisualStyleBackColor = false;
            this.btnAnexoApoliceAT.Click += new System.EventHandler(this.btnAnexoApoliceAT_Click);
            // 
            // lblAnexoApoliceRC
            // 
            this.lblAnexoApoliceRC.AutoSize = true;
            this.lblAnexoApoliceRC.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoApoliceRC.Location = new System.Drawing.Point(566, 449);
            this.lblAnexoApoliceRC.Name = "lblAnexoApoliceRC";
            this.lblAnexoApoliceRC.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoApoliceRC.TabIndex = 108;
            this.lblAnexoApoliceRC.Text = "Nenhum anexo";
            this.lblAnexoApoliceRC.Visible = false;
            this.lblAnexoApoliceRC.Click += new System.EventHandler(this.visualizarApoliceRC_Click);
            // 
            // btnAnexoApoliceRC
            // 
            this.btnAnexoApoliceRC.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoApoliceRC.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoApoliceRC.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoApoliceRC.Location = new System.Drawing.Point(513, 446);
            this.btnAnexoApoliceRC.Name = "btnAnexoApoliceRC";
            this.btnAnexoApoliceRC.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoApoliceRC.TabIndex = 109;
            this.btnAnexoApoliceRC.Text = "...";
            this.btnAnexoApoliceRC.UseVisualStyleBackColor = false;
            this.btnAnexoApoliceRC.Click += new System.EventHandler(this.btnAnexoApoliceRC_Click);
            // 
            // lblAnexoD
            // 
            this.lblAnexoD.AutoSize = true;
            this.lblAnexoD.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoD.Location = new System.Drawing.Point(566, 475);
            this.lblAnexoD.Name = "lblAnexoD";
            this.lblAnexoD.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoD.TabIndex = 112;
            this.lblAnexoD.Text = "Nenhum anexo";
            this.lblAnexoD.Visible = false;
            this.lblAnexoD.Click += new System.EventHandler(this.visualizarAnexoD_Click);
            // 
            // btnAnexoD
            // 
            this.btnAnexoD.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoD.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoD.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoD.Location = new System.Drawing.Point(513, 472);
            this.btnAnexoD.Name = "btnAnexoD";
            this.btnAnexoD.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoD.TabIndex = 113;
            this.btnAnexoD.Text = "...";
            this.btnAnexoD.UseVisualStyleBackColor = false;
            this.btnAnexoD.Click += new System.EventHandler(this.btnAnexoD_Click);
            // 
            // lblDecTrabEmigr
            // 
            this.lblDecTrabEmigr.AutoSize = true;
            this.lblDecTrabEmigr.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblDecTrabEmigr.Location = new System.Drawing.Point(566, 501);
            this.lblDecTrabEmigr.Name = "lblDecTrabEmigr";
            this.lblDecTrabEmigr.Size = new System.Drawing.Size(78, 13);
            this.lblDecTrabEmigr.TabIndex = 114;
            this.lblDecTrabEmigr.Text = "Nenhum anexo";
            this.lblDecTrabEmigr.Visible = false;
            this.lblDecTrabEmigr.Click += new System.EventHandler(this.visualizarDecTrabEmigr_Click);
            // 
            // btnDecTrabEmigr
            // 
            this.btnDecTrabEmigr.BackColor = System.Drawing.Color.LightBlue;
            this.btnDecTrabEmigr.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDecTrabEmigr.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnDecTrabEmigr.Location = new System.Drawing.Point(513, 498);
            this.btnDecTrabEmigr.Name = "btnDecTrabEmigr";
            this.btnDecTrabEmigr.Size = new System.Drawing.Size(47, 22);
            this.btnDecTrabEmigr.TabIndex = 115;
            this.btnDecTrabEmigr.Text = "...";
            this.btnDecTrabEmigr.UseVisualStyleBackColor = false;
            this.btnDecTrabEmigr.Click += new System.EventHandler(this.btnDecTrabEmigr_Click);
            // 
            // lblInscricaoSS
            // 
            this.lblInscricaoSS.AutoSize = true;
            this.lblInscricaoSS.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblInscricaoSS.Location = new System.Drawing.Point(584, 369);
            this.lblInscricaoSS.Name = "lblInscricaoSS";
            this.lblInscricaoSS.Size = new System.Drawing.Size(78, 13);
            this.lblInscricaoSS.TabIndex = 116;
            this.lblInscricaoSS.Text = "Nenhum anexo";
            this.lblInscricaoSS.Visible = false;
            this.lblInscricaoSS.Click += new System.EventHandler(this.visualizarInscricaoSS_Click);
            // 
            // btnInscricaoSS
            // 
            this.btnInscricaoSS.BackColor = System.Drawing.Color.LightBlue;
            this.btnInscricaoSS.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInscricaoSS.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnInscricaoSS.Location = new System.Drawing.Point(531, 366);
            this.btnInscricaoSS.Name = "btnInscricaoSS";
            this.btnInscricaoSS.Size = new System.Drawing.Size(47, 22);
            this.btnInscricaoSS.TabIndex = 117;
            this.btnInscricaoSS.Text = "...";
            this.btnInscricaoSS.UseVisualStyleBackColor = false;
            this.btnInscricaoSS.Click += new System.EventHandler(this.btnInscricaoSS_Click);
            // 
            // lblHorarioTrabalhoTitle
            // 
            this.lblHorarioTrabalhoTitle.AutoSize = true;
            this.lblHorarioTrabalhoTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblHorarioTrabalhoTitle.Location = new System.Drawing.Point(528, 323);
            this.lblHorarioTrabalhoTitle.Name = "lblHorarioTrabalhoTitle";
            this.lblHorarioTrabalhoTitle.Size = new System.Drawing.Size(102, 14);
            this.lblHorarioTrabalhoTitle.TabIndex = 118;
            this.lblHorarioTrabalhoTitle.Text = "Horário Trabalho:";
            this.lblHorarioTrabalhoTitle.Visible = false;
            // 
            // lblApoliceATTitle
            // 
            this.lblApoliceATTitle.AutoSize = true;
            this.lblApoliceATTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblApoliceATTitle.Location = new System.Drawing.Point(561, 332);
            this.lblApoliceATTitle.Name = "lblApoliceATTitle";
            this.lblApoliceATTitle.Size = new System.Drawing.Size(65, 14);
            this.lblApoliceATTitle.TabIndex = 119;
            this.lblApoliceATTitle.Text = "Apólice AT:";
            this.lblApoliceATTitle.Visible = false;
            // 
            // lblApoliceRCTitle
            // 
            this.lblApoliceRCTitle.AutoSize = true;
            this.lblApoliceRCTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblApoliceRCTitle.Location = new System.Drawing.Point(561, 358);
            this.lblApoliceRCTitle.Name = "lblApoliceRCTitle";
            this.lblApoliceRCTitle.Size = new System.Drawing.Size(67, 14);
            this.lblApoliceRCTitle.TabIndex = 120;
            this.lblApoliceRCTitle.Text = "Apólice RC:";
            this.lblApoliceRCTitle.Visible = false;
            // 
            // lblAnexoDTitle
            // 
            this.lblAnexoDTitle.AutoSize = true;
            this.lblAnexoDTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblAnexoDTitle.Location = new System.Drawing.Point(561, 384);
            this.lblAnexoDTitle.Name = "lblAnexoDTitle";
            this.lblAnexoDTitle.Size = new System.Drawing.Size(54, 14);
            this.lblAnexoDTitle.TabIndex = 121;
            this.lblAnexoDTitle.Text = "Anexo D:";
            this.lblAnexoDTitle.Visible = false;
            // 
            // lblDecTrabEmigrTitle
            // 
            this.lblDecTrabEmigrTitle.AutoSize = true;
            this.lblDecTrabEmigrTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblDecTrabEmigrTitle.Location = new System.Drawing.Point(332, 419);
            this.lblDecTrabEmigrTitle.Name = "lblDecTrabEmigrTitle";
            this.lblDecTrabEmigrTitle.Size = new System.Drawing.Size(97, 14);
            this.lblDecTrabEmigrTitle.TabIndex = 122;
            this.lblDecTrabEmigrTitle.Text = "Dec. Trab. Emigr.:";
            this.lblDecTrabEmigrTitle.Visible = false;
            // 
            // lblInscricaoSSTitle
            // 
            this.lblInscricaoSSTitle.AutoSize = true;
            this.lblInscricaoSSTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblInscricaoSSTitle.Location = new System.Drawing.Point(332, 445);
            this.lblInscricaoSSTitle.Name = "lblInscricaoSSTitle";
            this.lblInscricaoSSTitle.Size = new System.Drawing.Size(74, 14);
            this.lblInscricaoSSTitle.TabIndex = 123;
            this.lblInscricaoSSTitle.Text = "Inscrição SS:";
            this.lblInscricaoSSTitle.Visible = false;
            // 
            // btnAnexarDocumentoGeral
            // 
            this.btnAnexarDocumentoGeral.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnAnexarDocumentoGeral.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexarDocumentoGeral.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnAnexarDocumentoGeral.Location = new System.Drawing.Point(12, 78);
            this.btnAnexarDocumentoGeral.Name = "btnAnexarDocumentoGeral";
            this.btnAnexarDocumentoGeral.Size = new System.Drawing.Size(130, 24);
            this.btnAnexarDocumentoGeral.TabIndex = 124;
            this.btnAnexarDocumentoGeral.Text = "Anexar Documento";
            this.btnAnexarDocumentoGeral.UseVisualStyleBackColor = false;
            this.btnAnexarDocumentoGeral.Click += new System.EventHandler(this.btnAnexarDocumentoGeral_Click);
            // 
            // btnVerificarDocumentosFaltantes
            // 
            this.btnVerificarDocumentosFaltantes.BackColor = System.Drawing.Color.LightSalmon;
            this.btnVerificarDocumentosFaltantes.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnVerificarDocumentosFaltantes.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnVerificarDocumentosFaltantes.Location = new System.Drawing.Point(148, 48);
            this.btnVerificarDocumentosFaltantes.Name = "btnVerificarDocumentosFaltantes";
            this.btnVerificarDocumentosFaltantes.Size = new System.Drawing.Size(130, 24);
            this.btnVerificarDocumentosFaltantes.TabIndex = 125;
            this.btnVerificarDocumentosFaltantes.Text = "Documentos Faltantes";
            this.btnVerificarDocumentosFaltantes.UseVisualStyleBackColor = false;
            this.btnVerificarDocumentosFaltantes.Click += new System.EventHandler(this.btnVerificarDocumentosFaltantes_Click);
            // 
            // groupBoxApolices
            // 
            this.groupBoxApolices.Controls.Add(this.label9);
            this.groupBoxApolices.Controls.Add(this.cb_ApoliceAT);
            this.groupBoxApolices.Controls.Add(this.label10);
            this.groupBoxApolices.Controls.Add(this.TXT_ReciboApoliceAT);
            this.groupBoxApolices.Controls.Add(this.label4);
            this.label4.Controls.Add(this.TXT_Alvara);
            this.groupBoxApolices.Controls.Add(this.label11);
            this.groupBoxApolices.Controls.Add(this.cb_ApoliceRC);
            this.groupBoxApolices.Controls.Add(this.label22);
            this.groupBoxApolices.Controls.Add(this.TXT_AlvaraValidade);
            this.groupBoxApolices.Controls.Add(this.label12);
            this.groupBoxApolices.Controls.Add(this.AlertaValidadeAlvara);
            this.groupBoxApolices.Controls.Add(this.TXT_ReciboRC);
            this.groupBoxApolices.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxApolices.Location = new System.Drawing.Point(345, 5);
            this.groupBoxApolices.Name = "groupBoxApolices";
            this.groupBoxApolices.Size = new System.Drawing.Size(325, 132);
            this.groupBoxApolices.TabIndex = 2;
            this.groupBoxApolices.TabStop = false;
            this.groupBoxApolices.Text = "Apólices de Seguro";
            this.groupBoxApolices.Visible = false;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Calibri", 9F);
            this.label9.Location = new System.Drawing.Point(20, 22);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(63, 14);
            this.label9.TabIndex = 63;
            this.label9.Text = "Apólice AT";
            // 
            // cb_ApoliceAT
            // 
            this.cb_ApoliceAT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_ApoliceAT.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_ApoliceAT.FormattingEnabled = true;
            this.cb_ApoliceAT.Location = new System.Drawing.Point(88, 19);
            this.cb_ApoliceAT.Name = "cb_ApoliceAT";
            this.cb_ApoliceAT.Size = new System.Drawing.Size(50, 22);
            this.cb_ApoliceAT.TabIndex = 82;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Calibri", 9F);
            this.label10.Location = new System.Drawing.Point(143, 22);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(47, 14);
            this.label10.TabIndex = 64;
            this.label10.Text = "Recibo:";
            // 
            // TXT_ReciboApoliceAT
            // 
            this.TXT_ReciboApoliceAT.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_ReciboApoliceAT.Location = new System.Drawing.Point(198, 19);
            this.TXT_ReciboApoliceAT.Name = "TXT_ReciboApoliceAT";
            this.TXT_ReciboApoliceAT.Size = new System.Drawing.Size(113, 22);
            this.TXT_ReciboApoliceAT.TabIndex = 77;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri", 9F);
            this.label4.Location = new System.Drawing.Point(68, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 14);
            this.label4.TabIndex = 58;
            this.label4.Text = "Alvará";
            this.label4.Visible = false;
            // 
            // TXT_Alvara
            // 
            this.TXT_Alvara.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_Alvara.Location = new System.Drawing.Point(115, 71);
            this.TXT_Alvara.Name = "TXT_Alvara";
            this.TXT_Alvara.Size = new System.Drawing.Size(245, 22);
            this.TXT_Alvara.TabIndex = 73;
            this.TXT_Alvara.Visible = false;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Calibri", 9F);
            this.label11.Location = new System.Drawing.Point(19, 52);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(64, 14);
            this.label11.TabIndex = 65;
            this.label11.Text = "Apólice RC";
            // 
            // cb_ApoliceRC
            // 
            this.cb_ApoliceRC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_ApoliceRC.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_ApoliceRC.FormattingEnabled = true;
            this.cb_ApoliceRC.Items.AddRange(new object[] {
            "C",
            "N/C",
            "N/A"});
            this.cb_ApoliceRC.Location = new System.Drawing.Point(88, 49);
            this.cb_ApoliceRC.Name = "cb_ApoliceRC";
            this.cb_ApoliceRC.Size = new System.Drawing.Size(50, 22);
            this.cb_ApoliceRC.TabIndex = 83;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Calibri", 9F);
            this.label22.Location = new System.Drawing.Point(27, 105);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(56, 14);
            this.label22.TabIndex = 80;
            this.label22.Text = "Validade";
            this.label22.Visible = false;
            // 
            // TXT_AlvaraValidade
            // 
            this.TXT_AlvaraValidade.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_AlvaraValidade.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.TXT_AlvaraValidade.Location = new System.Drawing.Point(88, 102);
            this.TXT_AlvaraValidade.Name = "TXT_AlvaraValidade";
            this.TXT_AlvaraValidade.ShowCheckBox = true;
            this.TXT_AlvaraValidade.Size = new System.Drawing.Size(228, 22);
            this.TXT_AlvaraValidade.TabIndex = 79;
            this.TXT_AlvaraValidade.Visible = false;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Calibri", 9F);
            this.label12.Location = new System.Drawing.Point(143, 52);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(47, 14);
            this.label12.TabIndex = 66;
            this.label12.Text = "Recibo:";
            // 
            // AlertaValidadeAlvara
            // 
            this.AlertaValidadeAlvara.Location = new System.Drawing.Point(322, 102);
            this.AlertaValidadeAlvara.Name = "AlertaValidadeAlvara";
            this.AlertaValidadeAlvara.Size = new System.Drawing.Size(10, 10);
            this.AlertaValidadeAlvara.TabIndex = 88;
            this.AlertaValidadeAlvara.Visible = false;
            // 
            // TXT_ReciboRC
            // 
            this.TXT_ReciboRC.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_ReciboRC.Location = new System.Drawing.Point(198, 49);
            this.TXT_ReciboRC.Name = "TXT_ReciboRC";
            this.TXT_ReciboRC.Size = new System.Drawing.Size(113, 22);
            this.TXT_ReciboRC.TabIndex = 78;
            // 
            // groupBoxDeclaracoes
            // 
            this.groupBoxDeclaracoes.Controls.Add(this.label13);
            this.groupBoxDeclaracoes.Controls.Add(this.cb_HorarioTrabalho);
            this.groupBoxDeclaracoes.Controls.Add(this.label14);
            this.groupBoxDeclaracoes.Controls.Add(this.cb_DecTrabIlegais);
            this.groupBoxDeclaracoes.Controls.Add(this.label15);
            this.groupBoxDeclaracoes.Controls.Add(this.cb_DecRespEstaleiro);
            this.groupBoxDeclaracoes.Controls.Add(this.label16);
            this.groupBoxDeclaracoes.Controls.Add(this.cb_DecConhecimPSS);
            this.groupBoxDeclaracoes.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxDeclaracoes.Location = new System.Drawing.Point(11, 399);
            this.groupBoxDeclaracoes.Name = "groupBoxDeclaracoes";
            this.groupBoxDeclaracoes.Size = new System.Drawing.Size(662, 122);
            this.groupBoxDeclaracoes.TabIndex = 3;
            this.groupBoxDeclaracoes.TabStop = false;
            this.groupBoxDeclaracoes.Text = "Declarações";
            this.groupBoxDeclaracoes.Visible = false;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Calibri", 9F);
            this.label13.Location = new System.Drawing.Point(9, 22);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(116, 14);
            this.label13.TabIndex = 67;
            this.label13.Text = "Horário de Trabalho";
            // 
            // cb_HorarioTrabalho
            // 
            this.cb_HorarioTrabalho.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_HorarioTrabalho.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_HorarioTrabalho.FormattingEnabled = true;
            this.cb_HorarioTrabalho.Location = new System.Drawing.Point(131, 19);
            this.cb_HorarioTrabalho.Name = "cb_HorarioTrabalho";
            this.cb_HorarioTrabalho.Size = new System.Drawing.Size(180, 22);
            this.cb_HorarioTrabalho.TabIndex = 84;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Calibri", 9F);
            this.label14.Location = new System.Drawing.Point(22, 47);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(100, 14);
            this.label14.TabIndex = 68;
            this.label14.Text = "Dec. Trab. Ilegais";
            // 
            // cb_DecTrabIlegais
            // 
            this.cb_DecTrabIlegais.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_DecTrabIlegais.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_DecTrabIlegais.FormattingEnabled = true;
            this.cb_DecTrabIlegais.Location = new System.Drawing.Point(131, 44);
            this.cb_DecTrabIlegais.Name = "cb_DecTrabIlegais";
            this.cb_DecTrabIlegais.Size = new System.Drawing.Size(180, 22);
            this.cb_DecTrabIlegais.TabIndex = 85;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Calibri", 9F);
            this.label15.Location = new System.Drawing.Point(10, 70);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(115, 14);
            this.label15.TabIndex = 69;
            this.label15.Text = "Dec. Resp. Estaleiro";
            // 
            // cb_DecRespEstaleiro
            // 
            this.cb_DecRespEstaleiro.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_DecRespEstaleiro.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_DecRespEstaleiro.FormattingEnabled = true;
            this.cb_DecRespEstaleiro.Location = new System.Drawing.Point(131, 67);
            this.cb_DecRespEstaleiro.Name = "cb_DecRespEstaleiro";
            this.cb_DecRespEstaleiro.Size = new System.Drawing.Size(180, 22);
            this.cb_DecRespEstaleiro.TabIndex = 86;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Calibri", 9F);
            this.label16.Location = new System.Drawing.Point(12, 94);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(110, 14);
            this.label16.TabIndex = 70;
            this.label16.Text = "Dec. Conhecim. PSS";
            // 
            // cb_DecConhecimPSS
            // 
            this.cb_DecConhecimPSS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_DecConhecimPSS.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_DecConhecimPSS.FormattingEnabled = true;
            this.cb_DecConhecimPSS.Location = new System.Drawing.Point(131, 91);
            this.cb_DecConhecimPSS.Name = "cb_DecConhecimPSS";
            this.cb_DecConhecimPSS.Size = new System.Drawing.Size(180, 22);
            this.cb_DecConhecimPSS.TabIndex = 87;
            // 
            // panelObras
            // 
            this.panelObras.Controls.Add(this.groupBoxObras);
            this.panelObras.Controls.Add(this.pnlAutorizacaoObra);
            this.panelObras.Location = new System.Drawing.Point(3, 370);
            this.panelObras.Name = "panelObras";
            this.panelObras.Size = new System.Drawing.Size(712, 329);
            this.panelObras.TabIndex = 93;
            // 
            // groupBoxObras
            // 
            this.groupBoxObras.Controls.Add(this.label17);
            this.groupBoxObras.Controls.Add(this.cb_Obras);
            this.groupBoxObras.Controls.Add(this.btnGravarObra);
            this.groupBoxObras.Controls.Add(this.dataGridView1);
            this.groupBoxObras.Controls.Add(this.btnAutorizarEntrada);
            this.groupBoxObras.Controls.Add(this.lblDataEntrada);
            this.groupBoxObras.Controls.Add(this.dtpDataEntrada);
            this.groupBoxObras.Controls.Add(this.lblDataSaida);
            this.groupBoxObras.Controls.Add(this.dtpDataSaida);
            this.groupBoxObras.Controls.Add(this.lblContratoSubempreitada);
            this.groupBoxObras.Controls.Add(this.txtContratoSubempreitada);
            this.groupBoxObras.Controls.Add(this.pnlDadosObra);
            this.groupBoxObras.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxObras.Location = new System.Drawing.Point(0, 3);
            this.groupBoxObras.Name = "groupBoxObras";
            this.groupBoxObras.Size = new System.Drawing.Size(705, 319);
            this.groupBoxObras.TabIndex = 0;
            this.groupBoxObras.TabStop = false;
            this.groupBoxObras.Text = "Obras";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Calibri", 9F);
            this.label17.Location = new System.Drawing.Point(11, 24);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(103, 14);
            this.label17.TabIndex = 93;
            this.label17.Text = "Selecione a Obra:";
            // 
            // cb_Obras
            // 
            this.cb_Obras.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_Obras.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_Obras.FormattingEnabled = true;
            this.cb_Obras.Location = new System.Drawing.Point(123, 21);
            this.cb_Obras.Name = "cb_Obras";
            this.cb_Obras.Size = new System.Drawing.Size(470, 22);
            this.cb_Obras.TabIndex = 89;
            this.cb_Obras.SelectedIndexChanged += new System.EventHandler(this.cb_Obras_SelectedIndexChanged);
            // 
            // btnGravarObra
            // 
            this.btnGravarObra.BackColor = System.Drawing.Color.LightGreen;
            this.btnGravarObra.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGravarObra.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnGravarObra.Location = new System.Drawing.Point(599, 21);
            this.btnGravarObra.Name = "btnGravarObra";
            this.btnGravarObra.Size = new System.Drawing.Size(70, 22);
            this.btnGravarObra.TabIndex = 91;
            this.btnGravarObra.Text = "Gravar";
            this.btnGravarObra.UseVisualStyleBackColor = false;
            this.btnGravarObra.Click += new System.EventHandler(this.ProcessarGravarObra);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.BackgroundColor = System.Drawing.Color.WhiteSmoke;
            this.dataGridView1.ColumnHeadersHeight = 25;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.EntradaObra_,
            this.SaidaObra_,
            this.ContratoSubempreitada,
            this.StatusAutorizacao,
            this.AutorizacaoEntrada});
            this.dataGridView1.Location = new System.Drawing.Point(11, 110);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 25;
            this.dataGridView1.Size = new System.Drawing.Size(658, 75);
            this.dataGridView1.TabIndex = 90;
            // 
            // EntradaObra_
            // 
            this.EntradaObra_.HeaderText = "Entrada em Obra";
            this.EntradaObra_.Name = "EntradaObra_";
            // 
            // SaidaObra_
            // 
            this.SaidaObra_.HeaderText = "Saida de Obra";
            this.SaidaObra_.Name = "SaidaObra_";
            // 
            // ContratoSubempreitada
            // 
            this.ContratoSubempreitada.HeaderText = "Contrato Subempreitada";
            this.ContratoSubempreitada.Name = "ContratoSubempreitada";
            // 
            // StatusAutorizacao
            // 
            this.StatusAutorizacao.HeaderText = "Status";
            this.StatusAutorizacao.Name = "StatusAutorizacao";
            // 
            // AutorizacaoEntrada
            // 
            this.AutorizacaoEntrada.HeaderText = "Autorização de Entrada";
            this.AutorizacaoEntrada.Name = "AutorizacaoEntrada";
            // 
            // btnAutorizarEntrada
            // 
            this.btnAutorizarEntrada.BackColor = System.Drawing.Color.LightGreen;
            this.btnAutorizarEntrada.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAutorizarEntrada.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold);
            this.btnAutorizarEntrada.Location = new System.Drawing.Point(11, 49);
            this.btnAutorizarEntrada.Name = "btnAutorizarEntrada";
            this.btnAutorizarEntrada.Size = new System.Drawing.Size(180, 30);
            this.btnAutorizarEntrada.TabIndex = 92;
            this.btnAutorizarEntrada.Text = "Autorizar Nova Entrada em Obra";
            this.btnAutorizarEntrada.UseVisualStyleBackColor = false;
            this.btnAutorizarEntrada.Click += new System.EventHandler(this.btnAutorizarEntrada_Click);
            // 
            // lblDataEntrada
            // 
            this.lblDataEntrada.AutoSize = true;
            this.lblDataEntrada.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblDataEntrada.Location = new System.Drawing.Point(175, 55);
            this.lblDataEntrada.Name = "lblDataEntrada";
            this.lblDataEntrada.Size = new System.Drawing.Size(81, 14);
            this.lblDataEntrada.TabIndex = 95;
            this.lblDataEntrada.Text = "Data Entrada:";
            this.lblDataEntrada.Visible = false;
            // 
            // dtpDataEntrada
            // 
            this.dtpDataEntrada.Font = new System.Drawing.Font("Calibri", 9F);
            this.dtpDataEntrada.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataEntrada.Location = new System.Drawing.Point(260, 53);
            this.dtpDataEntrada.Name = "dtpDataEntrada";
            this.dtpDataEntrada.Size = new System.Drawing.Size(100, 22);
            this.dtpDataEntrada.TabIndex = 93;
            this.dtpDataEntrada.Visible = false;
            // 
            // lblDataSaida
            // 
            this.lblDataSaida.AutoSize = true;
            this.lblDataSaida.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblDataSaida.Location = new System.Drawing.Point(370, 55);
            this.lblDataSaida.Name = "lblDataSaida";
            this.lblDataSaida.Size = new System.Drawing.Size(69, 14);
            this.lblDataSaida.TabIndex = 96;
            this.lblDataSaida.Text = "Data Saída:";
            this.lblDataSaida.Visible = false;
            // 
            // dtpDataSaida
            // 
            this.dtpDataSaida.Font = new System.Drawing.Font("Calibri", 9F);
            this.dtpDataSaida.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataSaida.Location = new System.Drawing.Point(440, 53);
            this.dtpDataSaida.Name = "dtpDataSaida";
            this.dtpDataSaida.Size = new System.Drawing.Size(100, 22);
            this.dtpDataSaida.TabIndex = 94;
            this.dtpDataSaida.Visible = false;
            // 
            // lblContratoSubempreitada
            // 
            this.lblContratoSubempreitada.AutoSize = true;
            this.lblContratoSubempreitada.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblContratoSubempreitada.Location = new System.Drawing.Point(175, 82);
            this.lblContratoSubempreitada.Name = "lblContratoSubempreitada";
            this.lblContratoSubempreitada.Size = new System.Drawing.Size(143, 14);
            this.lblContratoSubempreitada.TabIndex = 97;
            this.lblContratoSubempreitada.Text = "Contrato Subempreitada:";
            this.lblContratoSubempreitada.Visible = false;
            // 
            // txtContratoSubempreitada
            // 
            this.txtContratoSubempreitada.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtContratoSubempreitada.Location = new System.Drawing.Point(325, 80);
            this.txtContratoSubempreitada.Name = "txtContratoSubempreitada";
            this.txtContratoSubempreitada.Size = new System.Drawing.Size(215, 22);
            this.txtContratoSubempreitada.TabIndex = 95;
            this.txtContratoSubempreitada.Visible = false;
            // 
            // pnlDadosObra
            // 
            this.pnlDadosObra.BackColor = System.Drawing.Color.WhiteSmoke;
            this.pnlDadosObra.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlDadosObra.Location = new System.Drawing.Point(170, 49);
            this.pnlDadosObra.Name = "pnlDadosObra";
            this.pnlDadosObra.Size = new System.Drawing.Size(499, 30);
            this.pnlDadosObra.TabIndex = 98;
            this.pnlDadosObra.Visible = false;
            // 
            // pnlAutorizacaoObra
            // 
            this.pnlAutorizacaoObra.BackColor = System.Drawing.Color.WhiteSmoke;
            this.pnlAutorizacaoObra.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlAutorizacaoObra.Controls.Add(this.lblAutorizacao);
            this.pnlAutorizacaoObra.Controls.Add(this.cmbAutorizacaoStatus);
            this.pnlAutorizacaoObra.Controls.Add(this.lblObservacao);
            this.pnlAutorizacaoObra.Controls.Add(this.txtObservacao);
            this.pnlAutorizacaoObra.Controls.Add(this.btnSalvarAutorizacao);
            this.pnlAutorizacaoObra.Location = new System.Drawing.Point(11, 385);
            this.pnlAutorizacaoObra.Name = "pnlAutorizacaoObra";
            this.pnlAutorizacaoObra.Size = new System.Drawing.Size(658, 45);
            this.pnlAutorizacaoObra.TabIndex = 94;
            this.pnlAutorizacaoObra.Visible = false;
            // 
            // lblAutorizacao
            // 
            this.lblAutorizacao.AutoSize = true;
            this.lblAutorizacao.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold);
            this.lblAutorizacao.Location = new System.Drawing.Point(10, 12);
            this.lblAutorizacao.Name = "lblAutorizacao";
            this.lblAutorizacao.Size = new System.Drawing.Size(50, 17);
            this.lblAutorizacao.TabIndex = 0;
            this.lblAutorizacao.Text = "Status:";
            // 
            // cmbAutorizacaoStatus
            // 
            this.cmbAutorizacaoStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAutorizacaoStatus.Font = new System.Drawing.Font("Calibri", 9.5F);
            this.cmbAutorizacaoStatus.Items.AddRange(new object[] {
            "Autorizado",
            "Pendente",
            "Não Autorizado",
            "Renovação Necessária",
            "Documentos Faltantes"});
            this.cmbAutorizacaoStatus.Location = new System.Drawing.Point(70, 10);
            this.cmbAutorizacaoStatus.Name = "cmbAutorizacaoStatus";
            this.cmbAutorizacaoStatus.Size = new System.Drawing.Size(150, 23);
            this.cmbAutorizacaoStatus.TabIndex = 1;
            this.cmbAutorizacaoStatus.SelectedIndexChanged += new System.EventHandler(this.cmbAutorizacaoStatus_SelectedIndexChanged);
            // 
            // lblObservacao
            // 
            this.lblObservacao.AutoSize = true;
            this.lblObservacao.Font = new System.Drawing.Font("Calibri", 9.5F);
            this.lblObservacao.Location = new System.Drawing.Point(230, 12);
            this.lblObservacao.Name = "lblObservacao";
            this.lblObservacao.Size = new System.Drawing.Size(32, 15);
            this.lblObservacao.TabIndex = 2;
            this.lblObservacao.Text = "Obs:";
            // 
            // txtObservacao
            // 
            this.txtObservacao.Font = new System.Drawing.Font("Calibri", 9.5F);
            this.txtObservacao.Location = new System.Drawing.Point(280, 10);
            this.txtObservacao.Name = "txtObservacao";
            this.txtObservacao.Size = new System.Drawing.Size(280, 23);
            this.txtObservacao.TabIndex = 3;
            this.toolTip.SetToolTip(this.txtObservacao, "Observações sobre a autorização...");
            // 
            // btnSalvarAutorizacao
            // 
            this.btnSalvarAutorizacao.BackColor = System.Drawing.Color.LightBlue;
            this.btnSalvarAutorizacao.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSalvarAutorizacao.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnSalvarAutorizacao.Location = new System.Drawing.Point(570, 10);
            this.btnSalvarAutorizacao.Name = "btnSalvarAutorizacao";
            this.btnSalvarAutorizacao.Size = new System.Drawing.Size(75, 25);
            this.btnSalvarAutorizacao.TabIndex = 4;
            this.btnSalvarAutorizacao.Text = "Salvar";
            this.btnSalvarAutorizacao.UseVisualStyleBackColor = false;
            this.btnSalvarAutorizacao.Click += new System.EventHandler(this.btnSalvarAutorizacao_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.panelTrabalhadores);
            this.tabPage2.Location = new System.Drawing.Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(718, 703);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Trabalhadores";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // panelTrabalhadores
            // 
            this.panelTrabalhadores.Controls.Add(this.groupBoxInfoTrabalhador);
            this.panelTrabalhadores.Controls.Add(this.groupBoxListaTrabalhadores);
            this.panelTrabalhadores.Controls.Add(this.panelBotoesTrabalhador);
            this.panelTrabalhadores.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelTrabalhadores.Location = new System.Drawing.Point(3, 3);
            this.panelTrabalhadores.Name = "panelTrabalhadores";
            this.panelTrabalhadores.Size = new System.Drawing.Size(712, 697);
            this.panelTrabalhadores.TabIndex = 0;
            // 
            // groupBoxInfoTrabalhador
            // 
            this.groupBoxInfoTrabalhador.Controls.Add(this.txtNomeTrabalhador);
            this.groupBoxInfoTrabalhador.Controls.Add(this.cmbTipoDocumentoTrabalhador);
            this.groupBoxInfoTrabalhador.Controls.Add(this.txtNumDocumento);
            this.groupBoxInfoTrabalhador.Controls.Add(this.dtpValidadeDocumento);
            this.groupBoxInfoTrabalhador.Controls.Add(this.txtNIF);
            this.groupBoxInfoTrabalhador.Controls.Add(this.txtNumSS);
            this.groupBoxInfoTrabalhador.Controls.Add(this.chkFichaAptidaoMedica);
            this.groupBoxInfoTrabalhador.Controls.Add(this.chkCredenciacao);
            this.groupBoxInfoTrabalhador.Controls.Add(this.txtCredenciacao);
            this.groupBoxInfoTrabalhador.Controls.Add(this.chkFichaEPI);
            this.groupBoxInfoTrabalhador.Controls.Add(this.lblFichaAptidaoAnexo);
            this.groupBoxInfoTrabalhador.Controls.Add(this.btnAnexarFichaAptidao);
            this.groupBoxInfoTrabalhador.Controls.Add(this.lblCredenciacaoAnexo);
            this.groupBoxInfoTrabalhador.Controls.Add(this.btnAnexarCredenciacao);
            this.groupBoxInfoTrabalhador.Controls.Add(this.lblFichaEPIAnexo);
            this.groupBoxInfoTrabalhador.Controls.Add(this.btnAnexarFichaEPI);
            this.groupBoxInfoTrabalhador.Location = new System.Drawing.Point(8, 5);
            this.groupBoxInfoTrabalhador.Name = "groupBoxInfoTrabalhador";
            this.groupBoxInfoTrabalhador.Size = new System.Drawing.Size(697, 200);
            this.groupBoxInfoTrabalhador.TabIndex = 0;
            this.groupBoxInfoTrabalhador.TabStop = false;
            this.groupBoxInfoTrabalhador.Text = "Informações do Trabalhador";
            // 
            // txtNomeTrabalhador
            // 
            this.txtNomeTrabalhador.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtNomeTrabalhador.Location = new System.Drawing.Point(70, 27);
            this.txtNomeTrabalhador.Name = "txtNomeTrabalhador";
            this.txtNomeTrabalhador.Size = new System.Drawing.Size(300, 22);
            this.txtNomeTrabalhador.TabIndex = 0;
            // 
            // cmbTipoDocumentoTrabalhador
            // 
            this.cmbTipoDocumentoTrabalhador.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTipoDocumentoTrabalhador.Font = new System.Drawing.Font("Calibri", 9F);
            this.cmbTipoDocumentoTrabalhador.FormattingEnabled = true;
            this.cmbTipoDocumentoTrabalhador.Location = new System.Drawing.Point(120, 57);
            this.cmbTipoDocumentoTrabalhador.Name = "cmbTipoDocumentoTrabalhador";
            this.cmbTipoDocumentoTrabalhador.Size = new System.Drawing.Size(150, 22);
            this.cmbTipoDocumentoTrabalhador.TabIndex = 1;
            this.cmbTipoDocumentoTrabalhador.Items.AddRange(new object[] {
            "Cartão Cidadão",
            "Bilhete de Identidade",
            "Passaporte",
            "Título de Residência",
            "Outro"});
            // 
            // txtNumDocumento
            // 
            this.txtNumDocumento.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtNumDocumento.Location = new System.Drawing.Point(340, 57);
            this.txtNumDocumento.Name = "txtNumDocumento";
            this.txtNumDocumento.Size = new System.Drawing.Size(150, 22);
            this.txtNumDocumento.TabIndex = 2;
            // 
            // dtpValidadeDocumento
            // 
            this.dtpValidadeDocumento.Font = new System.Drawing.Font("Calibri", 9F);
            this.dtpValidadeDocumento.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpValidadeDocumento.Location = new System.Drawing.Point(560, 57);
            this.dtpValidadeDocumento.Name = "dtpValidadeDocumento";
            this.dtpValidadeDocumento.Size = new System.Drawing.Size(120, 22);
            this.dtpValidadeDocumento.TabIndex = 3;
            this.dtpValidadeDocumento.ShowCheckBox = true;
            // 
            // txtNIF
            // 
            this.txtNIF.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtNIF.Location = new System.Drawing.Point(70, 87);
            this.txtNIF.Name = "txtNIF";
            this.txtNIF.Size = new System.Drawing.Size(150, 22);
            this.txtNIF.TabIndex = 4;
            // 
            // txtNumSS
            // 
            this.txtNumSS.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtNumSS.Location = new System.Drawing.Point(340, 87);
            this.txtNumSS.Name = "txtNumSS";
            this.txtNumSS.Size = new System.Drawing.Size(150, 22);
            this.txtNumSS.TabIndex = 5;
            // 
            // chkFichaAptidaoMedica
            // 
            this.chkFichaAptidaoMedica.AutoSize = true;
            this.chkFichaAptidaoMedica.Font = new System.Drawing.Font("Calibri", 9F);
            this.chkFichaAptidaoMedica.Location = new System.Drawing.Point(20, 120);
            this.chkFichaAptidaoMedica.Name = "chkFichaAptidaoMedica";
            this.chkFichaAptidaoMedica.Size = new System.Drawing.Size(160, 18);
            this.chkFichaAptidaoMedica.TabIndex = 6;
            this.chkFichaAptidaoMedica.Text = "Ficha de Aptidão Médica";
            this.chkFichaAptidaoMedica.UseVisualStyleBackColor = true;
            // 
            // chkCredenciacao
            // 
            this.chkCredenciacao.AutoSize = true;
            this.chkCredenciacao.Font = new System.Drawing.Font("Calibri", 9F);
            this.chkCredenciacao.Location = new System.Drawing.Point(20, 150);
            this.chkCredenciacao.Name = "chkCredenciacao";
            this.chkCredenciacao.Size = new System.Drawing.Size(98, 18);
            this.chkCredenciacao.TabIndex = 7;
            this.chkCredenciacao.Text = "Credenciação";
            this.chkCredenciacao.UseVisualStyleBackColor = true;
            // 
            // txtCredenciacao
            // 
            this.txtCredenciacao.Font = new System.Drawing.Font("Calibri", 9F);
            this.txtCredenciacao.Location = new System.Drawing.Point(120, 148);
            this.txtCredenciacao.Name = "txtCredenciacao";
            this.txtCredenciacao.Size = new System.Drawing.Size(180, 22);
            this.txtCredenciacao.TabIndex = 8;
            this.txtCredenciacao.Enabled = false;
            // 
            // chkFichaEPI
            // 
            this.chkFichaEPI.AutoSize = true;
            this.chkFichaEPI.Font = new System.Drawing.Font("Calibri", 9F);
            this.chkFichaEPI.Location = new System.Drawing.Point(20, 180);
            this.chkFichaEPI.Name = "chkFichaEPI";
            this.chkFichaEPI.Size = new System.Drawing.Size(167, 18);
            this.chkFichaEPI.TabIndex = 9;
            this.chkFichaEPI.Text = "Ficha de Distribuição de EPI";
            this.chkFichaEPI.UseVisualStyleBackColor = true;
            // 
            // gridTrabalhadores
            // 
            this.gridTrabalhadores.AllowUserToAddRows = false;
            this.gridTrabalhadores.AllowUserToDeleteRows = false;
            this.gridTrabalhadores.BackgroundColor = System.Drawing.Color.WhiteSmoke;
            this.gridTrabalhadores.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridTrabalhadores.Location = new System.Drawing.Point(10, 25);
            this.gridTrabalhadores.Name = "gridTrabalhadores";
            this.gridTrabalhadores.ReadOnly = true;
            this.gridTrabalhadores.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridTrabalhadores.Size = new System.Drawing.Size(677, 180);
            this.gridTrabalhadores.TabIndex = 0;
            // 
            // cmbObrasTrabalhador
            // 
            this.cmbObrasTrabalhador.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbObrasTrabalhador.Font = new System.Drawing.Font("Calibri", 9F);
            this.cmbObrasTrabalhador.FormattingEnabled = true;
            this.cmbObrasTrabalhador.Location = new System.Drawing.Point(120, 212);
            this.cmbObrasTrabalhador.Name = "cmbObrasTrabalhador";
            this.cmbObrasTrabalhador.Size = new System.Drawing.Size(400, 22);
            this.cmbObrasTrabalhador.TabIndex = 10;
            // 
            // lblFichaAptidaoAnexo
            // 
            this.lblFichaAptidaoAnexo.AutoSize = true;
            this.lblFichaAptidaoAnexo.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblFichaAptidaoAnexo.Location = new System.Drawing.Point(370, 121);
            this.lblFichaAptidaoAnexo.Name = "lblFichaAptidaoAnexo";
            this.lblFichaAptidaoAnexo.Size = new System.Drawing.Size(200, 13);
            this.lblFichaAptidaoAnexo.TabIndex = 11;
            this.lblFichaAptidaoAnexo.Text = "Ficha de Aptidão Médica:";
            this.lblFichaAptidaoAnexo.ForeColor = System.Drawing.Color.Blue;
            this.lblFichaAptidaoAnexo.Cursor = System.Windows.Forms.Cursors.Hand;
            // 
            // btnAnexarFichaAptidao
            // 
            this.btnAnexarFichaAptidao.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexarFichaAptidao.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexarFichaAptidao.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexarFichaAptidao.Location = new System.Drawing.Point(320, 119);
            this.btnAnexarFichaAptidao.Name = "btnAnexarFichaAptidao";
            this.btnAnexarFichaAptidao.Size = new System.Drawing.Size(40, 22);
            this.btnAnexarFichaAptidao.TabIndex = 12;
            this.btnAnexarFichaAptidao.Text = "...";
            this.btnAnexarFichaAptidao.UseVisualStyleBackColor = false;
            // 
            // lblCredenciacaoAnexo
            // 
            this.lblCredenciacaoAnexo.AutoSize = true;
            this.lblCredenciacaoAnexo.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblCredenciacaoAnexo.Location = new System.Drawing.Point(370, 151);
            this.lblCredenciacaoAnexo.Name = "lblCredenciacaoAnexo";
            this.lblCredenciacaoAnexo.Size = new System.Drawing.Size(100, 13);
            this.lblCredenciacaoAnexo.TabIndex = 13;
            this.lblCredenciacaoAnexo.Text = "Credenciação:";
            this.lblCredenciacaoAnexo.ForeColor = System.Drawing.Color.Blue;
            this.lblCredenciacaoAnexo.Cursor = System.Windows.Forms.Cursors.Hand;
            // 
            // btnAnexarCredenciacao
            // 
            this.btnAnexarCredenciacao.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexarCredenciacao.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexarCredenciacao.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexarCredenciacao.Location = new System.Drawing.Point(320, 149);
            this.btnAnexarCredenciacao.Name = "btnAnexarCredenciacao";
            this.btnAnexarCredenciacao.Size = new System.Drawing.Size(40, 22);
            this.btnAnexarCredenciacao.TabIndex = 14;
            this.btnAnexarCredenciacao.Text = "...";
            this.btnAnexarCredenciacao.UseVisualStyleBackColor = false;
            // 
            // lblFichaEPIAnexo
            // 
            this.lblFichaEPIAnexo.AutoSize = true;
            this.lblFichaEPIAnexo.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblFichaEPIAnexo.Location = new System.Drawing.Point(370, 181);
            this.lblFichaEPIAnexo.Name = "lblFichaEPIAnexo";
            this.lblFichaEPIAnexo.Size = new System.Drawing.Size(160, 13);
            this.lblFichaEPIAnexo.TabIndex = 15;
            this.lblFichaEPIAnexo.Text = "Ficha de Distribuição de EPI:";
            this.lblFichaEPIAnexo.ForeColor = System.Drawing.Color.Blue;
            this.lblFichaEPIAnexo.Cursor = System.Windows.Forms.Cursors.Hand;
            // 
            // btnAnexarFichaEPI
            // 
            this.btnAnexarFichaEPI.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexarFichaEPI.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexarFichaEPI.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexarFichaEPI.Location = new System.Drawing.Point(320, 179);
            this.btnAnexarFichaEPI.Name = "btnAnexarFichaEPI";
            this.btnAnexarFichaEPI.Size = new System.Drawing.Size(40, 22);
            this.btnAnexarFichaEPI.TabIndex = 16;
            this.btnAnexarFichaEPI.Text = "...";
            this.btnAnexarFichaEPI.UseVisualStyleBackColor = false;
            // 
            // groupBoxListaTrabalhadores
            // 
            this.groupBoxListaTrabalhadores.Controls.Add(this.gridTrabalhadores);
            this.groupBoxListaTrabalhadores.Controls.Add(this.cmbObrasTrabalhador);
            this.groupBoxListaTrabalhadores.Location = new System.Drawing.Point(8, 210);
            this.groupBoxListaTrabalhadores.Name = "groupBoxListaTrabalhadores";
            this.groupBoxListaTrabalhadores.Size = new System.Drawing.Size(697, 265);
            this.groupBoxListaTrabalhadores.TabIndex = 1;
            this.groupBoxListaTrabalhadores.TabStop = false;
            this.groupBoxListaTrabalhadores.Text = "Lista de Trabalhadores";
            // 
            // panelBotoesTrabalhador
            // 
            this.panelBotoesTrabalhador.Controls.Add(this.btnAdicionarTrabalhador);
            this.panelBotoesTrabalhador.Controls.Add(this.btnEditarTrabalhador);
            this.panelBotoesTrabalhador.Controls.Add(this.btnExcluirTrabalhador);
            this.panelBotoesTrabalhador.Controls.Add(this.btnSalvarTrabalhador);
            this.panelBotoesTrabalhador.Controls.Add(this.btnAutorizarTrabalhador);
            this.panelBotoesTrabalhador.Location = new System.Drawing.Point(8, 480);
            this.panelBotoesTrabalhador.Name = "panelBotoesTrabalhador";
            this.panelBotoesTrabalhador.Size = new System.Drawing.Size(697, 50);
            this.panelBotoesTrabalhador.TabIndex = 2;
            // 
            // btnAdicionarTrabalhador
            // 
            this.btnAdicionarTrabalhador.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnAdicionarTrabalhador.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdicionarTrabalhador.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnAdicionarTrabalhador.Location = new System.Drawing.Point(10, 10);
            this.btnAdicionarTrabalhador.Name = "btnAdicionarTrabalhador";
            this.btnAdicionarTrabalhador.Size = new System.Drawing.Size(100, 30);
            this.btnAdicionarTrabalhador.TabIndex = 0;
            this.btnAdicionarTrabalhador.Text = "Adicionar";
            this.btnAdicionarTrabalhador.UseVisualStyleBackColor = false;
            // 
            // btnEditarTrabalhador
            // 
            this.btnEditarTrabalhador.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnEditarTrabalhador.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEditarTrabalhador.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnEditarTrabalhador.Location = new System.Drawing.Point(120, 10);
            this.btnEditarTrabalhador.Name = "btnEditarTrabalhador";
            this.btnEditarTrabalhador.Size = new System.Drawing.Size(100, 30);
            this.btnEditarTrabalhador.TabIndex = 1;
            this.btnEditarTrabalhador.Text = "Editar";
            this.btnEditarTrabalhador.UseVisualStyleBackColor = false;
            // 
            // btnExcluirTrabalhador
            // 
            this.btnExcluirTrabalhador.BackColor = System.Drawing.Color.LightCoral;
            this.btnExcluirTrabalhador.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExcluirTrabalhador.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnExcluirTrabalhador.Location = new System.Drawing.Point(230, 10);
            this.btnExcluirTrabalhador.Name = "btnExcluirTrabalhador";
            this.btnExcluirTrabalhador.Size = new System.Drawing.Size(100, 30);
            this.btnExcluirTrabalhador.TabIndex = 2;
            this.btnExcluirTrabalhador.Text = "Excluir";
            this.btnExcluirTrabalhador.UseVisualStyleBackColor = false;
            // 
            // btnSalvarTrabalhador
            // 
            this.btnSalvarTrabalhador.BackColor = System.Drawing.Color.LightGreen;
            this.btnSalvarTrabalhador.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSalvarTrabalhador.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnSalvarTrabalhador.Location = new System.Drawing.Point(340, 10);
            this.btnSalvarTrabalhador.Name = "btnSalvarTrabalhador";
            this.btnSalvarTrabalhador.Size = new System.Drawing.Size(100, 30);
            this.btnSalvarTrabalhador.TabIndex = 3;
            this.btnSalvarTrabalhador.Text = "Salvar";
            this.btnSalvarTrabalhador.UseVisualStyleBackColor = false;
            // 
            // btnAutorizarTrabalhador
            // 
            this.btnAutorizarTrabalhador.BackColor = System.Drawing.Color.PaleGoldenrod;
            this.btnAutorizarTrabalhador.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAutorizarTrabalhador.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnAutorizarTrabalhador.Location = new System.Drawing.Point(450, 10);
            this.btnAutorizarTrabalhador.Name = "btnAutorizarTrabalhador";
            this.btnAutorizarTrabalhador.Size = new System.Drawing.Size(150, 30);
            this.btnAutorizarTrabalhador.TabIndex = 4;
            this.btnAutorizarTrabalhador.Text = "Autorizar para a Obra";
            this.btnAutorizarTrabalhador.UseVisualStyleBackColor = false;
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.BT_Salvar_Click});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(757, 25);
            this.toolStrip1.TabIndex = 4;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // BT_Salvar_Click
            // 
            this.BT_Salvar_Click.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.BT_Salvar_Click.Image = ((System.Drawing.Image)(resources.GetObject("BT_Salvar_Click.Image")));
            this.BT_Salvar_Click.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.BT_Salvar_Click.Name = "BT_Salvar_Click";
            this.BT_Salvar_Click.Size = new System.Drawing.Size(23, 22);
            this.BT_Salvar_Click.Text = "Gravar Alterações";
            this.BT_Salvar_Click.ToolTipText = "Gravar Alterações";
            this.BT_Salvar_Click.Click += new System.EventHandler(this.BT_Salvar_Click_Click);
            // 
            // lblSelecionar
            // 
            this.lblSelecionar.AutoSize = true;
            this.lblSelecionar.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.lblSelecionar.Location = new System.Drawing.Point(15, 40);
            this.lblSelecionar.Name = "lblSelecionar";
            this.lblSelecionar.Size = new System.Drawing.Size(43, 14);
            this.lblSelecionar.TabIndex = 5;
            this.lblSelecionar.Text = "Código:";
            // 
            // lblStatusEntrada
            // 
            this.lblStatusEntrada.AutoSize = true;
            this.lblStatusEntrada.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblStatusEntrada.Location = new System.Drawing.Point(175, 110);
            this.lblStatusEntrada.Name = "lblStatusEntrada";
            this.lblStatusEntrada.Size = new System.Drawing.Size(44, 14);
            this.lblStatusEntrada.TabIndex = 0;
            this.lblStatusEntrada.Text = "Status:";
            this.lblStatusEntrada.Visible = false;
            // 
            // cmbStatusEntrada
            // 
            this.cmbStatusEntrada.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbStatusEntrada.Font = new System.Drawing.Font("Calibri", 9F);
            this.cmbStatusEntrada.Items.AddRange(new object[] {
            "Autorizado",
            "Pendente",
            "Não Autorizado",
            "Renovação Necessária",
            "Documentos Faltantes"});
            this.cmbStatusEntrada.Location = new System.Drawing.Point(325, 108);
            this.cmbStatusEntrada.Name = "cmbStatusEntrada";
            this.cmbStatusEntrada.Size = new System.Drawing.Size(215, 22);
            this.cmbStatusEntrada.TabIndex = 96;
            this.cmbStatusEntrada.Visible = false;
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(757, 804);
            this.Controls.Add(this.lblSelecionar);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.TXT_Codigo);
            this.Controls.Add(this.BTF4);
            this.Controls.Add(this.TXT_Nome);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
    
            this.Name = "Menu";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestão de Subempreiteiros";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panelDadosEmpresa.ResumeLayout(false);
            this.groupBoxInfoBasica.ResumeLayout(false);
            this.groupBoxInfoBasica.PerformLayout();
            this.groupBoxSituacaoFiscal.ResumeLayout(false);
            this.groupBoxSituacaoFiscal.PerformLayout();
            this.panelModalDocumentos.ResumeLayout(false);
            this.panelModalDocumentos.PerformLayout();
            this.groupBoxApolices.ResumeLayout(false);
            this.groupBoxApolices.PerformLayout();
            this.groupBoxDeclaracoes.ResumeLayout(false);
            this.groupBoxDeclaracoes.PerformLayout();
            this.panelObras.ResumeLayout(false);
            this.groupBoxObras.ResumeLayout(false);
            this.groupBoxObras.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.pnlAutorizacaoObra.ResumeLayout(false);
            this.pnlAutorizacaoObra.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.panelTrabalhadores.ResumeLayout(false);
            this.groupBoxInfoTrabalhador.ResumeLayout(false);
            this.groupBoxInfoTrabalhador.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridTrabalhadores)).EndInit();
            this.groupBoxListaTrabalhadores.ResumeLayout(false);
            this.groupBoxListaTrabalhadores.PerformLayout();
            this.panelBotoesTrabalhador.ResumeLayout(false);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private GroupBox groupBoxApolices;
        private Label label9;
        private ComboBox cb_ApoliceAT;
        private Label label10;
        private TextBox TXT_ReciboApoliceAT;
        private Label label11;
        private ComboBox cb_ApoliceRC;
        private Label label12;
        private TextBox TXT_ReciboRC;
        private GroupBox groupBoxDeclaracoes;
        private Label label13;
        private ComboBox cb_HorarioTrabalho;
        private Label label14;
        private ComboBox cb_DecTrabIlegais;
        private Label label15;
        private ComboBox cb_DecRespEstaleiro;
        private Label label16;
        private ComboBox cb_DecConhecimPSS;
        private Label label1;
        private TextBox textBox1;
        private Label lblAutorizacao;
        private ComboBox cmbAutorizacaoStatus;
        private Label lblObservacao;
        private TextBox txtObservacao;
        private ToolTip toolTip;
        private Button btnSalvarAutorizacao;
    }
}