
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
        private ComboBox cb_DecConhecimPSS;
        private ComboBox cb_DecRespEstaleiro;
        private ComboBox cb_DecTrabIlegais;
        private ComboBox cb_HorarioTrabalho;
        private ComboBox cb_ApoliceRC;
        private ComboBox cb_ApoliceAT;
        private ComboBox cb_ReciboPagSegSocial;
        private Label label22;
        private DateTimePicker TXT_AlvaraValidade;
        private TextBox TXT_ReciboRC;
        private TextBox TXT_ReciboApoliceAT;
        private DateTimePicker TXT_FolhaPagSegSocial;
        private DateTimePicker TXT_NaoDivSegSocial;
        private DateTimePicker TXT_NaoDivFinancas;
        private TextBox TXT_Alvara;
        private TextBox TXT_Contribuinte;
        private TextBox TXT_Sede;
        private Label label16;
        private Label label15;
        private Label label14;
        private Label label13;
        private Label label12;
        private Label label11;
        private Label label10;
        private Label label9;
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
        private DataGridViewCheckBoxColumn AutorizacaoEntrada;
        private Panel panelDadosEmpresa;
        private GroupBox groupBoxInfoBasica;
        private GroupBox groupBoxSituacaoFiscal;
        private GroupBox groupBoxApolices;
        private GroupBox groupBoxDeclaracoes;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Menu));
            this.TXT_Nome = new System.Windows.Forms.TextBox();
            this.BTF4 = new System.Windows.Forms.Button();
            this.TXT_Codigo = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panelDadosEmpresa = new System.Windows.Forms.Panel();
            this.groupBoxInfoBasica = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TXT_Sede = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TXT_Contribuinte = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.TXT_Alvara = new System.Windows.Forms.TextBox();
            this.label22 = new System.Windows.Forms.Label();
            this.TXT_AlvaraValidade = new System.Windows.Forms.DateTimePicker();
            this.AlertaValidadeAlvara = new System.Windows.Forms.Panel();
            this.groupBoxSituacaoFiscal = new System.Windows.Forms.GroupBox();
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
            this.panelModalDocumentos = new System.Windows.Forms.Panel();
            this.cmbTipoDocumento = new System.Windows.Forms.ComboBox();
            this.btnConfirmarAnexo = new System.Windows.Forms.Button();
            this.btnCancelarAnexo = new System.Windows.Forms.Button();
            this.lblTipoDocumento = new System.Windows.Forms.Label();
            this.groupBoxApolices = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.cb_ApoliceAT = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.TXT_ReciboApoliceAT = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.cb_ApoliceRC = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
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
            this.AutorizacaoEntrada = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.BT_Salvar_Click = new System.Windows.Forms.ToolStripButton();
            this.lblSelecionar = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panelDadosEmpresa.SuspendLayout();
            this.groupBoxInfoBasica.SuspendLayout();
            this.groupBoxSituacaoFiscal.SuspendLayout();
            this.groupBoxApolices.SuspendLayout();
            this.groupBoxDeclaracoes.SuspendLayout();
            this.panelObras.SuspendLayout();
            this.groupBoxObras.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
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
            this.tabControl1.Size = new System.Drawing.Size(726, 720);
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
            this.tabPage1.Size = new System.Drawing.Size(718, 692);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Empresa";
            // 
            // panelDadosEmpresa
            // 
            this.panelDadosEmpresa.Controls.Add(this.groupBoxInfoBasica);
            this.panelDadosEmpresa.Controls.Add(this.groupBoxSituacaoFiscal);
            this.panelDadosEmpresa.Controls.Add(this.groupBoxApolices);
            this.panelDadosEmpresa.Controls.Add(this.groupBoxDeclaracoes);
            this.panelDadosEmpresa.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelDadosEmpresa.Location = new System.Drawing.Point(3, 3);
            this.panelDadosEmpresa.Name = "panelDadosEmpresa";
            this.panelDadosEmpresa.Size = new System.Drawing.Size(712, 450);
            this.panelDadosEmpresa.TabIndex = 92;
            // 
            // groupBoxInfoBasica
            // 
            this.groupBoxInfoBasica.Controls.Add(this.label2);
            this.groupBoxInfoBasica.Controls.Add(this.TXT_Sede);
            this.groupBoxInfoBasica.Controls.Add(this.label3);
            this.groupBoxInfoBasica.Controls.Add(this.TXT_Contribuinte);
            this.groupBoxInfoBasica.Controls.Add(this.label4);
            this.groupBoxInfoBasica.Controls.Add(this.TXT_Alvara);
            this.groupBoxInfoBasica.Controls.Add(this.label22);
            this.groupBoxInfoBasica.Controls.Add(this.TXT_AlvaraValidade);
            this.groupBoxInfoBasica.Controls.Add(this.AlertaValidadeAlvara);
            this.groupBoxInfoBasica.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxInfoBasica.Location = new System.Drawing.Point(8, 5);
            this.groupBoxInfoBasica.Name = "groupBoxInfoBasica";
            this.groupBoxInfoBasica.Size = new System.Drawing.Size(325, 132);
            this.groupBoxInfoBasica.TabIndex = 0;
            this.groupBoxInfoBasica.TabStop = false;
            this.groupBoxInfoBasica.Text = "Informações Básicas";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9F);
            this.label2.Location = new System.Drawing.Point(26, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 14);
            this.label2.TabIndex = 56;
            this.label2.Text = "Sede";
            // 
            // TXT_Sede
            // 
            this.TXT_Sede.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_Sede.Location = new System.Drawing.Point(66, 19);
            this.TXT_Sede.Name = "TXT_Sede";
            this.TXT_Sede.Size = new System.Drawing.Size(245, 22);
            this.TXT_Sede.TabIndex = 71;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 9F);
            this.label3.Location = new System.Drawing.Point(6, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 14);
            this.label3.TabIndex = 57;
            this.label3.Text = "NIF/NIPC";
            // 
            // TXT_Contribuinte
            // 
            this.TXT_Contribuinte.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_Contribuinte.Location = new System.Drawing.Point(66, 47);
            this.TXT_Contribuinte.Name = "TXT_Contribuinte";
            this.TXT_Contribuinte.Size = new System.Drawing.Size(245, 22);
            this.TXT_Contribuinte.TabIndex = 72;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri", 9F);
            this.label4.Location = new System.Drawing.Point(19, 78);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 14);
            this.label4.TabIndex = 58;
            this.label4.Text = "Alvará";
            // 
            // TXT_Alvara
            // 
            this.TXT_Alvara.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_Alvara.Location = new System.Drawing.Point(66, 75);
            this.TXT_Alvara.Name = "TXT_Alvara";
            this.TXT_Alvara.Size = new System.Drawing.Size(245, 22);
            this.TXT_Alvara.TabIndex = 73;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Calibri", 9F);
            this.label22.Location = new System.Drawing.Point(6, 108);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(56, 14);
            this.label22.TabIndex = 80;
            this.label22.Text = "Validade";
            // 
            // TXT_AlvaraValidade
            // 
            this.TXT_AlvaraValidade.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_AlvaraValidade.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.TXT_AlvaraValidade.Location = new System.Drawing.Point(67, 105);
            this.TXT_AlvaraValidade.Name = "TXT_AlvaraValidade";
            this.TXT_AlvaraValidade.ShowCheckBox = true;
            this.TXT_AlvaraValidade.Size = new System.Drawing.Size(228, 22);
            this.TXT_AlvaraValidade.TabIndex = 79;
            // 
            // AlertaValidadeAlvara
            // 
            this.AlertaValidadeAlvara.Location = new System.Drawing.Point(301, 105);
            this.AlertaValidadeAlvara.Name = "AlertaValidadeAlvara";
            this.AlertaValidadeAlvara.Size = new System.Drawing.Size(10, 10);
            this.AlertaValidadeAlvara.TabIndex = 88;
            // 
            // groupBoxSituacaoFiscal
            // 
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
            this.Controls.Add(this.panelModalDocumentos);
            this.groupBoxSituacaoFiscal.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxSituacaoFiscal.Location = new System.Drawing.Point(8, 143);
            this.groupBoxSituacaoFiscal.Name = "groupBoxSituacaoFiscal";
            this.groupBoxSituacaoFiscal.Size = new System.Drawing.Size(662, 170);
            this.groupBoxSituacaoFiscal.TabIndex = 1;
            this.groupBoxSituacaoFiscal.TabStop = false;
            this.groupBoxSituacaoFiscal.Text = "Situação Fiscal";
            // 
            // lblFolhaPagSS
            // 
            this.lblFolhaPagSS.AutoSize = true;
            this.lblFolhaPagSS.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblFolhaPagSS.Location = new System.Drawing.Point(269, 121);
            this.lblFolhaPagSS.Name = "lblFolhaPagSS";
            this.lblFolhaPagSS.Size = new System.Drawing.Size(78, 13);
            this.lblFolhaPagSS.TabIndex = 105;
            this.lblFolhaPagSS.Text = "Nenhum anexo";
            this.lblFolhaPagSS.Click += new System.EventHandler(this.visualizarFolhaPag_Click);
            // 
            // btnAnexoFolhaPag
            // 
            this.btnAnexoFolhaPag.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoFolhaPag.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoFolhaPag.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoFolhaPag.Location = new System.Drawing.Point(215, 117);
            this.btnAnexoFolhaPag.Name = "btnAnexoFolhaPag";
            this.btnAnexoFolhaPag.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoFolhaPag.TabIndex = 104;
            this.btnAnexoFolhaPag.Text = "...";
            this.btnAnexoFolhaPag.UseVisualStyleBackColor = false;
            this.btnAnexoFolhaPag.Click += new System.EventHandler(this.btnAnexoFolhaPag_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri", 9F);
            this.label5.Location = new System.Drawing.Point(9, 67);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(74, 14);
            this.label5.TabIndex = 59;
            this.label5.Text = "Não Div. Fin.";
            // 
            // TXT_NaoDivFinancas
            // 
            this.TXT_NaoDivFinancas.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_NaoDivFinancas.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.TXT_NaoDivFinancas.Location = new System.Drawing.Point(89, 61);
            this.TXT_NaoDivFinancas.Name = "TXT_NaoDivFinancas";
            this.TXT_NaoDivFinancas.Size = new System.Drawing.Size(120, 22);
            this.TXT_NaoDivFinancas.TabIndex = 74;
            // 
            // btnAnexoFinancas
            // 
            this.btnAnexoFinancas.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoFinancas.Location = new System.Drawing.Point(215, 61);
            this.btnAnexoFinancas.Name = "btnAnexoFinancas";
            this.btnAnexoFinancas.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoFinancas.TabIndex = 100;
            this.btnAnexoFinancas.Text = "...";
            this.btnAnexoFinancas.UseVisualStyleBackColor = true;
            this.btnAnexoFinancas.Click += new System.EventHandler(this.btnAnexoFinancas_Click);
            // 
            // lblAnexoFinancas
            // 
            this.lblAnexoFinancas.AutoSize = true;
            this.lblAnexoFinancas.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoFinancas.Location = new System.Drawing.Point(268, 66);
            this.lblAnexoFinancas.Name = "lblAnexoFinancas";
            this.lblAnexoFinancas.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoFinancas.TabIndex = 101;
            this.lblAnexoFinancas.Text = "Nenhum anexo";
            this.lblAnexoFinancas.Click += new System.EventHandler(this.visualizarAnexoFinancas_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Calibri", 9F);
            this.label6.Location = new System.Drawing.Point(11, 95);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 14);
            this.label6.TabIndex = 60;
            this.label6.Text = "Não Div. S.S.";
            // 
            // TXT_NaoDivSegSocial
            // 
            this.TXT_NaoDivSegSocial.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_NaoDivSegSocial.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.TXT_NaoDivSegSocial.Location = new System.Drawing.Point(89, 89);
            this.TXT_NaoDivSegSocial.Name = "TXT_NaoDivSegSocial";
            this.TXT_NaoDivSegSocial.Size = new System.Drawing.Size(120, 22);
            this.TXT_NaoDivSegSocial.TabIndex = 75;
            // 
            // btnAnexoSegSocial
            // 
            this.btnAnexoSegSocial.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoSegSocial.Location = new System.Drawing.Point(215, 91);
            this.btnAnexoSegSocial.Name = "btnAnexoSegSocial";
            this.btnAnexoSegSocial.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoSegSocial.TabIndex = 102;
            this.btnAnexoSegSocial.Text = "...";
            this.btnAnexoSegSocial.UseVisualStyleBackColor = true;
            this.btnAnexoSegSocial.Click += new System.EventHandler(this.btnAnexoSegSocial_Click);
            // 
            // lblAnexoSegSocial
            // 
            this.lblAnexoSegSocial.AutoSize = true;
            this.lblAnexoSegSocial.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoSegSocial.Location = new System.Drawing.Point(268, 95);
            this.lblAnexoSegSocial.Name = "lblAnexoSegSocial";
            this.lblAnexoSegSocial.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoSegSocial.TabIndex = 103;
            this.lblAnexoSegSocial.Text = "Nenhum anexo";
            this.lblAnexoSegSocial.Click += new System.EventHandler(this.visualizarAnexoSegSocial_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri", 9F);
            this.label7.Location = new System.Drawing.Point(9, 123);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(81, 14);
            this.label7.TabIndex = 61;
            this.label7.Text = "Folha Pag. S.S";
            // 
            // TXT_FolhaPagSegSocial
            // 
            this.TXT_FolhaPagSegSocial.Font = new System.Drawing.Font("Calibri", 9F);
            this.TXT_FolhaPagSegSocial.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.TXT_FolhaPagSegSocial.Location = new System.Drawing.Point(89, 117);
            this.TXT_FolhaPagSegSocial.Name = "TXT_FolhaPagSegSocial";
            this.TXT_FolhaPagSegSocial.Size = new System.Drawing.Size(120, 22);
            this.TXT_FolhaPagSegSocial.TabIndex = 76;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Calibri", 9F);
            this.label8.Location = new System.Drawing.Point(330, 25);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(47, 14);
            this.label8.TabIndex = 62;
            this.label8.Text = "Recibo:";
            // 
            // cb_ReciboPagSegSocial
            // 
            this.cb_ReciboPagSegSocial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_ReciboPagSegSocial.Font = new System.Drawing.Font("Calibri", 9F);
            this.cb_ReciboPagSegSocial.FormattingEnabled = true;
            this.cb_ReciboPagSegSocial.Location = new System.Drawing.Point(383, 22);
            this.cb_ReciboPagSegSocial.Name = "cb_ReciboPagSegSocial";
            this.cb_ReciboPagSegSocial.Size = new System.Drawing.Size(44, 22);
            this.cb_ReciboPagSegSocial.TabIndex = 81;
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
            // btnAnexarDocumentoGeral
            // 
            this.btnAnexarDocumentoGeral.BackColor = System.Drawing.Color.LightSteelBlue;
            this.btnAnexarDocumentoGeral.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexarDocumentoGeral.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.btnAnexarDocumentoGeral.Location = new System.Drawing.Point(450, 20);
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
            this.btnVerificarDocumentosFaltantes.Location = new System.Drawing.Point(450, 50);
            this.btnVerificarDocumentosFaltantes.Name = "btnVerificarDocumentosFaltantes";
            this.btnVerificarDocumentosFaltantes.Size = new System.Drawing.Size(130, 24);
            this.btnVerificarDocumentosFaltantes.TabIndex = 125;
            this.btnVerificarDocumentosFaltantes.Text = "Documentos Faltantes";
            this.btnVerificarDocumentosFaltantes.UseVisualStyleBackColor = false;
           // this.btnVerificarDocumentosFaltantes.Click += new System.EventHandler(this.btnVerificarDocumentosFaltantes_Click);
            // 
            // panelModalDocumentos
            // 
            this.panelModalDocumentos.BackColor = System.Drawing.Color.White;
            this.panelModalDocumentos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelModalDocumentos.Controls.Add(this.lblTipoDocumento);
            this.panelModalDocumentos.Controls.Add(this.cmbTipoDocumento);
            this.panelModalDocumentos.Controls.Add(this.btnConfirmarAnexo);
            this.panelModalDocumentos.Controls.Add(this.btnCancelarAnexo);
            this.panelModalDocumentos.Location = new System.Drawing.Point(250, 200);
            this.panelModalDocumentos.Name = "panelModalDocumentos";
            this.panelModalDocumentos.Size = new System.Drawing.Size(300, 150);
            this.panelModalDocumentos.TabIndex = 125;
            this.panelModalDocumentos.Visible = false;
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
            this.cmbTipoDocumento.Size = new System.Drawing.Size(240, 23);
            this.cmbTipoDocumento.TabIndex = 0;
            // 
            // btnConfirmarAnexo
            // 
            this.btnConfirmarAnexo.BackColor = System.Drawing.Color.LightGreen;
            this.btnConfirmarAnexo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnConfirmarAnexo.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnConfirmarAnexo.Location = new System.Drawing.Point(60, 100);
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
            this.btnCancelarAnexo.Location = new System.Drawing.Point(160, 100);
            this.btnCancelarAnexo.Name = "btnCancelarAnexo";
            this.btnCancelarAnexo.Size = new System.Drawing.Size(80, 30);
            this.btnCancelarAnexo.TabIndex = 2;
            this.btnCancelarAnexo.Text = "Cancelar";
            this.btnCancelarAnexo.UseVisualStyleBackColor = false;
            this.btnCancelarAnexo.Click += new System.EventHandler(this.btnCancelarAnexo_Click);
            // 
            // lblTipoDocumento
            // 
            this.lblTipoDocumento.AutoSize = true;
            this.lblTipoDocumento.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold);
            this.lblTipoDocumento.Location = new System.Drawing.Point(78, 20);
            this.lblTipoDocumento.Name = "lblTipoDocumento";
            this.lblTipoDocumento.Size = new System.Drawing.Size(144, 18);
            this.lblTipoDocumento.TabIndex = 3;
            this.lblTipoDocumento.Text = "Selecione o documento";
            // 
            // lblAnexoHorarioTrabalho
            // 
            this.lblAnexoHorarioTrabalho.AutoSize = true;
            this.lblAnexoHorarioTrabalho.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoHorarioTrabalho.Location = new System.Drawing.Point(560, 22);
            this.lblAnexoHorarioTrabalho.Name = "lblAnexoHorarioTrabalho";
            this.lblAnexoHorarioTrabalho.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoHorarioTrabalho.TabIndex = 110;
            this.lblAnexoHorarioTrabalho.Text = "Nenhum anexo";
            this.lblAnexoHorarioTrabalho.Click += new System.EventHandler(this.visualizarHorarioTrabalho_Click);
            // 
            // btnAnexoHorarioTrabalho
            // 
            this.btnAnexoHorarioTrabalho.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoHorarioTrabalho.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoHorarioTrabalho.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoHorarioTrabalho.Location = new System.Drawing.Point(510, 18);
            this.btnAnexoHorarioTrabalho.Name = "btnAnexoHorarioTrabalho";
            this.btnAnexoHorarioTrabalho.Size = new System.Drawing.Size(47, 22);
            this.btnAnexoHorarioTrabalho.TabIndex = 111;
            this.btnAnexoHorarioTrabalho.Text = "...";
            this.btnAnexoHorarioTrabalho.UseVisualStyleBackColor = false;
            this.btnAnexoHorarioTrabalho.Click += new System.EventHandler(this.btnAnexoHorarioTrabalho_Click);
            // 
            // lblAnexoApoliceAT
            // 
            this.lblAnexoApoliceAT.AutoSize = true;
            this.lblAnexoApoliceAT.Font = new System.Drawing.Font("Calibri", 8F);
            this.lblAnexoApoliceAT.Location = new System.Drawing.Point(560, 48);
            this.lblAnexoApoliceAT.Name = "lblAnexoApoliceAT";
            this.lblAnexoApoliceAT.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoApoliceAT.TabIndex = 106;
            this.lblAnexoApoliceAT.Text = "Nenhum anexo";
            this.lblAnexoApoliceAT.Click += new System.EventHandler(this.visualizarApoliceAT_Click);
            // 
            // btnAnexoApoliceAT
            // 
            this.btnAnexoApoliceAT.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoApoliceAT.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoApoliceAT.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoApoliceAT.Location = new System.Drawing.Point(510, 44);
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
            this.lblAnexoApoliceRC.Location = new System.Drawing.Point(560, 74);
            this.lblAnexoApoliceRC.Name = "lblAnexoApoliceRC";
            this.lblAnexoApoliceRC.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoApoliceRC.TabIndex = 108;
            this.lblAnexoApoliceRC.Text = "Nenhum anexo";
            this.lblAnexoApoliceRC.Click += new System.EventHandler(this.visualizarApoliceRC_Click);
            // 
            // btnAnexoApoliceRC
            // 
            this.btnAnexoApoliceRC.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoApoliceRC.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoApoliceRC.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoApoliceRC.Location = new System.Drawing.Point(510, 70);
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
            this.lblAnexoD.Location = new System.Drawing.Point(560, 100);
            this.lblAnexoD.Name = "lblAnexoD";
            this.lblAnexoD.Size = new System.Drawing.Size(78, 13);
            this.lblAnexoD.TabIndex = 112;
            this.lblAnexoD.Text = "Nenhum anexo";
            this.lblAnexoD.Click += new System.EventHandler(this.visualizarAnexoD_Click);
            // 
            // btnAnexoD
            // 
            this.btnAnexoD.BackColor = System.Drawing.Color.LightBlue;
            this.btnAnexoD.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAnexoD.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnAnexoD.Location = new System.Drawing.Point(510, 96);
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
            this.lblDecTrabEmigr.Location = new System.Drawing.Point(560, 126);
            this.lblDecTrabEmigr.Name = "lblDecTrabEmigr";
            this.lblDecTrabEmigr.Size = new System.Drawing.Size(78, 13);
            this.lblDecTrabEmigr.TabIndex = 114;
            this.lblDecTrabEmigr.Text = "Nenhum anexo";
            this.lblDecTrabEmigr.Click += new System.EventHandler(this.visualizarDecTrabEmigr_Click);
            // 
            // btnDecTrabEmigr
            // 
            this.btnDecTrabEmigr.BackColor = System.Drawing.Color.LightBlue;
            this.btnDecTrabEmigr.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDecTrabEmigr.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnDecTrabEmigr.Location = new System.Drawing.Point(510, 122);
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
            this.lblInscricaoSS.Location = new System.Drawing.Point(560, 152);
            this.lblInscricaoSS.Name = "lblInscricaoSS";
            this.lblInscricaoSS.Size = new System.Drawing.Size(78, 13);
            this.lblInscricaoSS.TabIndex = 116;
            this.lblInscricaoSS.Text = "Nenhum anexo";
            this.lblInscricaoSS.Click += new System.EventHandler(this.visualizarInscricaoSS_Click);
            // 
            // btnInscricaoSS
            // 
            this.btnInscricaoSS.BackColor = System.Drawing.Color.LightBlue;
            this.btnInscricaoSS.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInscricaoSS.Font = new System.Drawing.Font("Calibri", 9F);
            this.btnInscricaoSS.Location = new System.Drawing.Point(510, 148);
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
            this.lblHorarioTrabalhoTitle.Location = new System.Drawing.Point(350, 22);
            this.lblHorarioTrabalhoTitle.Name = "lblHorarioTrabalhoTitle";
            this.lblHorarioTrabalhoTitle.Size = new System.Drawing.Size(95, 14);
            this.lblHorarioTrabalhoTitle.TabIndex = 118;
            this.lblHorarioTrabalhoTitle.Text = "Horário Trabalho:";
            // 
            // lblApoliceATTitle
            // 
            this.lblApoliceATTitle.AutoSize = true;
            this.lblApoliceATTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblApoliceATTitle.Location = new System.Drawing.Point(350, 48);
            this.lblApoliceATTitle.Name = "lblApoliceATTitle";
            this.lblApoliceATTitle.Size = new System.Drawing.Size(62, 14);
            this.lblApoliceATTitle.TabIndex = 119;
            this.lblApoliceATTitle.Text = "Apólice AT:";
            // 
            // lblApoliceRCTitle
            // 
            this.lblApoliceRCTitle.AutoSize = true;
            this.lblApoliceRCTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblApoliceRCTitle.Location = new System.Drawing.Point(350, 74);
            this.lblApoliceRCTitle.Name = "lblApoliceRCTitle";
            this.lblApoliceRCTitle.Size = new System.Drawing.Size(63, 14);
            this.lblApoliceRCTitle.TabIndex = 120;
            this.lblApoliceRCTitle.Text = "Apólice RC:";
            // 
            // lblAnexoDTitle
            // 
            this.lblAnexoDTitle.AutoSize = true;
            this.lblAnexoDTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblAnexoDTitle.Location = new System.Drawing.Point(350, 100);
            this.lblAnexoDTitle.Name = "lblAnexoDTitle";
            this.lblAnexoDTitle.Size = new System.Drawing.Size(52, 14);
            this.lblAnexoDTitle.TabIndex = 121;
            this.lblAnexoDTitle.Text = "Anexo D:";
            // 
            // lblDecTrabEmigrTitle
            // 
            this.lblDecTrabEmigrTitle.AutoSize = true;
            this.lblDecTrabEmigrTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblDecTrabEmigrTitle.Location = new System.Drawing.Point(350, 126);
            this.lblDecTrabEmigrTitle.Name = "lblDecTrabEmigrTitle";
            this.lblDecTrabEmigrTitle.Size = new System.Drawing.Size(108, 14);
            this.lblDecTrabEmigrTitle.TabIndex = 122;
            this.lblDecTrabEmigrTitle.Text = "Dec. Trab. Emigr.:";
            // 
            // lblInscricaoSSTitle
            // 
            this.lblInscricaoSSTitle.AutoSize = true;
            this.lblInscricaoSSTitle.Font = new System.Drawing.Font("Calibri", 9F);
            this.lblInscricaoSSTitle.Location = new System.Drawing.Point(350, 152);
            this.lblInscricaoSSTitle.Name = "lblInscricaoSSTitle";
            this.lblInscricaoSSTitle.Size = new System.Drawing.Size(75, 14);
            this.lblInscricaoSSTitle.TabIndex = 123;
            this.lblInscricaoSSTitle.Text = "Inscrição SS:";
            // 
            // groupBoxApolices
            // 
            this.groupBoxApolices.Controls.Add(this.label9);
            this.groupBoxApolices.Controls.Add(this.cb_ApoliceAT);
            this.groupBoxApolices.Controls.Add(this.label10);
            this.groupBoxApolices.Controls.Add(this.TXT_ReciboApoliceAT);
            this.groupBoxApolices.Controls.Add(this.label11);
            this.groupBoxApolices.Controls.Add(this.cb_ApoliceRC);
            this.groupBoxApolices.Controls.Add(this.label12);
            this.groupBoxApolices.Controls.Add(this.TXT_ReciboRC);
            this.groupBoxApolices.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxApolices.Location = new System.Drawing.Point(345, 5);
            this.groupBoxApolices.Name = "groupBoxApolices";
            this.groupBoxApolices.Size = new System.Drawing.Size(325, 132);
            this.groupBoxApolices.TabIndex = 2;
            this.groupBoxApolices.TabStop = false;
            this.groupBoxApolices.Text = "Apólices de Seguro";
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
            this.groupBoxDeclaracoes.Location = new System.Drawing.Point(8, 312);
            this.groupBoxDeclaracoes.Name = "groupBoxDeclaracoes";
            this.groupBoxDeclaracoes.Size = new System.Drawing.Size(662, 122);
            this.groupBoxDeclaracoes.TabIndex = 3;
            this.groupBoxDeclaracoes.TabStop = false;
            this.groupBoxDeclaracoes.Text = "Declarações";
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
            this.panelObras.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelObras.Location = new System.Drawing.Point(3, 466);
            this.panelObras.Name = "panelObras";
            this.panelObras.Size = new System.Drawing.Size(712, 223);
            this.panelObras.TabIndex = 93;
            // 
            // groupBoxObras
            // 
            this.groupBoxObras.Controls.Add(this.label17);
            this.groupBoxObras.Controls.Add(this.cb_Obras);
            this.groupBoxObras.Controls.Add(this.btnGravarObra);
            this.groupBoxObras.Controls.Add(this.dataGridView1);
            this.groupBoxObras.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBoxObras.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold);
            this.groupBoxObras.Location = new System.Drawing.Point(0, 0);
            this.groupBoxObras.Name = "groupBoxObras";
            this.groupBoxObras.Size = new System.Drawing.Size(712, 223);
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
            this.btnGravarObra.Click += new System.EventHandler(this.button1_Click);
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
            this.AutorizacaoEntrada});
            this.dataGridView1.Location = new System.Drawing.Point(11, 49);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 25;
            this.dataGridView1.Size = new System.Drawing.Size(658, 168);
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
            // AutorizacaoEntrada
            // 
            this.AutorizacaoEntrada.HeaderText = "Autorização de Entrada";
            this.AutorizacaoEntrada.Name = "AutorizacaoEntrada";
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(718, 692);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Trabalhadores";
            this.tabPage2.UseVisualStyleBackColor = true;
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
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(757, 797);
            this.Controls.Add(this.lblSelecionar);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.TXT_Codigo);
            this.Controls.Add(this.BTF4);
            this.Controls.Add(this.TXT_Nome);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
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
            this.groupBoxApolices.ResumeLayout(false);
            this.groupBoxApolices.PerformLayout();
            this.groupBoxDeclaracoes.ResumeLayout(false);
            this.groupBoxDeclaracoes.PerformLayout();
            this.panelObras.ResumeLayout(false);
            this.groupBoxObras.ResumeLayout(false);
            this.groupBoxObras.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}
