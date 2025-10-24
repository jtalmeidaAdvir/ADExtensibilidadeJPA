namespace ADExtensibilidadeJPA
{
    partial class QuadroControlo
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button BT_DadosJPA;

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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.BT_Editar = new System.Windows.Forms.Button();
            this.Bt_Email = new System.Windows.Forms.Button();
            this.Bt_Validades = new System.Windows.Forms.Button();
            this.Bt_Avisos = new System.Windows.Forms.Button();
            this.Bt_imprimir = new System.Windows.Forms.Button();
            this.f4TabelaSQL1 = new PRISDK100.F4TabelaSQL();
            this.label1 = new System.Windows.Forms.Label();
            this.BT_ImprimirJPA = new System.Windows.Forms.Button();
            this.BT_CriarTrabalhadores = new System.Windows.Forms.Button();
            this.Bt_imprimir2 = new System.Windows.Forms.Button();
            this.bt_dadosJPAa = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(3, 134);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(800, 354);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.UseWaitCursor = true;
            this.dataGridView1.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentDoubleClick);
            // 
            // BT_Editar
            // 
            this.BT_Editar.Location = new System.Drawing.Point(3, 4);
            this.BT_Editar.Name = "BT_Editar";
            this.BT_Editar.Size = new System.Drawing.Size(75, 23);
            this.BT_Editar.TabIndex = 1;
            this.BT_Editar.Text = "Editar";
            this.BT_Editar.UseVisualStyleBackColor = true;
            this.BT_Editar.Visible = false;
            this.BT_Editar.Click += new System.EventHandler(this.BT_Editar_Click);
            // 
            // Bt_Email
            // 
            this.Bt_Email.Location = new System.Drawing.Point(84, 4);
            this.Bt_Email.Name = "Bt_Email";
            this.Bt_Email.Size = new System.Drawing.Size(75, 23);
            this.Bt_Email.TabIndex = 2;
            this.Bt_Email.Text = "Enviar Email";
            this.Bt_Email.UseVisualStyleBackColor = true;
            this.Bt_Email.Click += new System.EventHandler(this.Bt_Email_Click);
            // 
            // Bt_Validades
            // 
            this.Bt_Validades.Location = new System.Drawing.Point(508, 4);
            this.Bt_Validades.Name = "Bt_Validades";
            this.Bt_Validades.Size = new System.Drawing.Size(75, 23);
            this.Bt_Validades.TabIndex = 3;
            this.Bt_Validades.Text = "Validades";
            this.Bt_Validades.UseVisualStyleBackColor = true;
            this.Bt_Validades.Click += new System.EventHandler(this.Bt_Validades_Click);
            // 
            // Bt_Avisos
            // 
            this.Bt_Avisos.Location = new System.Drawing.Point(589, 4);
            this.Bt_Avisos.Name = "Bt_Avisos";
            this.Bt_Avisos.Size = new System.Drawing.Size(104, 23);
            this.Bt_Avisos.TabIndex = 4;
            this.Bt_Avisos.Text = "Enviar Avisos";
            this.Bt_Avisos.UseVisualStyleBackColor = true;
            this.Bt_Avisos.Click += new System.EventHandler(this.Bt_Avisos_Click);
            // 
            // Bt_imprimir
            // 
            this.Bt_imprimir.Location = new System.Drawing.Point(699, 4);
            this.Bt_imprimir.Name = "Bt_imprimir";
            this.Bt_imprimir.Size = new System.Drawing.Size(75, 23);
            this.Bt_imprimir.TabIndex = 5;
            this.Bt_imprimir.Text = "Imprimir";
            this.Bt_imprimir.UseVisualStyleBackColor = true;
            this.Bt_imprimir.Visible = false;
            this.Bt_imprimir.Click += new System.EventHandler(this.Bt_imprimir_Click);
            // 
            // f4TabelaSQL1
            // 
            this.f4TabelaSQL1.AliasCampoChave = "";
            this.f4TabelaSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.f4TabelaSQL1.CampoChave = "";
            this.f4TabelaSQL1.CampoDescricao = "";
            this.f4TabelaSQL1.Caption = "Obra:";
            this.f4TabelaSQL1.F4Modal = false;
            this.f4TabelaSQL1.Inicializado = false;
            this.f4TabelaSQL1.Location = new System.Drawing.Point(99, 105);
            this.f4TabelaSQL1.MaxLengthF4 = 50;
            this.f4TabelaSQL1.Modulo = "";
            this.f4TabelaSQL1.MostraCaption = true;
            this.f4TabelaSQL1.Name = "f4TabelaSQL1";
            this.f4TabelaSQL1.ResourceID = 0;
            this.f4TabelaSQL1.ResourceIDTituloLista = 0;
            this.f4TabelaSQL1.SelectionFormula = "";
            this.f4TabelaSQL1.Size = new System.Drawing.Size(691, 21);
            this.f4TabelaSQL1.TabIndex = 9;
            this.f4TabelaSQL1.TituloLista = "Obras";
            this.f4TabelaSQL1.WidthCaption = 1000;
            this.f4TabelaSQL1.WidthEspacamento = 60;
            this.f4TabelaSQL1.WidthF4 = 1590;
            this.f4TabelaSQL1.Load += new System.EventHandler(this.f4TabelaSQL1_Load);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(14, 109);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Filtrar por Obra:";
            // 
            // BT_ImprimirJPA
            // 
            this.BT_ImprimirJPA.Location = new System.Drawing.Point(650, 4);
            this.BT_ImprimirJPA.Name = "BT_ImprimirJPA";
            this.BT_ImprimirJPA.Size = new System.Drawing.Size(75, 23);
            this.BT_ImprimirJPA.TabIndex = 11;
            this.BT_ImprimirJPA.Text = "Imprimir JPA";
            this.BT_ImprimirJPA.UseVisualStyleBackColor = true;
            this.BT_ImprimirJPA.Visible = false;
            this.BT_ImprimirJPA.Click += new System.EventHandler(this.BT_ImprimirJPA_Click);
            // 
            // BT_CriarTrabalhadores
            // 
            this.BT_CriarTrabalhadores.Location = new System.Drawing.Point(731, 4);
            this.BT_CriarTrabalhadores.Name = "BT_CriarTrabalhadores";
            this.BT_CriarTrabalhadores.Size = new System.Drawing.Size(120, 23);
            this.BT_CriarTrabalhadores.TabIndex = 12;
            this.BT_CriarTrabalhadores.Text = "Criar Trabalhadores";
            this.BT_CriarTrabalhadores.UseVisualStyleBackColor = true;
            this.BT_CriarTrabalhadores.Click += new System.EventHandler(this.BT_CriarTrabalhadores_Click);
            // 
            // Bt_imprimir2
            // 
            this.Bt_imprimir2.Location = new System.Drawing.Point(427, 4);
            this.Bt_imprimir2.Name = "Bt_imprimir2";
            this.Bt_imprimir2.Size = new System.Drawing.Size(75, 23);
            this.Bt_imprimir2.TabIndex = 13;
            this.Bt_imprimir2.Text = "Imprimir";
            this.Bt_imprimir2.UseVisualStyleBackColor = true;
            this.Bt_imprimir2.Click += new System.EventHandler(this.Bt_imprimir2_Click);
            // 
            // bt_dadosJPAa
            // 
            this.bt_dadosJPAa.Location = new System.Drawing.Point(346, 4);
            this.bt_dadosJPAa.Name = "bt_dadosJPAa";
            this.bt_dadosJPAa.Size = new System.Drawing.Size(75, 23);
            this.bt_dadosJPAa.TabIndex = 14;
            this.bt_dadosJPAa.Text = "dados JPA";
            this.bt_dadosJPAa.UseVisualStyleBackColor = true;
            this.bt_dadosJPAa.Click += new System.EventHandler(this.bt_dadosJPAa_Click);
            // 
            // QuadroControlo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.bt_dadosJPAa);
            this.Controls.Add(this.Bt_imprimir2);
            this.Controls.Add(this.BT_CriarTrabalhadores);
            this.Controls.Add(this.BT_ImprimirJPA);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.f4TabelaSQL1);
            this.Controls.Add(this.Bt_imprimir);
            this.Controls.Add(this.Bt_Validades);
            this.Controls.Add(this.Bt_Email);
            this.Controls.Add(this.BT_Editar);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.Bt_Avisos);
            this.Name = "QuadroControlo";
            this.Size = new System.Drawing.Size(806, 491);
            this.Text = "Quadro de Controlo SGS de Subempreiteiros";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.QuadroControlo_FormClosed);
            this.Load += new System.EventHandler(this.QuadroControlo_Load_1);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button BT_Editar;
        private System.Windows.Forms.Button Bt_Email;
        private System.Windows.Forms.Button Bt_Validades;
        private System.Windows.Forms.Button Bt_Avisos;
        private System.Windows.Forms.Button Bt_imprimir;
        private PRISDK100.F4TabelaSQL f4TabelaSQL1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BT_ImprimirJPA;
        private System.Windows.Forms.Button BT_CriarTrabalhadores;
        private System.Windows.Forms.Button Bt_imprimir2;
        private System.Windows.Forms.Button bt_dadosJPAa;
    }
}