namespace ADExtensibilidadeJPA
{
    partial class QuadroControlo
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.dataGridView1.Location = new System.Drawing.Point(3, 33);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(794, 414);
            this.dataGridView1.TabIndex = 0;
            // 
            // BT_Editar
            // 
            this.BT_Editar.Location = new System.Drawing.Point(3, 4);
            this.BT_Editar.Name = "BT_Editar";
            this.BT_Editar.Size = new System.Drawing.Size(75, 23);
            this.BT_Editar.TabIndex = 1;
            this.BT_Editar.Text = "Editar";
            this.BT_Editar.UseVisualStyleBackColor = true;
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
            // QuadroControlo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.Bt_Validades);
            this.Controls.Add(this.Bt_Email);
            this.Controls.Add(this.BT_Editar);
            this.Controls.Add(this.dataGridView1);
            this.Name = "QuadroControlo";
            this.Size = new System.Drawing.Size(800, 450);
            this.Text = "Quadro de Controlo";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button BT_Editar;
        private System.Windows.Forms.Button Bt_Email;
        private System.Windows.Forms.Button Bt_Validades;
    }
}