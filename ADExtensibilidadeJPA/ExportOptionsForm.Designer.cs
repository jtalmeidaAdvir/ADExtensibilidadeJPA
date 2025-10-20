
namespace ADExtensibilidadeJPA
{
    partial class ExportOptionsForm
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
            this.checkBoxEmpresa = new System.Windows.Forms.CheckBox();
            this.checkBoxTrabalhadores = new System.Windows.Forms.CheckBox();
            this.checkBoxEquipamentos = new System.Windows.Forms.CheckBox();
            this.btnConfirmar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();

            // label1
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.FromArgb(59, 89, 152);
            this.label1.Location = new System.Drawing.Point(20, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(200, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "Selecione o que deseja exportar:";

            // checkBoxEmpresa
            this.checkBoxEmpresa.AutoSize = true;
            this.checkBoxEmpresa.Font = new System.Drawing.Font("Calibri", 10F);
            this.checkBoxEmpresa.Location = new System.Drawing.Point(40, 55);
            this.checkBoxEmpresa.Name = "checkBoxEmpresa";
            this.checkBoxEmpresa.Size = new System.Drawing.Size(150, 21);
            this.checkBoxEmpresa.TabIndex = 1;
            this.checkBoxEmpresa.Text = "Dados da Empresa";
            this.checkBoxEmpresa.UseVisualStyleBackColor = true;

            // checkBoxTrabalhadores
            this.checkBoxTrabalhadores.AutoSize = true;
            this.checkBoxTrabalhadores.Font = new System.Drawing.Font("Calibri", 10F);
            this.checkBoxTrabalhadores.Location = new System.Drawing.Point(40, 85);
            this.checkBoxTrabalhadores.Name = "checkBoxTrabalhadores";
            this.checkBoxTrabalhadores.Size = new System.Drawing.Size(110, 21);
            this.checkBoxTrabalhadores.TabIndex = 2;
            this.checkBoxTrabalhadores.Text = "Trabalhadores";
            this.checkBoxTrabalhadores.UseVisualStyleBackColor = true;

            // checkBoxEquipamentos
            this.checkBoxEquipamentos.AutoSize = true;
            this.checkBoxEquipamentos.Font = new System.Drawing.Font("Calibri", 10F);
            this.checkBoxEquipamentos.Location = new System.Drawing.Point(40, 115);
            this.checkBoxEquipamentos.Name = "checkBoxEquipamentos";
            this.checkBoxEquipamentos.Size = new System.Drawing.Size(108, 21);
            this.checkBoxEquipamentos.TabIndex = 3;
            this.checkBoxEquipamentos.Text = "Equipamentos";
            this.checkBoxEquipamentos.UseVisualStyleBackColor = true;

            // btnConfirmar
            this.btnConfirmar.BackColor = System.Drawing.Color.FromArgb(59, 89, 152);
            this.btnConfirmar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnConfirmar.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Bold);
            this.btnConfirmar.ForeColor = System.Drawing.Color.White;
            this.btnConfirmar.Location = new System.Drawing.Point(40, 160);
            this.btnConfirmar.Name = "btnConfirmar";
            this.btnConfirmar.Size = new System.Drawing.Size(100, 35);
            this.btnConfirmar.TabIndex = 4;
            this.btnConfirmar.Text = "Confirmar";
            this.btnConfirmar.UseVisualStyleBackColor = false;
            this.btnConfirmar.Click += new System.EventHandler(this.btnConfirmar_Click);

            // btnCancelar
            this.btnCancelar.BackColor = System.Drawing.Color.White;
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.Font = new System.Drawing.Font("Calibri", 10F);
            this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(59, 89, 152);
            this.btnCancelar.Location = new System.Drawing.Point(150, 160);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(100, 35);
            this.btnCancelar.TabIndex = 5;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = false;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);

            // ExportOptionsForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(300, 220);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnConfirmar);
            this.Controls.Add(this.checkBoxEquipamentos);
            this.Controls.Add(this.checkBoxTrabalhadores);
            this.Controls.Add(this.checkBoxEmpresa);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExportOptionsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Opções de Exportação";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
