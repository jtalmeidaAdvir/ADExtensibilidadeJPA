
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    partial class Validades
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
        private System.Windows.Forms.TextBox txtCaminhoPasta;
        private System.Windows.Forms.TextBox txt_caminhotrab;
        private System.Windows.Forms.TextBox txt_caminhoequi;
        private System.Windows.Forms.TextBox txtcaminhoAuto;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Validades));

            this.txtCaminhoPasta = new System.Windows.Forms.TextBox();
            this.txt_caminhotrab = new System.Windows.Forms.TextBox();
            this.txt_caminhoequi = new System.Windows.Forms.TextBox();
            this.txtcaminhoAuto = new System.Windows.Forms.TextBox();

            this.SuspendLayout();
            // 
            // Validades
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 600);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Validades";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Validades dos Documentos";
            this.ResumeLayout(false);

        }

        #endregion
    }
}
