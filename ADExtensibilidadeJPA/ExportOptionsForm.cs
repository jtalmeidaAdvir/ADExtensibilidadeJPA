
using System;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class ExportOptionsForm : Form
    {
        public bool ExportarEmpresa { get; private set; }
        public bool ExportarTrabalhadores { get; private set; }
        public bool ExportarEquipamentos { get; private set; }

        public ExportOptionsForm()
        {
            InitializeComponent();

            // Por defeito, tudo fica selecionado
            checkBoxEmpresa.Checked = true;
            checkBoxTrabalhadores.Checked = true;
            checkBoxEquipamentos.Checked = true;
        }

        private CheckBox checkBoxEmpresa;
        private CheckBox checkBoxTrabalhadores;
        private CheckBox checkBoxEquipamentos;
        private Button btnConfirmar;
        private Button btnCancelar;
        private Label label1;

        private void btnConfirmar_Click(object sender, EventArgs e)
        {
            if (!checkBoxEmpresa.Checked && !checkBoxTrabalhadores.Checked && !checkBoxEquipamentos.Checked)
            {
                MessageBox.Show("Por favor, selecione pelo menos uma opção para exportar.",
                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ExportarEmpresa = checkBoxEmpresa.Checked;
            ExportarTrabalhadores = checkBoxTrabalhadores.Checked;
            ExportarEquipamentos = checkBoxEquipamentos.Checked;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
