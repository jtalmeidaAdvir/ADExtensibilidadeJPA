using ErpBS100;
using Primavera.Extensibility.BusinessEntities;
using Primavera.Extensibility.CustomForm;
using StdBE100;
using StdPlatBS100;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace ADExtensibilidadeJPA
{
    public partial class Menu : Form
    {
        #region Properties and Fields
        public string _ID;
        public string IdSelecionado;
        public ErpBS BSO { get; private set; }
        public StdBSInterfPub PSO { get; private set; }

        // Managers
        private EmpresaManager _empresaManager;
        private TrabalhadorManager _trabalhadorManager;

        #endregion

        #region Initialization
        public Menu(ErpBS100.ErpBS bSO, StdPlatBS100.StdBSInterfPub pSO, string idSelecionado)
        {
            InitializeComponent();
            ConfigurarEstiloControles();

            BSO = bSO;
            PSO = pSO;
            IdSelecionado = idSelecionado;

            // Inicializar os managers
            _empresaManager = new EmpresaManager(BSO, PSO, IdSelecionado, this);
            _trabalhadorManager = new TrabalhadorManager(tabPage2);

            if (IdSelecionado != "")
            {
                _empresaManager.CarregarDados();
            }
        }

        private void ConfigurarEstiloControles()
        {
            // Configuração das cores dos controles para uma aparência mais moderna
            this.BackColor = System.Drawing.Color.WhiteSmoke;

            // Configurar estilo do DataGridView
            dataGridView1.BorderStyle = BorderStyle.None;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke;

            // Destacar botões
            BTF4.FlatStyle = FlatStyle.Flat;
            btnGravarObra.FlatStyle = FlatStyle.Flat;

            // Configurar os botões para anexar documentos específicos
            btnAnexoFinancas.FlatStyle = FlatStyle.Flat;
            btnAnexoFinancas.BackColor = Color.LightBlue;
            btnAnexoSegSocial.FlatStyle = FlatStyle.Flat;
            btnAnexoSegSocial.BackColor = Color.LightBlue;

            // Configurar painéis
            AlertaValidadeAlvara.BackColor = Color.Red;

            // Configurar valores iniciais para os DateTimePickers
            // Se a data atual não for apropriada como valor padrão, pode definir um valor mínimo
            TXT_NaoDivFinancas.Value = DateTime.Today;
            TXT_NaoDivSegSocial.Value = DateTime.Today;
            TXT_FolhaPagSegSocial.Value = DateTime.Today;
            TXT_AlvaraValidade.Value = DateTime.Today;

            // Permitir limpar as datas (opcional)
            TXT_NaoDivFinancas.ShowCheckBox = true;
            TXT_NaoDivSegSocial.ShowCheckBox = true;
            TXT_FolhaPagSegSocial.ShowCheckBox = true;
            TXT_AlvaraValidade.ShowCheckBox = true;

            // Definir o formato de exibição para mostrar a data completa
            TXT_NaoDivFinancas.Format = DateTimePickerFormat.Short;
            TXT_NaoDivSegSocial.Format = DateTimePickerFormat.Short;
            TXT_FolhaPagSegSocial.Format = DateTimePickerFormat.Short;
        }
        #endregion

        #region Eventos da Interface
        // Manipuladores de eventos delegados para os managers
        private void BTF4_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> entidade = new Dictionary<string, string>();
            _empresaManager.GetEntidades(ref entidade);

            if (entidade.Count > 0)
            {
                _empresaManager.SetInfoEntidades(entidade);
            }
        }

        private void BT_Salvar_Click_Click(object sender, EventArgs e)
        {
            _empresaManager.SalvarDados();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _empresaManager.SalvarObra();
        }

        private void cb_Obras_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Obras.SelectedItem is KeyValuePair<string, string> obraSelecionada)
            {
                string codigoObraSelecionada = obraSelecionada.Key;
                _empresaManager.CarregarObrasEmDataGridView(codigoObraSelecionada);
            }
        }

        private void btnSelecionarPasta_Click(object sender, EventArgs e)
        {
            _empresaManager.SelecionarPasta();
        }

        private void btnAnexarDocumento_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarDocumento();
        }

        private void btnAnexoFinancas_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarDocumentoFinancas();
        }

        private void btnAnexoSegSocial_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarDocumentoSegSocial();
        }

        private void btnAnexoFolhaPag_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarFolhaPag();
        }

        private void visualizarAnexoFinancas_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarAnexoFinancas();
        }

        private void visualizarAnexoSegSocial_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarAnexoSegSocial();
        }

        private void visualizarFolhaPag_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarFolhaPag();
        }

        private void dgvTrabalhadores_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Use directly the DataGridViewCellEventArgs
            if (e.RowIndex >= 0)
            {
                DataGridView dgv = sender as DataGridView;
                if (dgv.Columns[e.ColumnIndex].Name == "Editar" || dgv.Columns[e.ColumnIndex].Name == "Remover")
                {
                    // Call the handler in TrabalhadorManager
                    _trabalhadorManager.HandleCellClick(dgv, e);
                }
            }
        }
        #endregion
    }
}