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

        private void btnAnexoApoliceAT_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarDocumentoApoliceAT();
        }

        private void btnAnexoApoliceRC_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarDocumentoApoliceRC();
        }

        private void btnAnexoHorarioTrabalho_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarHorarioTrabalho();
        }

        private void btnAnexoD_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarAnexoD();
        }

        private void btnDecTrabEmigr_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarDecTrabEmigr();
        }

        private void btnInscricaoSS_Click(object sender, EventArgs e)
        {
            _empresaManager.AnexarInscricaoSS();
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

        private void visualizarApoliceAT_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarApoliceAT();
        }

        private void visualizarApoliceRC_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarApoliceRC();
        }

        private void visualizarHorarioTrabalho_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarHorarioTrabalho();
        }

        private void visualizarAnexoD_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarAnexoD();
        }

        private void visualizarDecTrabEmigr_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarDecTrabEmigr();
        }

        private void visualizarInscricaoSS_Click(object sender, EventArgs e)
        {
            _empresaManager.VisualizarInscricaoSS();
        }

        private void btnAnexarDocumentoGeral_Click(object sender, EventArgs e)
        {
            // Atualizar os itens do combobox para mostrar quais documentos já estão anexados
            UpdateDocumentComboBox();
            panelModalDocumentos.Visible = true;
            panelModalDocumentos.BringToFront();/*
            // Posicionar o modal no centro do formulário
            panelModalDocumentos.Location = new Point(
                (this.ClientSize.Width - panelModalDocumentos.Width) / 2,
                (this.ClientSize.Height - panelModalDocumentos.Height) / 2);

            // Exibir o modal
        

            // Selecionar a primeira opção por padrão
            if (cmbTipoDocumento.Items.Count > 0)
                cmbTipoDocumento.SelectedIndex = 0;*/
        }

        private void UpdateDocumentComboBox()
        {
            // Limpar os itens existentes
            cmbTipoDocumento.Items.Clear();

            // Verificar quais documentos já estão anexados
            bool[] documentosAnexados = _empresaManager.GetDocumentosAnexados();

            // Adicionar os itens com prefixo indicando status
            cmbTipoDocumento.Items.Add(documentosAnexados[0] ? "✓ Não Div. Financas" : "□ Não Div. Financas");
            cmbTipoDocumento.Items.Add(documentosAnexados[1] ? "✓ Não Div. Seg. Social" : "□ Não Div. Seg. Social");
            cmbTipoDocumento.Items.Add(documentosAnexados[2] ? "✓ Folha Pag. S.S." : "□ Folha Pag. S.S.");
            cmbTipoDocumento.Items.Add(documentosAnexados[3] ? "✓ Apólice AT" : "□ Apólice AT");
            cmbTipoDocumento.Items.Add(documentosAnexados[4] ? "✓ Apólice RC" : "□ Apólice RC");
            cmbTipoDocumento.Items.Add(documentosAnexados[5] ? "✓ Horário Trabalho" : "□ Horário Trabalho");
            cmbTipoDocumento.Items.Add(documentosAnexados[6] ? "✓ Anexo D" : "□ Anexo D");
            cmbTipoDocumento.Items.Add(documentosAnexados[7] ? "✓ Dec. Trab. Emigrantes" : "□ Dec. Trab. Emigrantes");
            cmbTipoDocumento.Items.Add(documentosAnexados[8] ? "✓ Inscrição SS" : "□ Inscrição SS");
            cmbTipoDocumento.Items.Add("Outro documento");
        }

        private void btnVerificarDocumentosFaltantes_Click(object sender, EventArgs e)
        {
            _empresaManager.VerificarDocumentosFaltantes();
        }

        private void btnAbrirPastaAnexos_Click(object sender, EventArgs e)
        {
            _empresaManager.AbrirPastaAnexos();
        }

        private void btnCancelarAnexo_Click(object sender, EventArgs e)
        {
            // Fechar o modal
            panelModalDocumentos.Visible = false;
        }

        private void btnConfirmarAnexo_Click(object sender, EventArgs e)
        {
            if (cmbTipoDocumento.SelectedIndex == -1)
            {
                MessageBox.Show("Por favor, selecione um tipo de documento.",
                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Fechar o modal
            panelModalDocumentos.Visible = false;

            // Anexar o documento de acordo com o tipo selecionado
            string tipoSelecionado = cmbTipoDocumento.SelectedItem.ToString();

            // Remover o prefixo de status (✓ ou □) se estiver presente
            if (tipoSelecionado.StartsWith("✓ ") || tipoSelecionado.StartsWith("□ "))
                tipoSelecionado = tipoSelecionado.Substring(2);

            switch (tipoSelecionado)
            {
                case "Não Div. Financas":
                    _empresaManager.AnexarDocumentoFinancas();
                    break;
                case "Não Div. Seg. Social":
                    _empresaManager.AnexarDocumentoSegSocial();
                    break;
                case "Folha Pag. S.S.":
                    _empresaManager.AnexarFolhaPag();
                    break;
                case "Apólice AT":
                    _empresaManager.AnexarDocumentoApoliceAT();
                    break;
                case "Apólice RC":
                    _empresaManager.AnexarDocumentoApoliceRC();
                    break;
                case "Horário Trabalho":
                    _empresaManager.AnexarHorarioTrabalho();
                    break;
                case "Anexo D":
                    _empresaManager.AnexarAnexoD();
                    break;
                case "Dec. Trab. Emigrantes":
                    _empresaManager.AnexarDecTrabEmigr();
                    break;
                case "Inscrição SS":
                    _empresaManager.AnexarInscricaoSS();
                    break;
                case "Outro documento":
                    _empresaManager.AnexarDocumento();
                    break;
            }
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

        private void btnAbrirPastaAnexos_Click_1(object sender, EventArgs e)
        {
            _empresaManager.AbrirPastaAnexos();
        }
    }
}