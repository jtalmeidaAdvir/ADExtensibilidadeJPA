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
            panelModalDocumentos.Visible = true;/*
            // Limpar campos
            dtpValidade.Checked = false;

            // Posicionar o modal no centro do formulário
            panelModalDocumentos.Location = new Point(
                (this.ClientSize.Width - panelModalDocumentos.Width) / 2,
                (this.ClientSize.Height - panelModalDocumentos.Height) / 2);

            // Exibir o modal
            panelModalDocumentos.Visible = true;
            panelModalDocumentos.BringToFront();

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

            // Verificar se a validade foi informada
            if (!dtpValidade.Checked)
            {
                MessageBox.Show("Por favor, informe a data de validade do documento.",
                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Obter validade
            DateTime? validade = dtpValidade.Checked ? dtpValidade.Value : (DateTime?)null;

            // Fechar o modal
            panelModalDocumentos.Visible = false;

            // Anexar o documento de acordo com o tipo selecionado
            string tipoSelecionado = cmbTipoDocumento.SelectedItem.ToString();

            // Remover o prefixo de status (✓ ou □) se estiver presente
            if (tipoSelecionado.StartsWith("✓ ") || tipoSelecionado.StartsWith("□ "))
                tipoSelecionado = tipoSelecionado.Substring(2);

            try
            {
                // Atualizar os campos de validade específicos e armazenar no banco de dados
                switch (tipoSelecionado)
                {
                    case "Não Div. Financas":
                        if (validade.HasValue)
                        {
                            TXT_NaoDivFinancas.Value = validade.Value;
                            TXT_NaoDivFinancas.Checked = true;
                            // Atualizar diretamente no banco de dados
                            string queryFinancas = $"UPDATE Geral_Entidade SET CDU_NaoDivFinancas = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryFinancas);
                        }
                        _empresaManager.AnexarDocumentoFinancas();
                        break;
                    case "Não Div. Seg. Social":
                        if (validade.HasValue)
                        {
                            TXT_NaoDivSegSocial.Value = validade.Value;
                            TXT_NaoDivSegSocial.Checked = true;
                            // Atualizar diretamente no banco de dados
                            string querySegSocial = $"UPDATE Geral_Entidade SET CDU_NaoDivSegSocial = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(querySegSocial);
                        }
                        _empresaManager.AnexarDocumentoSegSocial();
                        break;
                    case "Folha Pag. S.S.":
                        if (validade.HasValue)
                        {
                            TXT_FolhaPagSegSocial.Value = validade.Value;
                            TXT_FolhaPagSegSocial.Checked = true;
                            // Atualizar diretamente no banco de dados
                            string queryFolhaPag = $"UPDATE Geral_Entidade SET CDU_FolhaPagSegSocial = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryFolhaPag);
                        }
                        _empresaManager.AnexarFolhaPag();
                        break;
                    case "Horário Trabalho":
                        // Criar uma query para atualizar a validade do horário de trabalho
                        if (validade.HasValue)
                        {
                            string queryHorarioTrabalho = $"UPDATE Geral_Entidade SET CDU_ValidadeHorarioTrabalho = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryHorarioTrabalho);
                        }
                        _empresaManager.AnexarHorarioTrabalho("", validade);
                        break;
                    case "Anexo D":
                        // Criar uma query para atualizar a validade do Anexo D
                        if (validade.HasValue)
                        {

                

                            string queryAnexoD = $"UPDATE Geral_Entidade SET CDU_ValidadeAnexoD = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryAnexoD);
                        }
                        _empresaManager.AnexarAnexoD("", validade);
                        break;
                    case "Dec. Trab. Emigrantes":
                        // Criar uma query para atualizar a validade da declaração de trabalhadores emigrantes
                        if (validade.HasValue)
                        {
                            string queryDecTrabEmigr = $"UPDATE Geral_Entidade SET CDU_ValidadeDecTrabEmigr = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryDecTrabEmigr);
                        }
                        _empresaManager.AnexarDecTrabEmigr("", validade);
                        break;
                    case "Inscrição SS":
                        // Criar uma query para atualizar a validade da inscrição na segurança social
                        if (validade.HasValue)
                        {
                            string queryInscricaoSS = $"UPDATE Geral_Entidade SET CDU_ValidadeInscricaoSS = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryInscricaoSS);
                        }
                        _empresaManager.AnexarInscricaoSS("", validade);
                        break;
                    case "Outro documento":
                        // Para outros documentos, poderia-se criar um registro em outra tabela
                        _empresaManager.AnexarDocumento("", validade);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar a validade do documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void cmbTipoDocumento_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTipoDocumento.SelectedIndex == -1)
                return;

            string tipoSelecionado = cmbTipoDocumento.SelectedItem.ToString();

            // Remover o prefixo de status (✓ ou □) se estiver presente
            if (tipoSelecionado.StartsWith("✓ ") || tipoSelecionado.StartsWith("□ "))
                tipoSelecionado = tipoSelecionado.Substring(2);

            try
            {
                // Carregar a validade do documento selecionado
                string query = "";
                switch (tipoSelecionado)
                {
                    case "Não Div. Financas":
                        query = $"SELECT CDU_NaoDivFinancas FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Não Div. Seg. Social":
                        query = $"SELECT CDU_NaoDivSegSocial FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Folha Pag. S.S.":
                        query = $"SELECT CDU_FolhaPagSegSocial FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Apólice AT":
                        query = $"SELECT CDU_ReciboApoliceAT FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Apólice RC":
                        query = $"SELECT CDU_ReciboRC FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Horário Trabalho":
                        query = $"SELECT CDU_ValidadeHorarioTrabalho FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Anexo D":
                        query = $"SELECT CDU_ValidadeAnexoD FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Dec. Trab. Emigrantes":
                        query = $"SELECT CDU_ValidadeDecTrabEmigr FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Inscrição SS":
                        query = $"SELECT CDU_ValidadeInscricaoSS FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    default:
                        // Para "Outro documento" não há campo específico de validade
                        dtpValidade.Checked = false;
                        return;
                }

                var resultado = BSO.Consulta(query);
                if (resultado.NumLinhas() > 0)
                {
                    resultado.Inicio();
                    string dataStr = resultado.Valor(0);

                    if (tipoSelecionado == "Apólice AT" || tipoSelecionado == "Apólice RC")
                    {
                        // Extrair data de strings como "123456 (Val: 01/01/2023)"
                        if (!string.IsNullOrEmpty(dataStr) && dataStr.Contains("(Val:"))
                        {
                            int inicio = dataStr.IndexOf("(Val:") + 6;
                            int fim = dataStr.IndexOf(")", inicio);
                            if (fim > inicio)
                            {
                                dataStr = dataStr.Substring(inicio, fim - inicio);
                                DateTime data;
                                if (DateTime.TryParse(dataStr, out data))
                                {
                                    dtpValidade.Value = data;
                                    dtpValidade.Checked = true;
                                    return;
                                }
                            }
                        }
                        dtpValidade.Checked = false;
                    }
                    else
                    {
                        // Para outros campos que são diretamente datas
                        if (!string.IsNullOrEmpty(dataStr))
                        {
                            DateTime data;
                            if (DateTime.TryParse(dataStr, out data))
                            {
                                dtpValidade.Value = data;
                                dtpValidade.Checked = true;
                                return;
                            }
                        }
                    }
                }

                // Se não encontrou data ou não conseguiu converter
                dtpValidade.Checked = false;

            }
            catch (Exception ex)
            {
                // Em caso de erro, apenas não mostra a data
                dtpValidade.Checked = false;
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