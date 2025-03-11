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

                // Mostrar o DataGridView garantindo que seja visível
                dataGridView1.Visible = true;

                // Limpar datagrid e recarregar os dados da obra selecionada
                dataGridView1.Rows.Clear();
                _empresaManager.CarregarObrasEmDataGridView(codigoObraSelecionada);

                // Formatar DataGridView
                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.ClearSelection();

                    // Destacar células conforme o status
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells["StatusAutorizacao"].Value != null)
                        {
                            string status = row.Cells["StatusAutorizacao"].Value.ToString();

                            switch (status)
                            {
                                case "Autorizado":
                                    row.Cells["StatusAutorizacao"].Style.BackColor = Color.LightGreen;
                                    row.Cells["StatusAutorizacao"].Style.ForeColor = Color.DarkGreen;
                                    break;
                                case "Pendente":
                                    row.Cells["StatusAutorizacao"].Style.BackColor = Color.LightYellow;
                                    row.Cells["StatusAutorizacao"].Style.ForeColor = Color.DarkOrange;
                                    break;
                                case "Não Autorizado":
                                    row.Cells["StatusAutorizacao"].Style.BackColor = Color.LightCoral;
                                    row.Cells["StatusAutorizacao"].Style.ForeColor = Color.DarkRed;
                                    break;
                                case "Renovação Necessária":
                                    row.Cells["StatusAutorizacao"].Style.BackColor = Color.LightSalmon;
                                    row.Cells["StatusAutorizacao"].Style.ForeColor = Color.Brown;
                                    break;
                                case "Documentos Faltantes":
                                    row.Cells["StatusAutorizacao"].Style.BackColor = Color.LightPink;
                                    row.Cells["StatusAutorizacao"].Style.ForeColor = Color.Maroon;
                                    break;
                            }
                        }
                    }
                }

                // Atualizar o controle de autorizações quando seleciona uma obra
                AtualizarControleAutorizacaoObra(codigoObraSelecionada);

                // Verificar se existem botões ou painéis de nova entrada e escondê-los
                Button btnConfirmar = groupBoxObras.Controls["btnConfirmar"] as Button;
                if (btnConfirmar != null)
                    btnConfirmar.Visible = false;

                Panel panelNovaEntrada = groupBoxObras.Controls["panelNovaEntrada"] as Panel;
                if (panelNovaEntrada != null)
                    panelNovaEntrada.Visible = false;

                // Esconder campos de entrada
                lblDataEntrada.Visible = false;
                dtpDataEntrada.Visible = false;
                lblDataSaida.Visible = false;
                dtpDataSaida.Visible = false;
                lblContratoSubempreitada.Visible = false;
                txtContratoSubempreitada.Visible = false;
                lblStatusEntrada.Visible = false;
                cmbStatusEntrada.Visible = false;
                pnlDadosObra.Visible = false;

                // Mostrar botão gravar
                btnGravarObra.Visible = true;
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
        }

        private void UpdateDocumentComboBox()
        {
            // Limpar os itens existentes
            cmbTipoDocumento.Items.Clear();

            // Verificar quais documentos já estão anexados
            bool[] documentosAnexados = _empresaManager.GetDocumentosAnexados();

            // Adicionar os itens com prefixo indicando status
            cmbTipoDocumento.Items.Add(documentosAnexados[0] ? "✓ Certidão de não divida às Finanças" : "□ Certidão de não divida às Finanças");
            cmbTipoDocumento.Items.Add(documentosAnexados[1] ? "✓ Certidão de não divida à Segurança Social" : "□ Certidão de não divida à Segurança Social");
            cmbTipoDocumento.Items.Add(documentosAnexados[2] ? "✓ Folha de remuneração à segurança social do mês corrente com o nome dos funcionários a colocar em obra + comprovativo de pagamento" : "□ Folha de remuneração à segurança social do mês corrente com o nome dos funcionários a colocar em obra + comprovativo de pagamento");
            cmbTipoDocumento.Items.Add(documentosAnexados[3] ? "✓ Recibo do seguro de acidentes de trabalho" : "□ Recibo do seguro de acidentes de trabalho");
            cmbTipoDocumento.Items.Add(documentosAnexados[4] ? "✓ Seguro de responsabilidade civil" : "□ Seguro de responsabilidade civil");
            cmbTipoDocumento.Items.Add(documentosAnexados[5] ? "✓ Horário trabalho para a empreitada acima designada" : "□ Horário trabalho para a empreitada acima designada");
            cmbTipoDocumento.Items.Add(documentosAnexados[6] ? "✓ Condições particulares do seguro de acidentes de trabalho" : "□ Condições particulares do seguro de acidentes de trabalho");
            cmbTipoDocumento.Items.Add(documentosAnexados[7] ? "✓ Alvará/Certificado de construção ou alvará específico para a atividade (ex. trabalho temporário)" : "□ Alvará/Certificado de construção ou alvará específico para a atividade (ex. trabalho temporário)");
            cmbTipoDocumento.Items.Add(documentosAnexados[8] ? "✓ Certidão permanente" : "□ Certidão permanente");
            cmbTipoDocumento.Items.Add("Contrato de subcontratação/subempreitada/nota de encomenda");
            cmbTipoDocumento.Items.Add("Declaração de adesão ao PSS (segue em anexo modelo de declaração a preencher)");
            cmbTipoDocumento.Items.Add("Declaração do responsável no estaleiro (segue em anexo modelo de declaração a preencher)");
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
                    case "Certidão de não divida às Finanças":
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
                    case "Certidão de não divida à Segurança Social":
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
                    case "Folha de remuneração à segurança social do mês corrente com o nome dos funcionários a colocar em obra + comprovativo de pagamento":
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
                    case "Recibo do seguro de acidentes de trabalho":
                        if (validade.HasValue)
                        {
                            string queryApoliceAT = $"UPDATE Geral_Entidade SET CDU_ValidadeApoliceAT = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryApoliceAT);
                        }
                        _empresaManager.AnexarDocumentoApoliceAT();
                        break;
                    case "Seguro de responsabilidade civil":
                        if (validade.HasValue)
                        {
                            string queryApoliceRC = $"UPDATE Geral_Entidade SET CDU_ValidadeApoliceRC = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryApoliceRC);
                        }
                        _empresaManager.AnexarDocumentoApoliceRC();
                        break;
                    case "Horário trabalho para a empreitada acima designada":
                        if (validade.HasValue)
                        {
                            string queryHorarioTrabalho = $"UPDATE Geral_Entidade SET CDU_ValidadeHorarioTrabalho = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryHorarioTrabalho);
                        }
                        _empresaManager.AnexarHorarioTrabalho("", validade);
                        break;
                    case "Alvará/Certificado de construção ou alvará específico para a atividade (ex. trabalho temporário)":
                        if (validade.HasValue)
                        {
                            TXT_AlvaraValidade.Value = validade.Value;
                            TXT_AlvaraValidade.Checked = true;
                            string queryAlvara = $"UPDATE Geral_Entidade SET CDU_ValidadeAlvara = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryAlvara);
                        }
                        _empresaManager.AnexarDocumento("AlvaraConstrucao", validade);
                        break;
                    case "Certidão permanente":
                        if (validade.HasValue)
                        {
                            string queryCertidaoPermanente = $"UPDATE Geral_Entidade SET CDU_ValidadeCertidaoPermanente = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryCertidaoPermanente);
                        }
                        _empresaManager.AnexarDocumento("CertidaoPermanente", validade);
                        break;
                    case "Contrato de subcontratação/subempreitada/nota de encomenda":
                        if (validade.HasValue)
                        {
                            string queryContrato = $"UPDATE Geral_Entidade SET CDU_ValidadeContratoSubcontratacao = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryContrato);
                        }
                        _empresaManager.AnexarDocumento("ContratoSubcontratacao", validade);
                        break;
                    case "Condições particulares do seguro de acidentes de trabalho":
                        if (validade.HasValue)
                        {



                            string queryCondicoesAT = $"UPDATE Geral_Entidade SET CDU_ValidadeCondicoesAT = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryCondicoesAT);
                        }
                        _empresaManager.AnexarDocumento("CondicoesAT", validade);
                        break;
                    case "Declaração de adesão ao PSS (segue em anexo modelo de declaração a preencher)":
                        if (validade.HasValue)
                        {
                            string queryDeclaracaoPSS = $"UPDATE Geral_Entidade SET CDU_ValidadeDeclaracaoPSS = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryDeclaracaoPSS);
                        }
                        _empresaManager.AnexarDocumento("DeclaracaoPSS", validade);
                        break;
                    case "Declaração do responsável no estaleiro (segue em anexo modelo de declaração a preencher)":
                        if (validade.HasValue)
                        {
                            string queryDeclaracaoResponsavel = $"UPDATE Geral_Entidade SET CDU_ValidadeDeclaracaoResponsavel = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            BSO.DSO.ExecuteSQL(queryDeclaracaoResponsavel);
                        }
                        _empresaManager.AnexarDocumento("DeclaracaoResponsavel", validade);
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
        private void rbAutorizacao_CheckedChanged(object sender, EventArgs e)
        {
            // Make sure we have a valid obra selection
            if (cb_Obras.SelectedItem is KeyValuePair<string, string> obraSelecionada)
            {
                string codigoObraSelecionada = obraSelecionada.Key;

                // Determine which radio button was checked
                RadioButton rb = sender as RadioButton;
                if (rb != null && rb.Checked)
                {
                    // If the Autorizado radio button is checked, set autorizado to true
                    // Otherwise, set it to false (Não Autorizado)
                    bool autorizado = (rb.Name == "rbAutorizado");

                    // Update the authorization status
                    _empresaManager.AtualizarAutorizacaoObra(codigoObraSelecionada, autorizado);
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma obra primeiro.",
                    "Nenhuma obra selecionada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                // Reset radio button states to avoid confusion
                RadioButton rbAutorizado = pnlAutorizacaoObra.Controls["rbAutorizado"] as RadioButton;
                RadioButton rbNaoAutorizado = pnlAutorizacaoObra.Controls["rbNaoAutorizado"] as RadioButton;
                if (rbAutorizado != null && rbNaoAutorizado != null)
                {
                    rbAutorizado.Checked = false;
                    rbNaoAutorizado.Checked = false;
                }
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
                    case "Certidão de não divida às Finanças":
                        query = $"SELECT CDU_NaoDivFinancas FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Certidão de não divida à Segurança Social":
                        query = $"SELECT CDU_NaoDivSegSocial FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Folha de remuneração à segurança social do mês corrente com o nome dos funcionários a colocar em obra + comprovativo de pagamento":
                        query = $"SELECT CDU_FolhaPagSegSocial FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Apólice AT":
                        query = $"SELECT CDU_ReciboApoliceAT FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Apólice RC":
                        query = $"SELECT CDU_ReciboRC FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Horário trabalho para a empreitada acima designada":
                        query = $"SELECT CDU_ValidadeHorarioTrabalho FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Alvará/Certificado de construção ou alvará específico para a atividade (ex. trabalho temporário)":
                        query = $"SELECT CDU_ValidadeAlvara FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Certidão permanente":
                        query = $"SELECT CDU_ValidadeCertidaoPermanente FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Contrato de subcontratação/subempreitada/nota de encomenda":
                        query = $"SELECT CDU_ValidadeContratoSubcontratacao FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Condições particulares do seguro AT":
                        query = $"SELECT CDU_ValidadeCondicoesAT FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Declaração de adesão ao PSS (segue em anexo modelo de declaração a preencher)":
                        query = $"SELECT CDU_ValidadeDeclaracaoPSS FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
                        break;
                    case "Declaração do responsável no estaleiro (segue em anexo modelo de declaração a preencher)":
                        query = $"SELECT CDU_ValidadeDeclaracaoResponsavel FROM Geral_Entidade WHERE ID = '{_empresaManager.IdSelecionado}'";
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

        private void AtualizarControleAutorizacaoObra(string codigoObra)
        {
            // Buscar o status de autorização atual da obra selecionada
            KeyValuePair<int, string> statusInfo = _empresaManager.ObterStatusAutorizacaoObra(codigoObra);

            // Atualizar o controle no formulário
            if (pnlAutorizacaoObra != null && pnlAutorizacaoObra.Visible)
            {
                ComboBox cmbAutorizacaoStatus = pnlAutorizacaoObra.Controls["cmbAutorizacaoStatus"] as ComboBox;
                TextBox txtObservacao = pnlAutorizacaoObra.Controls["txtObservacao"] as TextBox;

                if (cmbAutorizacaoStatus != null && txtObservacao != null)
                {
                    // Converter o status para o índice correto
                    // Status: 0=Autorizado, 1=Pendente, 2=Não Autorizado, 3=Renovação Necessária, 4=Documentos Faltantes
                    if (statusInfo.Key >= 0 && statusInfo.Key < cmbAutorizacaoStatus.Items.Count)
                    {
                        cmbAutorizacaoStatus.SelectedIndex = statusInfo.Key;
                    }
                    else
                    {
                        cmbAutorizacaoStatus.SelectedIndex = 1; // Padrão "Pendente"
                    }

                    // Definir as observações
                    txtObservacao.Text = statusInfo.Value;

                    // Destacar o combobox conforme o status
                    switch (statusInfo.Key)
                    {
                        case 0: // Autorizado
                            cmbAutorizacaoStatus.BackColor = Color.LightGreen;
                            break;
                        case 1: // Pendente
                            cmbAutorizacaoStatus.BackColor = Color.LightYellow;
                            break;
                        case 2: // Não Autorizado
                            cmbAutorizacaoStatus.BackColor = Color.LightCoral;
                            break;
                        case 3: // Renovação Necessária
                            cmbAutorizacaoStatus.BackColor = Color.LightSalmon;
                            break;
                        case 4: // Documentos Faltantes
                            cmbAutorizacaoStatus.BackColor = Color.LightPink;
                            break;
                        default:
                            cmbAutorizacaoStatus.BackColor = SystemColors.Window;
                            break;
                    }
                }
            }

            // Também atualizar o DataGridView com os dados atualizados da obra
            _empresaManager.CarregarObrasEmDataGridView(codigoObra);
        }

        private void cmbAutorizacaoStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Mudança de cor conforme status
            ComboBox cmb = sender as ComboBox;
            if (cmb != null)
            {
                switch (cmb.SelectedIndex)
                {
                    case 0: // Autorizado
                        cmb.BackColor = Color.LightGreen;
                        break;
                    case 1: // Pendente
                        cmb.BackColor = Color.LightYellow;
                        break;
                    case 2: // Não Autorizado
                        cmb.BackColor = Color.LightCoral;
                        break;
                    case 3: // Renovação Necessária
                        cmb.BackColor = Color.LightSalmon;
                        break;
                    case 4: // Documentos Faltantes
                        cmb.BackColor = Color.LightPink;
                        break;
                    default:
                        cmb.BackColor = SystemColors.Window;
                        break;
                }
            }
        }

        private void btnSalvarAutorizacao_Click(object sender, EventArgs e)
        {
            // Verificar se uma obra está selecionada
            if (cb_Obras.SelectedItem is KeyValuePair<string, string> obraSelecionada)
            {
                string codigoObraSelecionada = obraSelecionada.Key;

                ComboBox cmbAutorizacaoStatus = pnlAutorizacaoObra.Controls["cmbAutorizacaoStatus"] as ComboBox;
                TextBox txtObservacao = pnlAutorizacaoObra.Controls["txtObservacao"] as TextBox;

                if (cmbAutorizacaoStatus != null && txtObservacao != null)
                {
                    int statusIndex = cmbAutorizacaoStatus.SelectedIndex;
                    string observacao = txtObservacao.Text;

                    // Salvar status e observação na base de dados
                    _empresaManager.AtualizarStatusAutorizacaoObra(codigoObraSelecionada, statusIndex, observacao);
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma obra primeiro.",
                    "Nenhuma obra selecionada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
        #endregion

        private void btnAbrirPastaAnexos_Click_1(object sender, EventArgs e)
        {
            _empresaManager.AbrirPastaAnexos();
        }

        private void btnAutorizarEntrada_Click(object sender, EventArgs e)
        {
            // Verifica se uma obra está selecionada
            if (cb_Obras.SelectedItem == null)
            {
                MessageBox.Show("Por favor, selecione uma obra primeiro.",
                    "Seleção de Obra",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // Esconder o DataGridView temporariamente e botão gravar
            btnGravarObra.Visible = false;
            dataGridView1.Visible = false;

            // Criar botão confirmar se não existir
            Button btnConfirmar = groupBoxObras.Controls["btnConfirmar"] as Button;
            if (btnConfirmar == null)
            {
                btnConfirmar = new Button
                {
                    Name = "btnConfirmar",
                    Text = "Confirmar",
                    BackColor = Color.LightGreen,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Calibri", 10F, FontStyle.Bold),
                    Size = new Size(100, 30),
                    Location = new Point(550, 95)
                };
                btnConfirmar.Click += new EventHandler(btnConfirmar_Click);
                groupBoxObras.Controls.Add(btnConfirmar);
                btnConfirmar.BringToFront();
            }
            else
            {
                btnConfirmar.Visible = true;
                btnConfirmar.BringToFront();
            }

            // Configurar e mostrar os controles dentro do formulário principal - sem depender do painel
            lblDataEntrada.Parent = groupBoxObras;
            dtpDataEntrada.Parent = groupBoxObras;
            lblDataSaida.Parent = groupBoxObras;
            dtpDataSaida.Parent = groupBoxObras;
            lblContratoSubempreitada.Parent = groupBoxObras;
            txtContratoSubempreitada.Parent = groupBoxObras;
            lblStatusEntrada.Parent = groupBoxObras;
            cmbStatusEntrada.Parent = groupBoxObras;

            // Posicionar os controles corretamente
            lblDataEntrada.Location = new Point(25, 95);
            dtpDataEntrada.Location = new Point(115, 95);
            lblDataSaida.Location = new Point(235, 95);
            dtpDataSaida.Location = new Point(315, 95);

            lblContratoSubempreitada.Location = new Point(25, 120);
            txtContratoSubempreitada.Location = new Point(165, 120);
            txtContratoSubempreitada.Size = new Size(250, 22);

            lblStatusEntrada.Location = new Point(25, 155);
            cmbStatusEntrada.Location = new Point(90, 155);
            cmbStatusEntrada.Size = new Size(150, 22);

            btnConfirmar.Location = new Point(515, 140);

            // Tornar os controles visíveis
            lblDataEntrada.Visible = true;
            dtpDataEntrada.Visible = true;
            lblDataSaida.Visible = true;
            dtpDataSaida.Visible = true;
            lblContratoSubempreitada.Visible = true;
            txtContratoSubempreitada.Visible = true;
            lblStatusEntrada.Visible = true;
            cmbStatusEntrada.Visible = true;

            // Titulo para a área de dados
            Label lblTituloNovaEntrada = groupBoxObras.Controls["lblTituloNovaEntrada"] as Label;
            if (lblTituloNovaEntrada == null)
            {
                lblTituloNovaEntrada = new Label
                {
                    Name = "lblTituloNovaEntrada",
                    Text = "AUTORIZAÇÃO DE NOVA ENTRADA EM OBRA",
                    Font = new Font("Calibri", 11F, FontStyle.Bold),
                    AutoSize = true,
                    Location = new Point(200, 70),
                    BackColor = Color.Transparent
                };
                groupBoxObras.Controls.Add(lblTituloNovaEntrada);
            }
            else
            {
                lblTituloNovaEntrada.Visible = true;
            }
            lblTituloNovaEntrada.BringToFront();

            // Destaque visual
            lblStatusEntrada.Font = new Font("Calibri", 9, FontStyle.Bold);
            lblStatusEntrada.ForeColor = Color.Firebrick;
            lblStatusEntrada.Text = "Status:";

            // Definir valores padrão
            dtpDataEntrada.Value = DateTime.Today;
            dtpDataSaida.Value = DateTime.Today.AddMonths(1);
            txtContratoSubempreitada.Text = "";

            // Certificar que o ComboBox de status tem itens e selecionar "Autorizado" como padrão
            if (cmbStatusEntrada.Items.Count == 0)
            {
                // Adicionar itens caso não existam
                cmbStatusEntrada.Items.AddRange(new object[] {
                    "Autorizado",
                    "Pendente",
                    "Não Autorizado",
                    "Renovação Necessária",
                    "Documentos Faltantes"
                });
            }
            cmbStatusEntrada.SelectedIndex = 0; // Status padrão "Autorizado"

            // Destaque visual dos campos
            cmbStatusEntrada.BackColor = Color.LightGreen;
            txtContratoSubempreitada.BackColor = Color.LightYellow;

            // Desenhar um retângulo como fundo visual
            Panel backgroundPanel = groupBoxObras.Controls["backgroundPanel"] as Panel;
            if (backgroundPanel == null)
            {
                backgroundPanel = new Panel
                {
                    Name = "backgroundPanel",
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = Color.WhiteSmoke,
                    Size = new Size(655, 120),
                    Location = new Point(20, 90)
                };
                groupBoxObras.Controls.Add(backgroundPanel);
                backgroundPanel.SendToBack();
            }
            else
            {
                backgroundPanel.Visible = true;
                backgroundPanel.SendToBack();
            }

            // Focar no primeiro campo
            dtpDataEntrada.Focus();
        }

        private void btnConfirmar_Click(object sender, EventArgs e)
        {
            // Verificar se o status foi selecionado (obrigatório)
            if (cmbStatusEntrada.SelectedIndex == -1)
            {
                MessageBox.Show("Por favor, selecione um status de autorização.",
                    "Status obrigatório",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                cmbStatusEntrada.Focus();
                return;
            }

            // Verificar se o contrato foi preenchido
            if (string.IsNullOrEmpty(txtContratoSubempreitada.Text))
            {
                MessageBox.Show("Por favor, informe o número do contrato de subempreitada.",
                    "Campo obrigatório",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                txtContratoSubempreitada.Focus();
                return;
            }

            // Obter o status selecionado
            string statusText = cmbStatusEntrada.SelectedItem.ToString();
            int statusIndex = cmbStatusEntrada.SelectedIndex;

            // Restaurar interface - mostrar o DataGridView novamente
            dataGridView1.Visible = true;

            // Adicionar os dados diretamente ao DataGridView incluindo o status
            int rowIndex = dataGridView1.Rows.Add(
                dtpDataEntrada.Value.ToString("dd/MM/yyyy"),
                dtpDataSaida.Value.ToString("dd/MM/yyyy"),
                txtContratoSubempreitada.Text,
                statusText
            );

            // Marcar a autorização conforme o status
            dataGridView1.Rows[rowIndex].Cells["AutorizacaoEntrada"].Value = statusIndex == 0 ? true : false;

            // Destacar a linha adicionada
            dataGridView1.Rows[rowIndex].DefaultCellStyle.BackColor = Color.AliceBlue;

            // Destacar célula de status conforme a seleção
            switch (statusIndex)
            {
                case 0: // Autorizado
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.BackColor = Color.LightGreen;
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.ForeColor = Color.DarkGreen;
                    break;
                case 1: // Pendente
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.BackColor = Color.LightYellow;
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.ForeColor = Color.DarkOrange;
                    break;
                case 2: // Não Autorizado
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.BackColor = Color.LightCoral;
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.ForeColor = Color.DarkRed;
                    break;
                case 3: // Renovação Necessária
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.BackColor = Color.LightSalmon;
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.ForeColor = Color.Brown;
                    break;
                case 4: // Documentos Faltantes
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.BackColor = Color.LightPink;
                    dataGridView1.Rows[rowIndex].Cells["StatusAutorizacao"].Style.ForeColor = Color.Maroon;
                    break;
            }

            // Ocultar os campos após adicionar
            lblDataEntrada.Visible = false;
            dtpDataEntrada.Visible = false;
            lblDataSaida.Visible = false;
            dtpDataSaida.Visible = false;
            lblContratoSubempreitada.Visible = false;
            txtContratoSubempreitada.Visible = false;
            lblStatusEntrada.Visible = false;
            cmbStatusEntrada.Visible = false;

            // Ocultar elementos adicionais
            Button btnConfirmar = groupBoxObras.Controls["btnConfirmar"] as Button;
            if (btnConfirmar != null)
                btnConfirmar.Visible = false;

            Label lblTituloNovaEntrada = groupBoxObras.Controls["lblTituloNovaEntrada"] as Label;
            if (lblTituloNovaEntrada != null)
                lblTituloNovaEntrada.Visible = false;

            Panel backgroundPanel = groupBoxObras.Controls["backgroundPanel"] as Panel;
            if (backgroundPanel != null)
                backgroundPanel.Visible = false;

            btnGravarObra.Visible = true;

            // Verificar se uma obra está selecionada
            if (cb_Obras.SelectedItem is KeyValuePair<string, string> obraSelecionada)
            {
                string codigoObraSelecionada = obraSelecionada.Key;

                // Criar uma observação detalhada
                string observacao = $"Autorização: {statusText}. Entrada: {dtpDataEntrada.Value:dd/MM/yyyy}, Saída: {dtpDataSaida.Value:dd/MM/yyyy}, Contrato: {txtContratoSubempreitada.Text}";

                // Atualizar o status de autorização para a obra
                _empresaManager.AtualizarStatusAutorizacaoObra(codigoObraSelecionada, statusIndex, observacao);

                // Atualizar a visualização de status
                AtualizarControleAutorizacaoObra(codigoObraSelecionada);

                // Salvar os dados imediatamente
                _empresaManager.SalvarObra();
            }

            // Mensagem de sucesso
            MessageBox.Show("Autorização de entrada registrada com sucesso!",
                "Sucesso",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // Método para processar dados de obra e chamar a gravação
        private void ProcessarGravarObra(object sender, EventArgs e)
        {
            // Se os campos de autorização estiverem visíveis, cancelar operação
            // pois deveria usar o botão confirmar
            if (lblDataEntrada.Visible)
            {
                MessageBox.Show("Por favor, use o botão 'Confirmar' para salvar a autorização de entrada.",
                    "Utilizar botão Confirmar",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // Chamar o método original para salvar
            _empresaManager.SalvarObra();
        }
    }
}