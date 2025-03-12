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
using Microsoft.Win32;

namespace ADExtensibilidadeJPA
{
    public partial class Menu : Form
    {
        private readonly ErpBS _BSO;
        private readonly StdBSInterfPub _PSO;
        private readonly string _idSelecionado;
        private EmpresaManager _empresaManager;
        private TrabalhadorManager _trabalhadorManager;

        public Menu()
        {
            InitializeComponent();
        }

        public Menu(ErpBS BSO, StdBSInterfPub PSO, string idSelecionado)
        {
            InitializeComponent();
            _BSO = BSO;
            _PSO = PSO;
            _idSelecionado = idSelecionado;
            _empresaManager = new EmpresaManager(_BSO, _PSO, _idSelecionado, this);
            _trabalhadorManager = new TrabalhadorManager(_BSO, _PSO, _idSelecionado, this);
            ConfigurarInterface();
            ConfigurarEventosTrabalhadores();
        }

        private void ConfigurarInterface()
        {
            // Configuração da interface do usuário, especialmente para os DateTimePickers
            // Definir valores iniciais para os DateTimePickers
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

        private void ConfigurarEventosTrabalhadores()
        {
            // Configurar eventos para a aba de Trabalhadores

            // Configurar evento para credenciação (habilitar/desabilitar campo de descrição)
            CheckBox chkCredenciacao = this.Controls.Find("chkCredenciacao", true).FirstOrDefault() as CheckBox;
            TextBox txtCredenciacao = this.Controls.Find("txtCredenciacao", true).FirstOrDefault() as TextBox;

            if (chkCredenciacao != null && txtCredenciacao != null)
            {
                chkCredenciacao.CheckedChanged += (sender, e) => {
                    txtCredenciacao.Enabled = chkCredenciacao.Checked;
                };
            }

            // Configurar eventos de botões
            Button btnAdicionarTrabalhador = this.Controls.Find("btnAdicionarTrabalhador", true).FirstOrDefault() as Button;
            Button btnEditarTrabalhador = this.Controls.Find("btnEditarTrabalhador", true).FirstOrDefault() as Button;
            Button btnExcluirTrabalhador = this.Controls.Find("btnExcluirTrabalhador", true).FirstOrDefault() as Button;
            Button btnSalvarTrabalhador = this.Controls.Find("btnSalvarTrabalhador", true).FirstOrDefault() as Button;
            Button btnAutorizarTrabalhador = this.Controls.Find("btnAutorizarTrabalhador", true).FirstOrDefault() as Button;

            if (btnAdicionarTrabalhador != null)
                btnAdicionarTrabalhador.Click += new EventHandler(btnAdicionarTrabalhador_Click);

            if (btnEditarTrabalhador != null)
                btnEditarTrabalhador.Click += new EventHandler(btnEditarTrabalhador_Click);

            if (btnExcluirTrabalhador != null)
                btnExcluirTrabalhador.Click += new EventHandler(btnExcluirTrabalhador_Click);

            if (btnSalvarTrabalhador != null)
                btnSalvarTrabalhador.Click += new EventHandler(btnSalvarTrabalhador_Click);

            if (btnAutorizarTrabalhador != null)
                btnAutorizarTrabalhador.Click += new EventHandler(btnAutorizarTrabalhador_Click);

            // Configurar eventos para botões de anexos
            Button btnAnexarFichaAptidao = this.Controls.Find("btnAnexarFichaAptidao", true).FirstOrDefault() as Button;
            Button btnAnexarCredenciacao = this.Controls.Find("btnAnexarCredenciacao", true).FirstOrDefault() as Button;
            Button btnAnexarFichaEPI = this.Controls.Find("btnAnexarFichaEPI", true).FirstOrDefault() as Button;

            if (btnAnexarFichaAptidao != null)
                btnAnexarFichaAptidao.Click += new EventHandler(btnAnexarFichaAptidao_Click);

            if (btnAnexarCredenciacao != null)
                btnAnexarCredenciacao.Click += new EventHandler(btnAnexarCredenciacao_Click);

            if (btnAnexarFichaEPI != null)
                btnAnexarFichaEPI.Click += new EventHandler(btnAnexarFichaEPI_Click);

            // Configurar evento de seleção de trabalhador na grid
            DataGridView gridTrabalhadores = this.Controls.Find("gridTrabalhadores", true).FirstOrDefault() as DataGridView;
            if (gridTrabalhadores != null)
            {
                gridTrabalhadores.CellDoubleClick += new DataGridViewCellEventHandler(gridTrabalhadores_CellDoubleClick);
            }

            // Configurar labels para visualizar anexos
            Label lblFichaAptidaoAnexo = this.Controls.Find("lblFichaAptidaoAnexo", true).FirstOrDefault() as Label;
            Label lblCredenciacaoAnexo = this.Controls.Find("lblCredenciacaoAnexo", true).FirstOrDefault() as Label;
            Label lblFichaEPIAnexo = this.Controls.Find("lblFichaEPIAnexo", true).FirstOrDefault() as Label;

            if (lblFichaAptidaoAnexo != null)
                lblFichaAptidaoAnexo.Click += new EventHandler(lblFichaAptidaoAnexo_Click);

            if (lblCredenciacaoAnexo != null)
                lblCredenciacaoAnexo.Click += new EventHandler(lblCredenciacaoAnexo_Click);

            if (lblFichaEPIAnexo != null)
                lblFichaEPIAnexo.Click += new EventHandler(lblFichaEPIAnexo_Click);
        }

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
                            _BSO.DSO.ExecuteSQL(queryFinancas);
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
                            _BSO.DSO.ExecuteSQL(querySegSocial);
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
                            _BSO.DSO.ExecuteSQL(queryFolhaPag);
                        }
                        _empresaManager.AnexarFolhaPag();
                        break;
                    case "Recibo do seguro de acidentes de trabalho":
                        if (validade.HasValue)
                        {
                            string queryApoliceAT = $"UPDATE Geral_Entidade SET CDU_ValidadeApoliceAT = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryApoliceAT);
                        }
                        _empresaManager.AnexarDocumentoApoliceAT();
                        break;
                    case "Seguro de responsabilidade civil":
                        if (validade.HasValue)
                        {
                            string queryApoliceRC = $"UPDATE Geral_Entidade SET CDU_ValidadeApoliceRC = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryApoliceRC);
                        }
                        _empresaManager.AnexarDocumentoApoliceRC();
                        break;
                    case "Horário trabalho para a empreitada acima designada":
                        if (validade.HasValue)
                        {
                            string queryHorarioTrabalho = $"UPDATE Geral_Entidade SET CDU_ValidadeHorarioTrabalho = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryHorarioTrabalho);
                        }
                        _empresaManager.AnexarHorarioTrabalho("", validade);
                        break;
                    case "Alvará/Certificado de construção ou alvará específico para a atividade (ex. trabalho temporário)":
                        if (validade.HasValue)
                        {
                            TXT_AlvaraValidade.Value = validade.Value;
                            TXT_AlvaraValidade.Checked = true;
                            string queryAlvara = $"UPDATE Geral_Entidade SET CDU_ValidadeAlvara = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryAlvara);
                        }
                        _empresaManager.AnexarDocumento("AlvaraConstrucao", validade);
                        break;
                    case "Certidão permanente":
                        if (validade.HasValue)
                        {
                            string queryCertidaoPermanente = $"UPDATE Geral_Entidade SET CDU_ValidadeCertidaoPermanente = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryCertidaoPermanente);
                        }
                        _empresaManager.AnexarDocumento("CertidaoPermanente", validade);
                        break;
                    case "Contrato de subcontratação/subempreitada/nota de encomenda":
                        if (validade.HasValue)
                        {
                            string queryContrato = $"UPDATE Geral_Entidade SET CDU_ValidadeContratoSubcontratacao = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryContrato);
                        }
                        _empresaManager.AnexarDocumento("ContratoSubcontratacao", validade);
                        break;
                    case "Condições particulares do seguro de acidentes de trabalho":
                        if (validade.HasValue)
                        {
                            string queryCondicoesAT = $"UPDATE Geral_Entidade SET CDU_ValidadeCondicoesAT = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryCondicoesAT);
                        }
                        _empresaManager.AnexarDocumento("CondicoesAT", validade);
                        break;
                    case "Declaração de adesão ao PSS (segue em anexo modelo de declaração a preencher)":
                        if (validade.HasValue)
                        {
                            string queryDeclaracaoPSS = $"UPDATE Geral_Entidade SET CDU_ValidadeDeclaracaoPSS = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryDeclaracaoPSS);
                        }
                        _empresaManager.AnexarDocumento("DeclaracaoPSS", validade);
                        break;
                    case "Declaração do responsável no estaleiro (segue em anexo modelo de declaração a preencher)":
                        if (validade.HasValue)
                        {
                            string queryDeclaracaoResponsavel = $"UPDATE Geral_Entidade SET CDU_ValidadeDeclaracaoResponsavel = '{validade.Value.ToString("yyyy-MM-dd")}' WHERE ID = '{_empresaManager.IdSelecionado}'";
                            _BSO.DSO.ExecuteSQL(queryDeclaracaoResponsavel);
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

                var resultado = _BSO.Consulta(query);
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

            // Criar um painel personalizado moderno para o formulário de autorização
            Panel pnlNovaAutorizacao = groupBoxObras.Controls["pnlNovaAutorizacao"] as Panel;
            if (pnlNovaAutorizacao == null)
            {
                pnlNovaAutorizacao = new Panel
                {
                    Name = "pnlNovaAutorizacao",
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(245, 245, 250),
                    Size = new Size(680, 180),
                    Location = new Point(10, 80)
                };
                groupBoxObras.Controls.Add(pnlNovaAutorizacao);
            }
            pnlNovaAutorizacao.Visible = true;
            pnlNovaAutorizacao.BringToFront();

            // Adicionar um cabeçalho ao painel
            Panel pnlCabecalho = pnlNovaAutorizacao.Controls["pnlCabecalho"] as Panel;
            if (pnlCabecalho == null)
            {
                pnlCabecalho = new Panel
                {
                    Name = "pnlCabecalho",
                    Dock = DockStyle.Top,
                    Height = 32,
                    BackColor = Color.FromArgb(59, 89, 152)
                };
                pnlNovaAutorizacao.Controls.Add(pnlCabecalho);

                Label lblTitulo = new Label
                {
                    Name = "lblTituloNovaEntrada",
                    Text = "AUTORIZAÇÃO DE NOVA ENTRADA EM OBRA",
                    Font = new Font("Calibri", 11F, FontStyle.Bold),
                    ForeColor = Color.White,
                    AutoSize = true,
                    Location = new Point(15, 7)
                };
                pnlCabecalho.Controls.Add(lblTitulo);
            }
            pnlCabecalho.Visible = true;

            // Adicionar controles no painel
            int baseY = 45;
            int spacing = 35;

            // Data de Entrada
            Label lblDataEntrada = pnlNovaAutorizacao.Controls["lblDataEntrada"] as Label;
            DateTimePicker dtpDataEntrada = pnlNovaAutorizacao.Controls["dtpDataEntrada"] as DateTimePicker;

            if (lblDataEntrada == null)
            {
                lblDataEntrada = new Label
                {
                    Name = "lblDataEntrada",
                    Text = "Data de Entrada:",
                    Font = new Font("Calibri", 9.5F),
                    AutoSize = true,
                    Location = new Point(20, baseY)
                };
                pnlNovaAutorizacao.Controls.Add(lblDataEntrada);
            }
            else
            {
                lblDataEntrada.Location = new Point(20, baseY);
                lblDataEntrada.Visible = true;
            }

            if (dtpDataEntrada == null)
            {
                dtpDataEntrada = new DateTimePicker
                {
                    Name = "dtpDataEntrada",
                    Format = DateTimePickerFormat.Short,
                    Font = new Font("Calibri", 9.5F),
                    Size = new Size(120, 24),
                    Location = new Point(120, baseY - 3),
                    Value = DateTime.Today
                };
                pnlNovaAutorizacao.Controls.Add(dtpDataEntrada);
            }
            else
            {
                dtpDataEntrada.Location = new Point(120, baseY - 3);
                dtpDataEntrada.Value = DateTime.Today;
                dtpDataEntrada.Visible = true;
            }

            // Data de Saída
            Label lblDataSaida = pnlNovaAutorizacao.Controls["lblDataSaida"] as Label;
            DateTimePicker dtpDataSaida = pnlNovaAutorizacao.Controls["dtpDataSaida"] as DateTimePicker;

            if (lblDataSaida == null)
            {
                lblDataSaida = new Label
                {
                    Name = "lblDataSaida",
                    Text = "Data de Saída:",
                    Font = new Font("Calibri", 9.5F),
                    AutoSize = true,
                    Location = new Point(270, baseY)
                };
                pnlNovaAutorizacao.Controls.Add(lblDataSaida);
            }
            else
            {
                lblDataSaida.Location = new Point(270, baseY);
                lblDataSaida.Visible = true;
            }

            if (dtpDataSaida == null)
            {
                dtpDataSaida = new DateTimePicker
                {
                    Name = "dtpDataSaida",
                    Format = DateTimePickerFormat.Short,
                    Font = new Font("Calibri", 9.5F),
                    Size = new Size(120, 24),
                    Location = new Point(360, baseY - 3),
                    Value = DateTime.Today.AddMonths(1)
                };
                pnlNovaAutorizacao.Controls.Add(dtpDataSaida);
            }
            else
            {
                dtpDataSaida.Location = new Point(360, baseY - 3);
                dtpDataSaida.Value = DateTime.Today.AddMonths(1);
                dtpDataSaida.Visible = true;
            }

            // Contrato Subempreitada
            Label lblContratoSubempreitada = pnlNovaAutorizacao.Controls["lblContratoSubempreitada"] as Label;
            TextBox txtContratoSubempreitada = pnlNovaAutorizacao.Controls["txtContratoSubempreitada"] as TextBox;

            if (lblContratoSubempreitada == null)
            {
                lblContratoSubempreitada = new Label
                {
                    Name = "lblContratoSubempreitada",
                    Text = "Contrato Subempreitada:",
                    Font = new Font("Calibri", 9.5F),
                    AutoSize = true,
                    Location = new Point(20, baseY + spacing)
                };
                pnlNovaAutorizacao.Controls.Add(lblContratoSubempreitada);
            }
            else
            {
                lblContratoSubempreitada.Location = new Point(20, baseY + spacing);
                lblContratoSubempreitada.Visible = true;
            }

            if (txtContratoSubempreitada == null)
            {
                txtContratoSubempreitada = new TextBox
                {
                    Name = "txtContratoSubempreitada",
                    Size = new Size(280, 24),
                    Font = new Font("Calibri", 9.5F),
                    Location = new Point(170, baseY + spacing - 3),
                    BackColor = Color.LightYellow,
                    BorderStyle = BorderStyle.FixedSingle
                };
                pnlNovaAutorizacao.Controls.Add(txtContratoSubempreitada);
            }
            else
            {
                txtContratoSubempreitada.Location = new Point(170, baseY + spacing - 3);
                txtContratoSubempreitada.Text = "";
                txtContratoSubempreitada.BackColor = Color.LightYellow;
                txtContratoSubempreitada.Visible = true;
            }

            // Status
            Label lblStatusEntrada = pnlNovaAutorizacao.Controls["lblStatusEntrada"] as Label;
            ComboBox cmbStatusEntrada = pnlNovaAutorizacao.Controls["cmbStatusEntrada"] as ComboBox;

            if (lblStatusEntrada == null)
            {
                lblStatusEntrada = new Label
                {
                    Name = "lblStatusEntrada",
                    Text = "Status de Autorização:",
                    Font = new Font("Calibri", 9.5F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(59, 89, 152),
                    AutoSize = true,
                    Location = new Point(20, baseY + spacing * 2)
                };
                pnlNovaAutorizacao.Controls.Add(lblStatusEntrada);
            }
            else
            {
                lblStatusEntrada.Location = new Point(20, baseY + spacing * 2);
                lblStatusEntrada.Visible = true;
            }

            if (cmbStatusEntrada == null)
            {
                cmbStatusEntrada = new ComboBox
                {
                    Name = "cmbStatusEntrada",
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Size = new Size(180, 24),
                    Font = new Font("Calibri", 9.5F),
                    Location = new Point(170, baseY + spacing * 2 - 3)
                };
                cmbStatusEntrada.Items.AddRange(new object[] {
                    "Autorizado",
                    "Pendente",
                    "Não Autorizado",
                    "Renovação Necessária",
                    "Documentos Faltantes"
                });
                cmbStatusEntrada.SelectedIndex = 0;
                pnlNovaAutorizacao.Controls.Add(cmbStatusEntrada);
            }
            else
            {
                cmbStatusEntrada.Location = new Point(170, baseY + spacing * 2 - 3);
                if (cmbStatusEntrada.Items.Count == 0)
                {
                    cmbStatusEntrada.Items.AddRange(new object[] {
                        "Autorizado",
                        "Pendente",
                        "Não Autorizado",
                        "Renovação Necessária",
                        "Documentos Faltantes"
                    });
                }
                cmbStatusEntrada.SelectedIndex = 0;
                cmbStatusEntrada.Visible = true;
            }

            // Colorir o status de acordo com a seleção
            cmbStatusEntrada.BackColor = Color.LightGreen;
            cmbStatusEntrada.SelectedIndexChanged += (s, ev) => {
                ComboBox cmb = s as ComboBox;
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
            };

            // Botões de ação
            Button btnConfirmar = pnlNovaAutorizacao.Controls["btnConfirmar"] as Button;
            Button btnCancelar = pnlNovaAutorizacao.Controls["btnCancelar"] as Button;

            if (btnConfirmar == null)
            {
                btnConfirmar = new Button
                {
                    Name = "btnConfirmar",
                    Text = "Confirmar",
                    BackColor = Color.FromArgb(76, 175, 80),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Calibri", 10F, FontStyle.Bold),
                    Size = new Size(100, 32),
                    Location = new Point(470, baseY + spacing * 2 - 5)
                };
                btnConfirmar.FlatAppearance.BorderColor = Color.FromArgb(56, 142, 60);
                btnConfirmar.Click += new EventHandler(btnConfirmar_Click);
                pnlNovaAutorizacao.Controls.Add(btnConfirmar);
            }
            else
            {
                btnConfirmar.Location = new Point(470, baseY + spacing * 2 - 5);
                btnConfirmar.Visible = true;
            }

            if (btnCancelar == null)
            {
                btnCancelar = new Button
                {
                    Name = "btnCancelar",
                    Text = "Cancelar",
                    BackColor = Color.FromArgb(239, 239, 239),
                    ForeColor = Color.FromArgb(59, 89, 152),
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Calibri", 10F),
                    Size = new Size(90, 32),
                    Location = new Point(580, baseY + spacing * 2 - 5)
                };
                btnCancelar.FlatAppearance.BorderColor = Color.LightGray;
                btnCancelar.Click += (s, ev) => {
                    pnlNovaAutorizacao.Visible = false;
                    dataGridView1.Visible = true;
                    btnGravarObra.Visible = true;
                };
                pnlNovaAutorizacao.Controls.Add(btnCancelar);
            }
            else
            {
                btnCancelar.Location = new Point(580, baseY + spacing * 2 - 5);
                btnCancelar.Visible = true;
            }

            // Adicionar informação de ajuda
            Label lblAjuda = pnlNovaAutorizacao.Controls["lblAjuda"] as Label;
            if (lblAjuda == null)
            {
                lblAjuda = new Label
                {
                    Name = "lblAjuda",
                    Text = "Nota: O status 'Autorizado' permitirá entrada imediata na obra.",
                    Font = new Font("Calibri", 8F, FontStyle.Italic),
                    ForeColor = Color.Gray,
                    AutoSize = true,
                    Location = new Point(20, baseY + spacing * 3 + 5)
                };
                pnlNovaAutorizacao.Controls.Add(lblAjuda);
            }
            else
            {
                lblAjuda.Location = new Point(20, baseY + spacing * 3 + 5);
                lblAjuda.Visible = true;
            }

            // Focar no primeiro campo
            dtpDataEntrada.Focus();
        }

        private void btnConfirmar_Click(object sender, EventArgs e)
        {
            // Obter as referências dos controles do painel
            Panel pnlNovaAutorizacao = groupBoxObras.Controls["pnlNovaAutorizacao"] as Panel;
            if (pnlNovaAutorizacao == null) return;

            DateTimePicker dtpDataEntrada = pnlNovaAutorizacao.Controls["dtpDataEntrada"] as DateTimePicker;
            DateTimePicker dtpDataSaida = pnlNovaAutorizacao.Controls["dtpDataSaida"] as DateTimePicker;
            TextBox txtContratoSubempreitada = pnlNovaAutorizacao.Controls["txtContratoSubempreitada"] as TextBox;
            ComboBox cmbStatusEntrada = pnlNovaAutorizacao.Controls["cmbStatusEntrada"] as ComboBox;

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

            // Verificar se uma obra está selecionada
            if (!(cb_Obras.SelectedItem is KeyValuePair<string, string> obraSelecionada))
            {
                MessageBox.Show("Por favor, selecione uma obra primeiro.",
                    "Nenhuma obra selecionada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string codigoObraSelecionada = obraSelecionada.Key;

            // Obter o status selecionado
            string statusText = cmbStatusEntrada.SelectedItem.ToString();
            int statusIndex = cmbStatusEntrada.SelectedIndex;

            // Criar uma observação detalhada
            string observacao = $"Autorização: {statusText}. Entrada: {dtpDataEntrada.Value:dd/MM/yyyy}, Saída: {dtpDataSaida.Value:dd/MM/yyyy}, Contrato: {txtContratoSubempreitada.Text}";

            try
            {
                // Verificar se a tabela TDU_AD_Obras existe
                string queryCheckTable = @"
                    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'TDU_AD_Obras')
                    BEGIN
                        CREATE TABLE TDU_AD_Obras (
                            CDU_Codigo UNIQUEIDENTIFIER PRIMARY KEY,
                            CDU_Obra NVARCHAR(50) NOT NULL,
                            CDU_EntradaObra NVARCHAR(50) NULL,
                            CDU_SaidaObra NVARCHAR(50) NULL,
                            CDU_ContratoSubempreitada NVARCHAR(100) NULL,
                            CDU_AutorizacaoEntrada BIT DEFAULT 0,
                            CDU_StatusAutorizacao INT DEFAULT 1,
                            CDU_ObservacaoAutorizacao NVARCHAR(500) NULL
                        )
                    END;

                    IF NOT EXISTS (
                        SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = 'TDU_AD_Obras' AND COLUMN_NAME = 'CDU_StatusAutorizacao'
                    )
                    BEGIN
                        ALTER TABLE TDU_AD_Obras ADD CDU_StatusAutorizacao INT DEFAULT 1
                    END;

                    IF NOT EXISTS (
                        SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = 'TDU_AD_Obras' AND COLUMN_NAME = 'CDU_ObservacaoAutorizacao'
                    )
                    BEGIN
                        ALTER TABLE TDU_AD_Obras ADD CDU_ObservacaoAutorizacao NVARCHAR(500) NULL
                    END;
                ";

                _BSO.DSO.ExecuteSQL(queryCheckTable);

                // Criar um novo registro na tabela
                Guid id = Guid.NewGuid();
                string queryInsert = $@"
                    INSERT INTO TDU_AD_Obras 
                    (CDU_Codigo, CDU_Obra, CDU_EntradaObra, CDU_SaidaObra, CDU_ContratoSubempreitada, CDU_AutorizacaoEntrada, CDU_StatusAutorizacao, CDU_ObservacaoAutorizacao) 
                    VALUES 
                    ('{id}', '{codigoObraSelecionada}', '{dtpDataEntrada.Value.ToString("yyyy-MM-dd")}', '{dtpDataSaida.Value.ToString("yyyy-MM-dd")}', 
                    '{txtContratoSubempreitada.Text.Replace("'", "''")}', {(statusIndex == 0 ? 1 : 0)}, {statusIndex}, '{observacao.Replace("'", "''")}');
                ";

                _BSO.DSO.ExecuteSQL(queryInsert);

                // Atualizar o status de autorização para a obra
                _empresaManager.AtualizarStatusAutorizacaoObra(codigoObraSelecionada, statusIndex, observacao);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar autorização: {ex.Message}",
                    "Erro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            // Restaurar interface - mostrar o DataGridView novamente
            pnlNovaAutorizacao.Visible = false;
            dataGridView1.Visible = true;
            btnGravarObra.Visible = true;

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

            // Atualizar a visualização de status
            AtualizarControleAutorizacaoObra(codigoObraSelecionada);

            // Salvar os dados imediatamente
            _empresaManager.SalvarObra();

            // Mensagem de sucesso com um design mais moderno
            using (Form msgForm = new Form())
            {
                msgForm.Size = new Size(400, 150);
                msgForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                msgForm.StartPosition = FormStartPosition.CenterParent;
                msgForm.MaximizeBox = false;
                msgForm.MinimizeBox = false;
                msgForm.Text = "Autorização Registrada";
                msgForm.BackColor = Color.White;

                Panel statusPanel = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 8,
                    BackColor = Color.FromArgb(76, 175, 80) // Verde para sucesso
                };
                msgForm.Controls.Add(statusPanel);

                // Ícone de sucesso (poderia ser substituído por PictureBox com imagem)
                Label iconLabel = new Label
                {
                    Text = "✓",
                    Font = new Font("Calibri", 24F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(76, 175, 80),
                    Size = new Size(50, 50),
                    Location = new Point(20, 30),
                    TextAlign = ContentAlignment.MiddleCenter
                };
                msgForm.Controls.Add(iconLabel);

                // Mensagem
                Label msgLabel = new Label
                {
                    Text = "Autorização de entrada registrada com sucesso!",
                    Font = new Font("Calibri", 10F),
                    Size = new Size(300, 50),
                    Location = new Point(80, 30),
                    TextAlign = ContentAlignment.MiddleLeft
                };
                msgForm.Controls.Add(msgLabel);

                // Botão OK
                Button btnOk = new Button
                {
                    Text = "OK",
                    Size = new Size(80, 30),
                    Location = new Point(300, 80),
                    BackColor = Color.FromArgb(76, 175, 80),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnOk.FlatAppearance.BorderSize = 0;
                btnOk.Click += (s, args) => msgForm.Close();
                msgForm.Controls.Add(btnOk);
                msgForm.AcceptButton = btnOk;

                msgForm.ShowDialog();
            }
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

        // Using the handler already defined in Menu.Designer.cs

        #region Eventos da Aba Trabalhadores

        private void btnAdicionarTrabalhador_Click(object sender, EventArgs e)
        {
            _trabalhadorManager.LimparCampos();
        }

        private void btnEditarTrabalhador_Click(object sender, EventArgs e)
        {
            DataGridView gridTrabalhadores = this.Controls.Find("gridTrabalhadores", true).FirstOrDefault() as DataGridView;

            if (gridTrabalhadores != null && gridTrabalhadores.SelectedRows.Count > 0)
            {
                string idTrabalhador = gridTrabalhadores.SelectedRows[0].Tag?.ToString();
                if (!string.IsNullOrEmpty(idTrabalhador))
                {
                    _trabalhadorManager.CarregarDadosTrabalhador(idTrabalhador);
                }
            }
            else
            {
                MessageBox.Show("Selecione um trabalhador na lista.", "Seleção necessária",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnExcluirTrabalhador_Click(object sender, EventArgs e)
        {
            DataGridView gridTrabalhadores = this.Controls.Find("gridTrabalhadores", true).FirstOrDefault() as DataGridView;

            if (gridTrabalhadores != null && gridTrabalhadores.SelectedRows.Count > 0)
            {
                string idTrabalhador = gridTrabalhadores.SelectedRows[0].Tag?.ToString();
                if (!string.IsNullOrEmpty(idTrabalhador))
                {
                    DialogResult result = MessageBox.Show(
                        "Tem certeza que deseja excluir este trabalhador?",
                        "Confirmar exclusão",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button2);

                    if (result == DialogResult.Yes)
                    {
                        _trabalhadorManager.ExcluirTrabalhador(idTrabalhador);
                    }
                }
            }
            else
            {
                MessageBox.Show("Selecione um trabalhador na lista.", "Seleção necessária",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnSalvarTrabalhador_Click(object sender, EventArgs e)
        {
            _trabalhadorManager.SalvarTrabalhador();
        }

        private void btnAutorizarTrabalhador_Click(object sender, EventArgs e)
        {
            DataGridView gridTrabalhadores = this.Controls.Find("gridTrabalhadores", true).FirstOrDefault() as DataGridView;
            ComboBox cmbObrasTrabalhador = this.Controls.Find("cmbObrasTrabalhador", true).FirstOrDefault() as ComboBox;

            if (gridTrabalhadores != null && gridTrabalhadores.SelectedRows.Count > 0 &&
                cmbObrasTrabalhador != null && cmbObrasTrabalhador.SelectedItem != null)
            {
                string idTrabalhador = gridTrabalhadores.SelectedRows[0].Tag?.ToString();
                string codigoObra = ((KeyValuePair<string, string>)cmbObrasTrabalhador.SelectedItem).Key;

                // Criar formulário de autorização
                using (Form formAutorizacao = new Form())
                {
                    formAutorizacao.Text = "Autorizar Trabalhador";
                    formAutorizacao.Width = 400;
                    formAutorizacao.Height = 250;
                    formAutorizacao.StartPosition = FormStartPosition.CenterParent;
                    formAutorizacao.FormBorderStyle = FormBorderStyle.FixedDialog;
                    formAutorizacao.MaximizeBox = false;
                    formAutorizacao.MinimizeBox = false;

                    // Adicionar controles
                    Label lblStatus = new Label
                    {
                        Text = "Status:",
                        Location = new Point(20, 20),
                        AutoSize = true
                    };
                    formAutorizacao.Controls.Add(lblStatus);

                    ComboBox cmbStatus = new ComboBox
                    {
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Location = new Point(150, 17),
                        Width = 200
                    };
                    cmbStatus.Items.AddRange(new object[] {
                        "Autorizado",
                        "Pendente",
                        "Não Autorizado",
                        "Renovação Necessária",
                        "Documentos Faltantes"
                    });
                    cmbStatus.SelectedIndex = 0;
                    formAutorizacao.Controls.Add(cmbStatus);

                    Label lblObservacoes = new Label
                    {
                        Text = "Observações:",
                        Location = new Point(20, 60),
                        AutoSize = true
                    };
                    formAutorizacao.Controls.Add(lblObservacoes);

                    TextBox txtObservacoes = new TextBox
                    {
                        Multiline = true,
                        Location = new Point(20, 80),
                        Width = 330,
                        Height = 80
                    };
                    formAutorizacao.Controls.Add(txtObservacoes);

                    Button btnConfirmar = new Button
                    {
                        Text = "Confirmar",
                        DialogResult = DialogResult.OK,
                        Location = new Point(120, 170),
                        Width = 80
                    };
                    formAutorizacao.Controls.Add(btnConfirmar);

                    Button btnCancelar = new Button
                    {
                        Text = "Cancelar",
                        DialogResult = DialogResult.Cancel,
                        Location = new Point(220, 170),
                        Width = 80
                    };
                    formAutorizacao.Controls.Add(btnCancelar);

                    // Mostrar formulário
                    if (formAutorizacao.ShowDialog() == DialogResult.OK)
                    {
                        string status = cmbStatus.Text;
                        string observacoes = txtObservacoes.Text;

                        _trabalhadorManager.AutorizarTrabalhador(idTrabalhador, codigoObra, status, observacoes);
                    }
                }
            }
            else
            {
                MessageBox.Show("Selecione um trabalhador na lista e uma obra.", "Seleção necessária",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void gridTrabalhadores_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridView gridTrabalhadores = sender as DataGridView;
                string idTrabalhador = gridTrabalhadores.Rows[e.RowIndex].Tag?.ToString();

                if (!string.IsNullOrEmpty(idTrabalhador))
                {
                    _trabalhadorManager.CarregarDadosTrabalhador(idTrabalhador);
                }
            }
        }

        private void btnAnexarFichaAptidao_Click(object sender, EventArgs e)
        {
            // Mostrar modal de anexos para Ficha de Aptidão Médica
            MostrarModalAnexos("Ficha de Aptidão Médica", "fichaAptidao");
        }

        private void btnAnexarCredenciacao_Click(object sender, EventArgs e)
        {
            // Mostrar modal de anexos para Credenciação
            MostrarModalAnexos("Credenciação", "credenciacao");
        }

        private void btnAnexarFichaEPI_Click(object sender, EventArgs e)
        {
            // Mostrar modal de anexos para Ficha de EPI
            MostrarModalAnexos("Ficha de Distribuição de EPI", "fichaEPI");
        }

        private void MostrarModalAnexos(string tipoDocumento, string tipoAnexo)
        {
            try
            {
                // Criar um form modal para anexos
                using (Form modalAnexos = new Form())
                {
                    modalAnexos.Text = $"Anexar {tipoDocumento}";
                    modalAnexos.Size = new Size(550, 350);
                    modalAnexos.StartPosition = FormStartPosition.CenterParent;
                    modalAnexos.FormBorderStyle = FormBorderStyle.FixedDialog;
                    modalAnexos.MaximizeBox = false;
                    modalAnexos.MinimizeBox = false;
                    modalAnexos.BackColor = Color.White;

                    // Painel de cabeçalho
                    Panel headerPanel = new Panel
                    {
                        Dock = DockStyle.Top,
                        Height = 40,
                        BackColor = Color.FromArgb(59, 89, 152)
                    };
                    modalAnexos.Controls.Add(headerPanel);

                    Label lblTitulo = new Label
                    {
                        Text = $"Anexar {tipoDocumento}",
                        ForeColor = Color.White,
                        Font = new Font("Calibri", 14F, FontStyle.Bold),
                        AutoSize = true,
                        Location = new Point(20, 10)
                    };
                    headerPanel.Controls.Add(lblTitulo);

                    // Painel de conteúdo
                    Panel contentPanel = new Panel
                    {
                        Dock = DockStyle.Fill,
                        BackColor = Color.White,
                        Padding = new Padding(20)
                    };
                    modalAnexos.Controls.Add(contentPanel);

                    // Inputs e controles
                    Label lblDescricao = new Label
                    {
                        Text = "Descrição:",
                        Font = new Font("Calibri", 10F),
                        AutoSize = true,
                        Location = new Point(10, 20)
                    };
                    contentPanel.Controls.Add(lblDescricao);

                    TextBox txtDescricao = new TextBox
                    {
                        Location = new Point(110, 18),
                        Width = 380,
                        Font = new Font("Calibri", 10F)
                    };
                    contentPanel.Controls.Add(txtDescricao);

                    Label lblValidade = new Label
                    {
                        Text = "Validade:",
                        Font = new Font("Calibri", 10F),
                        AutoSize = true,
                        Location = new Point(10, 60)
                    };
                    contentPanel.Controls.Add(lblValidade);

                    DateTimePicker dtpValidade = new DateTimePicker
                    {
                        Location = new Point(110, 58),
                        Width = 180,
                        Format = DateTimePickerFormat.Short,
                        Font = new Font("Calibri", 10F),
                        ShowCheckBox = true,
                        Checked = true,
                        Value = DateTime.Now.AddYears(1)
                    };
                    contentPanel.Controls.Add(dtpValidade);

                    Label lblArquivo = new Label
                    {
                        Text = "Arquivo:",
                        Font = new Font("Calibri", 10F),
                        AutoSize = true,
                        Location = new Point(10, 100)
                    };
                    contentPanel.Controls.Add(lblArquivo);

                    TextBox txtArquivo = new TextBox
                    {
                        Location = new Point(110, 98),
                        Width = 300,
                        Font = new Font("Calibri", 10F),
                        ReadOnly = true,
                        BackColor = Color.WhiteSmoke
                    };
                    contentPanel.Controls.Add(txtArquivo);

                    Button btnSelecionarArquivo = new Button
                    {
                        Text = "...",
                        Location = new Point(420, 97),
                        Width = 70,
                        Font = new Font("Calibri", 10F),
                        UseVisualStyleBackColor = true,
                        FlatStyle = FlatStyle.Flat
                    };
                    btnSelecionarArquivo.FlatAppearance.BorderColor = Color.LightGray;
                    contentPanel.Controls.Add(btnSelecionarArquivo);

                    // Área de observações
                    Label lblObservacoes = new Label
                    {
                        Text = "Observações:",
                        Font = new Font("Calibri", 10F),
                        AutoSize = true,
                        Location = new Point(10, 140)
                    };
                    contentPanel.Controls.Add(lblObservacoes);

                    TextBox txtObservacoes = new TextBox
                    {
                        Location = new Point(110, 138),
                        Width = 380,
                        Height = 80,
                        Font = new Font("Calibri", 10F),
                        Multiline = true,
                        ScrollBars = ScrollBars.Vertical
                    };
                    contentPanel.Controls.Add(txtObservacoes);

                    // Painel de botões
                    Panel buttonPanel = new Panel
                    {
                        Dock = DockStyle.Bottom,
                        Height = 60,
                        BackColor = Color.WhiteSmoke
                    };
                    modalAnexos.Controls.Add(buttonPanel);

                    Button btnSalvar = new Button
                    {
                        Text = "Salvar",
                        DialogResult = DialogResult.OK,
                        Location = new Point(350, 15),
                        Width = 80,
                        Font = new Font("Calibri", 10F, FontStyle.Bold),
                        BackColor = Color.FromArgb(76, 175, 80),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat
                    };
                    btnSalvar.FlatAppearance.BorderColor = Color.FromArgb(56, 142, 60);
                    buttonPanel.Controls.Add(btnSalvar);

                    Button btnCancelar = new Button
                    {
                        Text = "Cancelar",
                        DialogResult = DialogResult.Cancel,
                        Location = new Point(440, 15),
                        Width = 80,
                        Font = new Font("Calibri", 10F),
                        UseVisualStyleBackColor = true,
                        FlatStyle = FlatStyle.Flat
                    };
                    btnCancelar.FlatAppearance.BorderColor = Color.LightGray;
                    buttonPanel.Controls.Add(btnCancelar);

                    // Evento para selecionar arquivo
                    btnSelecionarArquivo.Click += (s, e) =>
                    {
                        using (OpenFileDialog openFileDialog = new OpenFileDialog())
                        {
                            openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png";
                            openFileDialog.FilterIndex = 1;
                            openFileDialog.Multiselect = false;
                            openFileDialog.Title = $"Selecionar {tipoDocumento}";

                            if (openFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                txtArquivo.Text = openFileDialog.FileName;
                            }
                        }
                    };

                    // Exibir modal e processar resultado
                    if (modalAnexos.ShowDialog() == DialogResult.OK)
                    {
                        string caminhoArquivo = txtArquivo.Text;
                        string descricao = txtDescricao.Text;
                        DateTime? validade = dtpValidade.Checked ? dtpValidade.Value : (DateTime?)null;
                        string observacoes = txtObservacoes.Text;

                        if (string.IsNullOrEmpty(caminhoArquivo))
                        {
                            MessageBox.Show("Nenhum arquivo selecionado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        // Chamar o método apropriado no TrabalhadorManager
                        switch (tipoAnexo)
                        {
                            case "fichaAptidao":
                                _trabalhadorManager.AnexarFichaAptidao(caminhoArquivo, descricao, validade, observacoes);
                                break;
                            case "credenciacao":
                                _trabalhadorManager.AnexarCredenciacao(caminhoArquivo, descricao, validade, observacoes);
                                break;
                            case "fichaEPI":
                                _trabalhadorManager.AnexarFichaEPI(caminhoArquivo, descricao, validade, observacoes);
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao anexar documento: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void lblFichaAptidaoAnexo_Click(object sender, EventArgs e)
        {
            // Buscar o nome do arquivo no texto do label
            Label label = sender as Label;
            string texto = label.Text;

            if (texto.Contains(":"))
            {
                string[] partes = texto.Split(':');
                if (partes.Length > 1 && !string.IsNullOrEmpty(partes[1].Trim()))
                {
                    // O caminho do anexo já deve estar na propriedade Tag do label
                    // Implementar o método para visualizar anexo
                    _trabalhadorManager.VisualizarAnexo(_trabalhadorManager.CaminhoAnexoFichaAptidao);
                }
            }
        }

        private void lblCredenciacaoAnexo_Click(object sender, EventArgs e)
        {
            // Buscar o nome do arquivo no texto do label
            Label label = sender as Label;
            string texto = label.Text;

            if (texto.Contains(":"))
            {
                string[] partes = texto.Split(':');
                if (partes.Length > 1 && !string.IsNullOrEmpty(partes[1].Trim()))
                {
                    // O caminho do anexo já deve estar na propriedade Tag do label
                    // Implementar o método para visualizar anexo
                    _trabalhadorManager.VisualizarAnexo(_trabalhadorManager.CaminhoAnexoCredenciacao);
                }
            }
        }

        private void lblFichaEPIAnexo_Click(object sender, EventArgs e)
        {
            // Buscar o nome do arquivo no texto do label
            Label label = sender as Label;
            string texto = label.Text;

            if (texto.Contains(":"))
            {
                string[] partes = texto.Split(':');
                if (partes.Length > 1 && !string.IsNullOrEmpty(partes[1].Trim()))
                {
                    // O caminho do anexo já deve estar na propriedade Tag do label
                    // Implementar o método para visualizar anexo
                    _trabalhadorManager.VisualizarAnexo(_trabalhadorManager.CaminhoAnexoFichaEPI);
                }
            }
        }

        #endregion
        private void btnVerificarDocumentosFaltantes_Click(object sender, EventArgs e)
        {
            _empresaManager.VerificarDocumentosFaltantes();
        }

    }
}