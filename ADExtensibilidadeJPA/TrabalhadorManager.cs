using ErpBS100;
using StdBE100;
using StdPlatBS100;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Linq;

namespace ADExtensibilidadeJPA
{
    public class TrabalhadorManager
    {
        private ErpBS _bso;
        private StdBSInterfPub _pso;
        private string _idEmpresa;
        private string entidadeIDA;
        private string _idTrabalhadorSelecionado;

        // Referências para os controles de interface
        private TextBox _txtNomeTrabalhador;
        private ComboBox _cmbTipoDocumento;
        private TextBox _txtNumDocumento;
        private DateTimePicker _dtpValidadeDocumento;
        private TextBox _txtNIF;
        private TextBox _txtNumSS;
        private CheckBox _chkFichaAptidaoMedica;
        private CheckBox _chkCredenciacao;
        private TextBox _txtCredenciacao;
        private CheckBox _chkFichaEPI;
        private DataGridView _gridTrabalhadores;
        private ComboBox _cmbObrasTrabalhador;

        // Caminhos dos anexos
        private string _caminhoAnexoFichaAptidao = "";
        private string _caminhoAnexoCredenciacao = "";
        private string _caminhoAnexoFichaEPI = "";
        private string _pastaDocumentos = "";

        // Propriedades para acessar caminhos dos anexos
        public string CaminhoAnexoFichaAptidao { get { return _caminhoAnexoFichaAptidao; } }
        public string CaminhoAnexoCredenciacao { get { return _caminhoAnexoCredenciacao; } }
        public string CaminhoAnexoFichaEPI { get { return _caminhoAnexoFichaEPI; } }

        // Labels para mostrar informações dos anexos
        private Label _lblFichaAptidaoAnexo;
        private Label _lblCredenciacaoAnexo;
        private Label _lblFichaEPIAnexo;

        public TrabalhadorManager(ErpBS bso, StdBSInterfPub pso, string idEmpresa, Menu menu)
        {
            try
            {
                if (bso == null)
                    throw new ArgumentNullException("bso", "O objeto de conexão com o ERP não pode ser null");

                if (pso == null)
                    throw new ArgumentNullException("pso", "O objeto de interface não pode ser null");

                if (string.IsNullOrEmpty(idEmpresa))
                    throw new ArgumentException("O ID da empresa não pode ser vazio", "idEmpresa");

                if (menu == null)
                    throw new ArgumentNullException("menu", "O formulário principal não pode ser null");

                _bso = bso;
                _pso = pso;
                _idEmpresa = idEmpresa;
                _idTrabalhadorSelecionado = "";

                // Inicializar referências aos controles da UI
                Control[] controls = menu.Controls.Find("txtNomeTrabalhador", true);
                _txtNomeTrabalhador = (controls != null && controls.Length > 0) ? controls[0] as TextBox : null;

                controls = menu.Controls.Find("cmbTipoDocumentoTrabalhador", true);
                _cmbTipoDocumento = (controls != null && controls.Length > 0) ? controls[0] as ComboBox : null;

                controls = menu.Controls.Find("txtNumDocumento", true);
                _txtNumDocumento = (controls != null && controls.Length > 0) ? controls[0] as TextBox : null;

                controls = menu.Controls.Find("dtpValidadeDocumento", true);
                _dtpValidadeDocumento = (controls != null && controls.Length > 0) ? controls[0] as DateTimePicker : null;

                controls = menu.Controls.Find("txtNIF", true);
                _txtNIF = (controls != null && controls.Length > 0) ? controls[0] as TextBox : null;

                controls = menu.Controls.Find("txtNumSS", true);
                _txtNumSS = (controls != null && controls.Length > 0) ? controls[0] as TextBox : null;

                controls = menu.Controls.Find("chkFichaAptidaoMedica", true);
                _chkFichaAptidaoMedica = (controls != null && controls.Length > 0) ? controls[0] as CheckBox : null;

                controls = menu.Controls.Find("chkCredenciacao", true);
                _chkCredenciacao = (controls != null && controls.Length > 0) ? controls[0] as CheckBox : null;

                controls = menu.Controls.Find("txtCredenciacao", true);
                _txtCredenciacao = (controls != null && controls.Length > 0) ? controls[0] as TextBox : null;

                controls = menu.Controls.Find("chkFichaEPI", true);
                _chkFichaEPI = (controls != null && controls.Length > 0) ? controls[0] as CheckBox : null;

                controls = menu.Controls.Find("gridTrabalhadores", true);
                _gridTrabalhadores = (controls != null && controls.Length > 0) ? controls[0] as DataGridView : null;

                controls = menu.Controls.Find("cmbObrasTrabalhador", true);
                _cmbObrasTrabalhador = (controls != null && controls.Length > 0) ? controls[0] as ComboBox : null;

                // Labels para informações dos anexos
                controls = menu.Controls.Find("lblFichaAptidaoAnexo", true);
                _lblFichaAptidaoAnexo = (controls != null && controls.Length > 0) ? controls[0] as Label : null;

                controls = menu.Controls.Find("lblCredenciacaoAnexo", true);
                _lblCredenciacaoAnexo = (controls != null && controls.Length > 0) ? controls[0] as Label : null;

                controls = menu.Controls.Find("lblFichaEPIAnexo", true);
                _lblFichaEPIAnexo = (controls != null && controls.Length > 0) ? controls[0] as Label : null;

                // Verificar controles críticos
                if (_gridTrabalhadores == null)
                {
                    MessageBox.Show("Erro: Grid de trabalhadores não encontrada. Algumas funcionalidades podem não funcionar corretamente.",
                                    "Erro de inicialização", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // Verificar se a empresa tem uma pasta de documentos definida e usar a mesma para os trabalhadores
                ObterPastaDocumentos();

                // Carregar obras no combobox
                if (_cmbObrasTrabalhador != null)
                {
                    CarregarObrasComboBox();
                }

                // Carregar lista de trabalhadores da empresa
                if (_gridTrabalhadores != null)
                {
                    CarregarTrabalhadores();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao inicializar gerenciador de trabalhadores: {ex.Message}",
                               "Erro de inicialização", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ObterPastaDocumentos()
        {
            try
            {
                string query = $"SELECT CDU_Caminho,EntidadeId FROM Geral_Entidade WHERE ID = '{_idEmpresa}'";

                var resultado = _bso.Consulta(query);
                if (resultado.NumLinhas() > 0)
                {
                    resultado.Inicio();
                    _pastaDocumentos = resultado.DaValor<string>("CDU_Caminho");
                    entidadeIDA = resultado.DaValor<string>("EntidadeId");
                    // Criar subpasta para documentos de trabalhadores caso não exista
                    if (!string.IsNullOrEmpty(_pastaDocumentos))
                    {
                        string pastaTrabalhadores = Path.Combine(_pastaDocumentos, "Trabalhadores");
                        if (!Directory.Exists(pastaTrabalhadores))
                        {
                            Directory.CreateDirectory(pastaTrabalhadores);
                        }
                        _pastaDocumentos = pastaTrabalhadores;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao obter pasta de documentos: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CarregarObrasComboBox()
        {
            try
            {
                // Consultar obras associadas a esta empresa
                string query = $@"SELECT Codigo, Descricao FROM COP_Obras 
                                WHERE Tipo = 'S' AND EntidadeIDA = '{entidadeIDA}'";
                var obras = _bso.Consulta(query);
                _cmbObrasTrabalhador.Items.Clear();

                if (obras != null && obras.NumLinhas() > 0)
                {
                    obras.Inicio();
                    while (!obras.NoFim())
                    {
                        string codigo = obras.DaValor<string>("Codigo");
                        string descricao = obras.DaValor<string>("Descricao");

                        _cmbObrasTrabalhador.Items.Add(new KeyValuePair<string, string>(codigo, $"{codigo} - {descricao}"));

                        obras.Seguinte();
                    }

                    _cmbObrasTrabalhador.DisplayMember = "Value";
                    _cmbObrasTrabalhador.ValueMember = "Key";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao carregar obras: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CarregarTrabalhadores()
        {
            try
            {
                // Verificar se os controles necessários foram inicializados
                if (_gridTrabalhadores == null)
                {
                    MessageBox.Show("Erro: O controle de lista de trabalhadores não foi inicializado corretamente.",
                                    "Erro de inicialização", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (string.IsNullOrEmpty(_idEmpresa))
                {
                    MessageBox.Show("Erro: ID da empresa não foi definido.",
                                    "Erro de inicialização", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Verificar se a tabela existe, se não, criá-la
                string queryCheckTable = @"
                    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                               WHERE TABLE_NAME = 'TDU_AD_Trabalhadores')
                    BEGIN
                        CREATE TABLE [dbo].[TDU_AD_Trabalhadores](
                            [CDU_Id] [uniqueidentifier] NOT NULL,
                            [CDU_IdEmpresa] [uniqueidentifier] NOT NULL,
                            [CDU_Nome] [nvarchar](255) NOT NULL,
                            [CDU_TipoDocumento] [nvarchar](50) NOT NULL,
                            [CDU_NumDocumento] [nvarchar](50) NOT NULL,
                            [CDU_ValidadeDocumento] [date] NULL,
                            [CDU_NIF] [nvarchar](50) NULL,
                            [CDU_NumSS] [nvarchar](50) NULL,
                            [CDU_FichaAptidao] [bit] NOT NULL DEFAULT(0),
                            [CDU_CaminhoFichaAptidao] [nvarchar](500) NULL,
                            [CDU_Credenciacao] [bit] NOT NULL DEFAULT(0),
                            [CDU_DescCredenciacao] [nvarchar](255) NULL,
                            [CDU_CaminhoCredenciacao] [nvarchar](500) NULL,
                            [CDU_FichaEPI] [bit] NOT NULL DEFAULT(0),
                            [CDU_CaminhoFichaEPI] [nvarchar](500) NULL,
                            [CDU_Status] [nvarchar](50) NULL DEFAULT('Pendente'),
                            [CDU_Observacoes] [nvarchar](500) NULL,
                            CONSTRAINT [PK_TDU_AD_Trabalhadores] PRIMARY KEY CLUSTERED ([CDU_Id] ASC)
                        );
                    END
                ";

                if (_bso != null && _bso.DSO != null)
                {
                    _bso.DSO.ExecuteSQL(queryCheckTable);

                    // Consultar trabalhadores associados a esta empresa
                    string query = $@"SELECT * FROM TDU_AD_Trabalhadores WHERE CDU_IdEmpresa = '{_idEmpresa}'";
                    var trabalhadores = _bso.Consulta(query);

                    if (trabalhadores != null)
                    {
                        _gridTrabalhadores.Rows.Clear();

                        if (trabalhadores.NumLinhas() > 0)
                        {
                            trabalhadores.Inicio();
                            while (!trabalhadores.NoFim())
                            {
                                string idTrabalhador = trabalhadores.DaValor<string>("CDU_Id") ?? "";
                                string nome = trabalhadores.DaValor<string>("CDU_Nome") ?? "";
                                string tipoDoc = trabalhadores.DaValor<string>("CDU_TipoDocumento") ?? "";
                                string numDoc = trabalhadores.DaValor<string>("CDU_NumDocumento") ?? "";
                                string nif = trabalhadores.DaValor<string>("CDU_NIF") ?? "";
                                string numSS = trabalhadores.DaValor<string>("CDU_NumSS") ?? "";
                                bool fichaAptidao = trabalhadores.DaValor<bool>("CDU_FichaAptidao");
                                bool credenciacao = trabalhadores.DaValor<bool>("CDU_Credenciacao");
                                bool fichaEPI = trabalhadores.DaValor<bool>("CDU_FichaEPI");
                                string status = trabalhadores.DaValor<string>("CDU_Status") ?? "Pendente";

                                // Combinar tipo e número do documento
                                string documento = $"{tipoDoc}: {numDoc}";

                                // Adicionar linha na grid
                                int rowIndex = _gridTrabalhadores.Rows.Add(nome, documento, nif, numSS, fichaAptidao, credenciacao, fichaEPI, status);

                                // Verificar se a linha foi adicionada com sucesso antes de continuar
                                if (rowIndex >= 0 && rowIndex < _gridTrabalhadores.Rows.Count)
                                {
                                    // Armazenar o ID do trabalhador na tag da linha
                                    _gridTrabalhadores.Rows[rowIndex].Tag = idTrabalhador;

                                    // Colorir a linha conforme o status
                                    if (status == "Autorizado")
                                    {
                                        _gridTrabalhadores.Rows[rowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                                    }
                                    else if (status == "Não Autorizado")
                                    {
                                        _gridTrabalhadores.Rows[rowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightCoral;
                                    }
                                    else if (status == "Renovação Necessária")
                                    {
                                        _gridTrabalhadores.Rows[rowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightSalmon;
                                    }
                                    else if (status == "Documentos Faltantes")
                                    {
                                        _gridTrabalhadores.Rows[rowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightPink;
                                    }
                                }

                                trabalhadores.Seguinte();
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Erro: Falha ao consultar trabalhadores no banco de dados.",
                                        "Erro de consulta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Erro: Conexão com o banco de dados não foi inicializada corretamente.",
                                    "Erro de conexão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao carregar trabalhadores: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CarregarDadosTrabalhador(string idTrabalhador)
        {
            try
            {
                LimparCampos();

                string query = $@"SELECT * FROM TDU_AD_Trabalhadores WHERE CDU_Id = '{idTrabalhador}'";
                var trabalhador = _bso.Consulta(query);

                if (trabalhador != null && trabalhador.NumLinhas() > 0)
                {
                    trabalhador.Inicio();

                    _idTrabalhadorSelecionado = idTrabalhador;
                    _txtNomeTrabalhador.Text = trabalhador.DaValor<string>("CDU_Nome");
                    _cmbTipoDocumento.Text = trabalhador.DaValor<string>("CDU_TipoDocumento");
                    _txtNumDocumento.Text = trabalhador.DaValor<string>("CDU_NumDocumento");

                    // Validar data e tratar se for nula
                    string validadeStr = trabalhador.DaValor<string>("CDU_ValidadeDocumento");
                    if (!string.IsNullOrEmpty(validadeStr) && DateTime.TryParse(validadeStr, out DateTime validadeDoc))
                    {
                        _dtpValidadeDocumento.Value = validadeDoc;
                        _dtpValidadeDocumento.Checked = true;
                    }
                    else
                    {
                        _dtpValidadeDocumento.Checked = false;
                    }

                    _txtNIF.Text = trabalhador.DaValor<string>("CDU_NIF");
                    _txtNumSS.Text = trabalhador.DaValor<string>("CDU_NumSS");
                    _chkFichaAptidaoMedica.Checked = trabalhador.DaValor<bool>("CDU_FichaAptidao");
                    _chkCredenciacao.Checked = trabalhador.DaValor<bool>("CDU_Credenciacao");
                    _txtCredenciacao.Text = trabalhador.DaValor<string>("CDU_DescCredenciacao");
                    _txtCredenciacao.Enabled = _chkCredenciacao.Checked;
                    _chkFichaEPI.Checked = trabalhador.DaValor<bool>("CDU_FichaEPI");

                    // Caminhos dos anexos
                    _caminhoAnexoFichaAptidao = trabalhador.DaValor<string>("CDU_CaminhoFichaAptidao") ?? "";
                    _caminhoAnexoCredenciacao = trabalhador.DaValor<string>("CDU_CaminhoCredenciacao") ?? "";
                    _caminhoAnexoFichaEPI = trabalhador.DaValor<string>("CDU_CaminhoFichaEPI") ?? "";

                    // Atualizar labels dos anexos
                    AtualizarLabelsAnexos();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao carregar dados do trabalhador: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AtualizarLabelsAnexos()
        {
            _lblFichaAptidaoAnexo.Text = string.IsNullOrEmpty(_caminhoAnexoFichaAptidao) ?
                "Ficha de Aptidão Médica:" : "Ficha de Aptidão Médica: " + Path.GetFileName(_caminhoAnexoFichaAptidao);

            _lblCredenciacaoAnexo.Text = string.IsNullOrEmpty(_caminhoAnexoCredenciacao) ?
                "Credenciação:" : "Credenciação: " + Path.GetFileName(_caminhoAnexoCredenciacao);

            _lblFichaEPIAnexo.Text = string.IsNullOrEmpty(_caminhoAnexoFichaEPI) ?
                "Ficha de Distribuição EPI:" : "Ficha de Distribuição EPI: " + Path.GetFileName(_caminhoAnexoFichaEPI);
        }

        public void LimparCampos()
        {
            _idTrabalhadorSelecionado = "";
            _txtNomeTrabalhador.Text = "";
            _cmbTipoDocumento.SelectedIndex = -1;
            _txtNumDocumento.Text = "";
            _dtpValidadeDocumento.Value = DateTime.Today;
            _dtpValidadeDocumento.Checked = false;
            _txtNIF.Text = "";
            _txtNumSS.Text = "";
            _chkFichaAptidaoMedica.Checked = false;
            _chkCredenciacao.Checked = false;
            _txtCredenciacao.Text = "";
            _txtCredenciacao.Enabled = false;
            _chkFichaEPI.Checked = false;

            _caminhoAnexoFichaAptidao = "";
            _caminhoAnexoCredenciacao = "";
            _caminhoAnexoFichaEPI = "";

            AtualizarLabelsAnexos();
        }

        public bool SalvarTrabalhador()
        {
            try
            {
                // Validar dados obrigatórios
                if (string.IsNullOrWhiteSpace(_txtNomeTrabalhador.Text))
                {
                    MessageBox.Show("O nome do trabalhador é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (_cmbTipoDocumento.SelectedIndex == -1)
                {
                    MessageBox.Show("Selecione o tipo de documento.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (string.IsNullOrWhiteSpace(_txtNumDocumento.Text))
                {
                    MessageBox.Show("O número do documento é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Validar credenciação se marcada
                if (_chkCredenciacao.Checked && string.IsNullOrWhiteSpace(_txtCredenciacao.Text))
                {
                    MessageBox.Show("Especifique o tipo de credenciação.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Verificar se tabela existe, se não, criá-la
                string queryCheckTable = @"
                    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                               WHERE TABLE_NAME = 'TDU_AD_Trabalhadores')
                    BEGIN
                        CREATE TABLE [dbo].[TDU_AD_Trabalhadores](
                            [CDU_Id] [uniqueidentifier] NOT NULL,
                            [CDU_IdEmpresa] [uniqueidentifier] NOT NULL,
                            [CDU_Nome] [nvarchar](255) NOT NULL,
                            [CDU_TipoDocumento] [nvarchar](50) NOT NULL,
                            [CDU_NumDocumento] [nvarchar](50) NOT NULL,
                            [CDU_ValidadeDocumento] [date] NULL,
                            [CDU_NIF] [nvarchar](50) NULL,
                            [CDU_NumSS] [nvarchar](50) NULL,
                            [CDU_FichaAptidao] [bit] NOT NULL DEFAULT(0),
                            [CDU_CaminhoFichaAptidao] [nvarchar](500) NULL,
                            [CDU_Credenciacao] [bit] NOT NULL DEFAULT(0),
                            [CDU_DescCredenciacao] [nvarchar](255) NULL,
                            [CDU_CaminhoCredenciacao] [nvarchar](500) NULL,
                            [CDU_FichaEPI] [bit] NOT NULL DEFAULT(0),
                            [CDU_CaminhoFichaEPI] [nvarchar](500) NULL,
                            [CDU_Status] [nvarchar](50) NULL DEFAULT('Pendente'),
                            [CDU_Observacoes] [nvarchar](500) NULL,
                            CONSTRAINT [PK_TDU_AD_Trabalhadores] PRIMARY KEY CLUSTERED ([CDU_Id] ASC)
                        );
                    END
                ";
                _bso.DSO.ExecuteSQL(queryCheckTable);

                // Preparar dados para salvar
                string id = string.IsNullOrEmpty(_idTrabalhadorSelecionado) ? Guid.NewGuid().ToString() : _idTrabalhadorSelecionado;
                string nome = _txtNomeTrabalhador.Text.Replace("'", "''");
                string tipoDoc = _cmbTipoDocumento.Text.Replace("'", "''");
                string numDoc = _txtNumDocumento.Text.Replace("'", "''");
                string validadeDoc = _dtpValidadeDocumento.Checked ? _dtpValidadeDocumento.Value.ToString("yyyy-MM-dd") : "NULL";
                string nif = _txtNIF.Text.Replace("'", "''");
                string numSS = _txtNumSS.Text.Replace("'", "''");
                bool fichaAptidao = _chkFichaAptidaoMedica.Checked;
                bool credenciacao = _chkCredenciacao.Checked;
                string descCredenciacao = _txtCredenciacao.Text.Replace("'", "''");
                bool fichaEPI = _chkFichaEPI.Checked;

                // Verificar se é novo registro ou atualização
                if (string.IsNullOrEmpty(_idTrabalhadorSelecionado))
                {
                    // Inserir novo trabalhador - Corrigindo inserção
                    string queryInsert = $@"
                        INSERT INTO TDU_AD_Trabalhadores (
                            CDU_Id, CDU_IdEmpresa, CDU_Nome, CDU_TipoDocumento, CDU_NumDocumento, 
                            CDU_ValidadeDocumento, CDU_NIF, CDU_NumSS, CDU_FichaAptidao, 
                            CDU_CaminhoFichaAptidao, CDU_Credenciacao, CDU_DescCredenciacao, 
                            CDU_CaminhoCredenciacao, CDU_FichaEPI, CDU_CaminhoFichaEPI, CDU_Status
                        ) VALUES (
                            '{id}', '{_idEmpresa}', '{nome}', '{tipoDoc}', '{numDoc}', 
                            {(validadeDoc == "NULL" ? validadeDoc : $"'{validadeDoc}'")}, '{nif}', '{numSS}', {(fichaAptidao ? 1 : 0)}, 
                            '{_caminhoAnexoFichaAptidao?.Replace("'", "''")}', {(credenciacao ? 1 : 0)}, '{descCredenciacao}', 
                            '{_caminhoAnexoCredenciacao?.Replace("'", "''")}', {(fichaEPI ? 1 : 0)}, '{_caminhoAnexoFichaEPI?.Replace("'", "''")}', 'Pendente'
                        )";

                    _bso.DSO.ExecuteSQL(queryInsert);
                    _idTrabalhadorSelecionado = id; // Atualizar ID do trabalhador selecionado
                }
                else
                {
                    // Atualizar trabalhador existente
                    string queryUpdate = $@"
                        UPDATE TDU_AD_Trabalhadores SET
                            CDU_Nome = '{nome}', 
                            CDU_TipoDocumento = '{tipoDoc}', 
                            CDU_NumDocumento = '{numDoc}', 
                            CDU_ValidadeDocumento = {(validadeDoc == "NULL" ? validadeDoc : $"'{validadeDoc}'")}, 
                            CDU_NIF = '{nif}', 
                            CDU_NumSS = '{numSS}', 
                            CDU_FichaAptidao = {(fichaAptidao ? 1 : 0)}, 
                            CDU_CaminhoFichaAptidao = '{_caminhoAnexoFichaAptidao?.Replace("'", "''")}', 
                            CDU_Credenciacao = {(credenciacao ? 1 : 0)}, 
                            CDU_DescCredenciacao = '{descCredenciacao}', 
                            CDU_CaminhoCredenciacao = '{_caminhoAnexoCredenciacao?.Replace("'", "''")}', 
                            CDU_FichaEPI = {(fichaEPI ? 1 : 0)}, 
                            CDU_CaminhoFichaEPI = '{_caminhoAnexoFichaEPI?.Replace("'", "''")}'
                        WHERE CDU_Id = '{id}'";
                    _bso.DSO.ExecuteSQL(queryUpdate);
                }

                MessageBox.Show("Trabalhador salvo com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Recarregar lista de trabalhadores
                CarregarTrabalhadores();

                // Limpar formulário
                LimparCampos();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar trabalhador: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public bool VerificarAutorizacaoTrabalhador(string idTrabalhador, string idObra)
        {
            try
            {
                // Verificar se tabela TDU_AD_AutorizacaoTrabalhador existe, se não, criá-la
                string queryCheckTable = @"
                    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                               WHERE TABLE_NAME = 'TDU_AD_AutorizacaoTrabalhador')
                    BEGIN
                        CREATE TABLE [dbo].[TDU_AD_AutorizacaoTrabalhador](
                            [CDU_Id] [uniqueidentifier] NOT NULL,
                            [CDU_IdTrabalhador] [uniqueidentifier] NOT NULL,
                            [CDU_IdEmpresa] [uniqueidentifier] NOT NULL,
                            [CDU_IdObra] [nvarchar](50) NOT NULL,
                            [CDU_Status] [nvarchar](50) NOT NULL DEFAULT('Pendente'),
                            [CDU_DataAutorizacao] [datetime] NULL,
                            [CDU_Observacoes] [nvarchar](500) NULL,
                            CONSTRAINT [PK_TDU_AD_AutorizacaoTrabalhador] PRIMARY KEY CLUSTERED ([CDU_Id] ASC)
                        );
                    END
                ";
                _bso.DSO.ExecuteSQL(queryCheckTable);

                // Verificar se existe autorização
                string queryCheck = $@"
                    SELECT CDU_Status FROM TDU_AD_AutorizacaoTrabalhador 
                    WHERE CDU_IdTrabalhador = '{idTrabalhador}' 
                    AND CDU_IdObra = '{idObra}'";
                var resultado = _bso.Consulta(queryCheck);

                if (resultado != null && resultado.NumLinhas() > 0)
                {
                    resultado.Inicio();
                    string status = resultado.DaValor<string>("CDU_Status");
                    return status == "Autorizado";
                }

                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao verificar autorização: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public void AutorizarTrabalhador(string idTrabalhador, string idObra, string status, string observacoes)
        {
            try
            {
                // Verificar se trabalhador existe
                string query = $@"SELECT COUNT(*) AS Total FROM TDU_AD_Trabalhadores WHERE CDU_Id = '{idTrabalhador}'";
                var resultadoTrabalhador = _bso.Consulta(query);
                if (resultadoTrabalhador != null && resultadoTrabalhador.NumLinhas() > 0)
                {
                    resultadoTrabalhador.Inicio();
                    int totalTrabalhador = resultadoTrabalhador.DaValor<int>("Total");

                    if (totalTrabalhador == 0)
                    {
                        MessageBox.Show("Trabalhador não encontrado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Verificar se já existe autorização
                    string queryCheckAutorizacao = $@"
                        SELECT CDU_Id FROM TDU_AD_AutorizacaoTrabalhador 
                        WHERE CDU_IdTrabalhador = '{idTrabalhador}' 
                        AND CDU_IdObra = '{idObra}'";
                    var resultadoAutorizacao = _bso.Consulta(queryCheckAutorizacao);

                    string idAutorizacao;
                    if (resultadoAutorizacao != null && resultadoAutorizacao.NumLinhas() > 0)
                    {
                        // Atualizar autorização existente
                        resultadoAutorizacao.Inicio();
                        idAutorizacao = resultadoAutorizacao.DaValor<string>("CDU_Id");

                        string queryUpdate = $@"
                            UPDATE TDU_AD_AutorizacaoTrabalhador SET
                                CDU_Status = '{status}',
                                CDU_DataAutorizacao = '{DateTime.Now:yyyy-MM-dd HH:mm:ss}',
                                CDU_Observacoes = '{observacoes}'
                            WHERE CDU_Id = '{idAutorizacao}'";
                        _bso.DSO.ExecuteSQL(queryUpdate);
                    }
                    else
                    {
                        // Criar nova autorização
                        idAutorizacao = Guid.NewGuid().ToString();

                        string queryInsert = $@"
                            INSERT INTO TDU_AD_AutorizacaoTrabalhador (
                                CDU_Id, CDU_IdTrabalhador, CDU_IdEmpresa, CDU_IdObra, 
                                CDU_Status, CDU_DataAutorizacao, CDU_Observacoes
                            ) VALUES (
                                '{idAutorizacao}', '{idTrabalhador}', '{_idEmpresa}', '{idObra}', 
                                '{status}', '{DateTime.Now:yyyy-MM-dd HH:mm:ss}', '{observacoes}'
                            )";
                        _bso.DSO.ExecuteSQL(queryInsert);
                    }

                    // Atualizar o status do trabalhador na tabela principal
                    string queryUpdateTrabalhador = $@"
                        UPDATE TDU_AD_Trabalhadores SET
                            CDU_Status = '{status}'
                        WHERE CDU_Id = '{idTrabalhador}'";
                    _bso.DSO.ExecuteSQL(queryUpdateTrabalhador);

                    MessageBox.Show($"Trabalhador {status} com sucesso!", "Autorização", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Recarregar lista de trabalhadores
                    CarregarTrabalhadores();
                }
                else
                {
                    MessageBox.Show("Erro ao consultar o trabalhador.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao autorizar trabalhador: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool ExcluirTrabalhador(string idTrabalhador)
        {
            try
            {
                // Verificar se existe autorizações para este trabalhador
                string queryCheckAutorizacoes = $@"
                    SELECT COUNT(*) AS Total FROM TDU_AD_AutorizacaoTrabalhador 
                    WHERE CDU_IdTrabalhador = '{idTrabalhador}'";
                var resultadoAutorizacoes = _bso.Consulta(queryCheckAutorizacoes);
                if (resultadoAutorizacoes != null && resultadoAutorizacoes.NumLinhas() > 0)
                {
                    resultadoAutorizacoes.Inicio();
                    int totalAutorizacoes = resultadoAutorizacoes.DaValor<int>("Total");

                    // Excluir autorizações
                    if (totalAutorizacoes > 0)
                    {
                        string queryDeleteAutorizacoes = $@"
                            DELETE FROM TDU_AD_AutorizacaoTrabalhador 
                            WHERE CDU_IdTrabalhador = '{idTrabalhador}'";
                        _bso.DSO.ExecuteSQL(queryDeleteAutorizacoes);
                    }
                }

                // Excluir trabalhador
                string queryDeleteTrabalhador = $@"
                    DELETE FROM TDU_AD_Trabalhadores 
                    WHERE CDU_Id = '{idTrabalhador}'";
                _bso.DSO.ExecuteSQL(queryDeleteTrabalhador);

                MessageBox.Show("Trabalhador excluído com sucesso!", "Exclusão", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Recarregar lista de trabalhadores
                CarregarTrabalhadores();

                // Limpar formulário
                LimparCampos();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao excluir trabalhador: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        #region Anexos

        public void AnexarFichaAptidao(string caminhoArquivo, string descricao, DateTime? validade, string observacoes)
        {
            AnexarDocumentoAvancado("Ficha de Aptidão Médica", ref _caminhoAnexoFichaAptidao, caminhoArquivo, descricao, validade, observacoes);
        }

        public void AnexarCredenciacao(string caminhoArquivo, string descricao, DateTime? validade, string observacoes)
        {
            AnexarDocumentoAvancado("Credenciação", ref _caminhoAnexoCredenciacao, caminhoArquivo, descricao, validade, observacoes);
        }

        public void AnexarFichaEPI(string caminhoArquivo, string descricao, DateTime? validade, string observacoes)
        {
            AnexarDocumentoAvancado("Ficha de Distribuição de EPI", ref _caminhoAnexoFichaEPI, caminhoArquivo, descricao, validade, observacoes);
        }

        // Manter método antigo para retrocompatibilidade
        public void AnexarFichaAptidao()
        {
            AnexarDocumento("Ficha de Aptidão Médica", ref _caminhoAnexoFichaAptidao);
        }

        public void AnexarCredenciacao()
        {
            AnexarDocumento("Credenciação", ref _caminhoAnexoCredenciacao);
        }

        public void AnexarFichaEPI()
        {
            AnexarDocumento("Ficha de Distribuição de EPI", ref _caminhoAnexoFichaEPI);
        }

        private void AnexarDocumento(string tipoDocumento, ref string caminhoPasta)
        {
            // Verificar se a pasta de documentos foi definida
            if (string.IsNullOrEmpty(_pastaDocumentos))
            {
                MessageBox.Show("A pasta de documentos não foi definida. Configure a pasta na aba Empresa.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;
                openFileDialog.Title = $"Selecionar {tipoDocumento}";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string nomeArquivo = string.IsNullOrEmpty(_txtNomeTrabalhador.Text)
                            ? "Sem_Nome"
                            : _txtNomeTrabalhador.Text.Replace(" ", "_");

                        string fileName = $"{tipoDocumento.Replace(" ", "_")}_{nomeArquivo}_{DateTime.Now.ToString("yyyyMMdd")}{Path.GetExtension(sourceFile)}";
                        string destFile = Path.Combine(_pastaDocumentos, fileName);

                        // Verificar se o arquivo já existe
                        if (File.Exists(destFile))
                        {
                            DialogResult result = MessageBox.Show(
                                $"O arquivo {fileName} já existe na pasta de destino. Deseja substituí-lo?",
                                "Arquivo já existe",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (result == DialogResult.No)
                                return;
                        }

                        // Copia o arquivo para a pasta de destino
                        File.Copy(sourceFile, destFile, true);

                        // Atualiza o caminho do anexo
                        caminhoPasta = destFile;

                        // Atualiza os labels
                        AtualizarLabelsAnexos();

                        MessageBox.Show("Documento anexado com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void AnexarDocumentoAvancado(string tipoDocumento, ref string caminhoPasta, string arquivoOrigem, string descricao, DateTime? validade, string observacoes)
        {
            // Verificar se a pasta de documentos foi definida
            if (string.IsNullOrEmpty(_pastaDocumentos))
            {
                MessageBox.Show("A pasta de documentos não foi definida. Configure a pasta na aba Empresa.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string sourceFile = arquivoOrigem;
                string nomeArquivo = string.IsNullOrEmpty(_txtNomeTrabalhador.Text)
                    ? "Sem_Nome"
                    : _txtNomeTrabalhador.Text.Replace(" ", "_");

                string fileName = $"{tipoDocumento.Replace(" ", "_")}_{nomeArquivo}_{DateTime.Now.ToString("yyyyMMdd")}{Path.GetExtension(sourceFile)}";
                string destFile = Path.Combine(_pastaDocumentos, fileName);

                // Verificar se o arquivo já existe
                if (File.Exists(destFile))
                {
                    DialogResult result = MessageBox.Show(
                        $"O arquivo {fileName} já existe na pasta de destino. Deseja substituí-lo?",
                        "Arquivo já existe",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                        return;
                }

                // Copia o arquivo para a pasta de destino
                File.Copy(sourceFile, destFile, true);

                // Atualiza o caminho do anexo
                caminhoPasta = destFile;

                // Atualiza os labels
                AtualizarLabelsAnexos();

                // Salvar informações adicionais na tabela TDU_AD_Trabalhadores_Anexos
                SalvarInformacoesAnexo(tipoDocumento, caminhoPasta, descricao, validade, observacoes);

                MessageBox.Show("Documento anexado com sucesso!",
                    "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SalvarInformacoesAnexo(string tipoDocumento, string caminhoPasta, string descricao, DateTime? validade, string observacoes)
        {
            // Verificar se tabela existe, se não, criá-la
            string queryCheckTable = @"
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                           WHERE TABLE_NAME = 'TDU_AD_Trabalhadores_Anexos')
                BEGIN
                    CREATE TABLE [dbo].[TDU_AD_Trabalhadores_Anexos](
                        [CDU_Id] [uniqueidentifier] NOT NULL,
                        [CDU_IdTrabalhador] [uniqueidentifier] NOT NULL,
                        [CDU_TipoDocumento] [nvarchar](100) NOT NULL,
                        [CDU_Descricao] [nvarchar](255) NULL,
                        [CDU_CaminhoAnexo] [nvarchar](500) NOT NULL,
                        [CDU_DataInclusao] [datetime] NOT NULL,
                        [CDU_Validade] [date] NULL,
                        [CDU_Observacoes] [nvarchar](500) NULL,
                        CONSTRAINT [PK_TDU_AD_Trabalhadores_Anexos] PRIMARY KEY CLUSTERED ([CDU_Id] ASC)
                    );
                END
            ";
            _bso.DSO.ExecuteSQL(queryCheckTable);

            try
            {
                // Verificar se já existe um anexo deste tipo para este trabalhador
                string queryCheck = $@"
                    SELECT CDU_Id FROM TDU_AD_Trabalhadores_Anexos 
                    WHERE CDU_IdTrabalhador = '{_idTrabalhadorSelecionado}' 
                    AND CDU_TipoDocumento = '{tipoDocumento}'";

                var resultado = _bso.Consulta(queryCheck);
                string idAnexo;

                if (resultado != null && resultado.NumLinhas() > 0)
                {
                    // Atualizar registro existente
                    resultado.Inicio();
                    idAnexo = resultado.DaValor<string>("CDU_Id");

                    string queryUpdate = $@"
                        UPDATE TDU_AD_Trabalhadores_Anexos SET
                            CDU_Descricao = '{descricao?.Replace("'", "''")}',
                            CDU_CaminhoAnexo = '{caminhoPasta?.Replace("'", "''")}',
                            CDU_DataInclusao = '{DateTime.Now:yyyy-MM-dd HH:mm:ss}',
                            CDU_Validade = {(validade.HasValue ? $"'{validade.Value:yyyy-MM-dd}'" : "NULL")},
                            CDU_Observacoes = '{observacoes?.Replace("'", "''")}'
                        WHERE CDU_Id = '{idAnexo}'";

                    _bso.DSO.ExecuteSQL(queryUpdate);
                }
                else if (!string.IsNullOrEmpty(_idTrabalhadorSelecionado))
                {
                    // Inserir novo registro
                    idAnexo = Guid.NewGuid().ToString();

                    string queryInsert = $@"
                        INSERT INTO TDU_AD_Trabalhadores_Anexos (
                            CDU_Id, CDU_IdTrabalhador, CDU_TipoDocumento, CDU_Descricao, 
                            CDU_CaminhoAnexo, CDU_DataInclusao, CDU_Validade, CDU_Observacoes
                        ) VALUES (
                            '{idAnexo}', '{_idTrabalhadorSelecionado}', '{tipoDocumento}', '{descricao?.Replace("'", "''")}',
                            '{caminhoPasta?.Replace("'", "''")}', '{DateTime.Now:yyyy-MM-dd HH:mm:ss}', 
                            {(validade.HasValue ? $"'{validade.Value:yyyy-MM-dd}'" : "NULL")}, '{observacoes?.Replace("'", "''")}'
                        )";

                    _bso.DSO.ExecuteSQL(queryInsert);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar informações do anexo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void VisualizarAnexo(string caminhoAnexo)
        {
            if (string.IsNullOrEmpty(caminhoAnexo) || !File.Exists(caminhoAnexo))
            {
                MessageBox.Show("Anexo não encontrado.", "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(caminhoAnexo);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}