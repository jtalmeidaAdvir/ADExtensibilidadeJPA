
using ErpBS100;
using StdBE100;
using StdPlatBS100;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public class EmpresaManager
    {
        private ErpBS _bso;
        private StdBSInterfPub _pso;
        private string _id;
        private string _idSelecionado;

        // Controles da UI
        private TextBox _txtCodigo;
        private TextBox _txtNome;
        private TextBox _txtSede;
        private TextBox _txtContribuinte;
        private TextBox _txtAlvara;
        private DateTimePicker _txtAlvaraValidade;
        private DateTimePicker _txtNaoDivFinancas;
        private DateTimePicker _txtNaoDivSegSocial;
        private DateTimePicker _txtFolhaPagSegSocial;
        private TextBox _txtReciboApoliceAT;
        private TextBox _txtReciboRC;
        private TextBox _txtCaminhoPasta;
        private ComboBox _cbReciboPagSegSocial;
        private ComboBox _cbApoliceAT;
        private ComboBox _cbApoliceRC;
        private ComboBox _cbHorarioTrabalho;
        private ComboBox _cbDecTrabIlegais;
        private ComboBox _cbDecRespEstaleiro;
        private ComboBox _cbDecConhecimPSS;
        private Label _lblAnexoFinancas;
        private Label _lblAnexoSegSocial;
        private Label _lblFolhaPagSS;
        private ComboBox _cbObras;
        private DataGridView _dataGridView;

        // Variáveis para armazenar os caminhos dos anexos específicos
        private string _caminhoAnexoFinancas = "";
        private string _caminhoAnexoSegSocial = "";
        private string _caminhoAnexoFolhaPag = "";

        public EmpresaManager(ErpBS bso, StdBSInterfPub pso, string idSelecionado, Menu menu)
        {
            _bso = bso;
            _pso = pso;
            _idSelecionado = idSelecionado;

            // Use Form.Controls.Find para encontrar os controles por nome
            // Inicializar referências aos controles da UI
            _txtCodigo = menu.Controls.Find("TXT_Codigo", true).FirstOrDefault() as TextBox;
            _txtNome = menu.Controls.Find("TXT_Nome", true).FirstOrDefault() as TextBox;
            _txtSede = menu.Controls.Find("TXT_Sede", true).FirstOrDefault() as TextBox;
            _txtContribuinte = menu.Controls.Find("TXT_Contribuinte", true).FirstOrDefault() as TextBox;
            _txtAlvara = menu.Controls.Find("TXT_Alvara", true).FirstOrDefault() as TextBox;
            _txtAlvaraValidade = menu.Controls.Find("TXT_AlvaraValidade", true).FirstOrDefault() as DateTimePicker;
            _txtNaoDivFinancas = menu.Controls.Find("TXT_NaoDivFinancas", true).FirstOrDefault() as DateTimePicker;
            _txtNaoDivSegSocial = menu.Controls.Find("TXT_NaoDivSegSocial", true).FirstOrDefault() as DateTimePicker;
            _txtFolhaPagSegSocial = menu.Controls.Find("TXT_FolhaPagSegSocial", true).FirstOrDefault() as DateTimePicker;
            _txtReciboApoliceAT = menu.Controls.Find("TXT_ReciboApoliceAT", true).FirstOrDefault() as TextBox;
            _txtReciboRC = menu.Controls.Find("TXT_ReciboRC", true).FirstOrDefault() as TextBox;
            _txtCaminhoPasta = menu.Controls.Find("txtCaminhoPasta", true).FirstOrDefault() as TextBox;
            _cbReciboPagSegSocial = menu.Controls.Find("cb_ReciboPagSegSocial", true).FirstOrDefault() as ComboBox;
            _cbApoliceAT = menu.Controls.Find("cb_ApoliceAT", true).FirstOrDefault() as ComboBox;
            _cbApoliceRC = menu.Controls.Find("cb_ApoliceRC", true).FirstOrDefault() as ComboBox;
            _cbHorarioTrabalho = menu.Controls.Find("cb_HorarioTrabalho", true).FirstOrDefault() as ComboBox;
            _cbDecTrabIlegais = menu.Controls.Find("cb_DecTrabIlegais", true).FirstOrDefault() as ComboBox;
            _cbDecRespEstaleiro = menu.Controls.Find("cb_DecRespEstaleiro", true).FirstOrDefault() as ComboBox;
            _cbDecConhecimPSS = menu.Controls.Find("cb_DecConhecimPSS", true).FirstOrDefault() as ComboBox;
            _lblAnexoFinancas = menu.Controls.Find("lblAnexoFinancas", true).FirstOrDefault() as Label;
            _lblAnexoSegSocial = menu.Controls.Find("lblAnexoSegSocial", true).FirstOrDefault() as Label;
            _lblFolhaPagSS = menu.Controls.Find("lblFolhaPagSS", true).FirstOrDefault() as Label;
            _cbObras = menu.Controls.Find("cb_Obras", true).FirstOrDefault() as ComboBox;
            _dataGridView = menu.Controls.Find("dataGridView1", true).FirstOrDefault() as DataGridView;

            if (_idSelecionado != "")
            {
                CarregarDados();
            }
        }

        public void CarregarDados()
        {
            Dictionary<string, string> entidade = new Dictionary<string, string>();
            GetEntidadesID(ref entidade);
            if (entidade.Count > 0)
            {
                SetInfoEntidades(entidade);
            }
        }

        // Keep this for backward compatibility
        public void DaValores()
        {
            CarregarDados();
        }

        private void GetEntidadesID(ref Dictionary<string, string> entidade)
        {
            // Consulta SQL para pegar os dados
            var query = $@"SELECT * FROM Geral_Entidade WHERE CDU_TrataSGS = 0 AND Id='{_idSelecionado}'";
            var dados = _bso.Consulta(query);

            // Iniciando a leitura dos dados
            dados.Inicio();

            // Verificando se há resultados
            if (dados.NumLinhas() > 0)
            {
                // Definindo as colunas esperadas na consulta
                string[] colunas = new string[] { "Codigo", "Nome", "NIPC", "AlvaraNumero", "AlvaraValidade", "CDU_NaoDivFinancas",
                                          "CDU_NaoDivSegSocial", "CDU_FolhaPagSegSocial", "CDU_ReciboApoliceAT",
                                          "CDU_ReciboRC", "CDU_Caminho", "CDU_ReciboPagSegSocial", "CDU_ApoliceAT",
                                          "CDU_ApoliceRC", "CDU_HorarioTrabalho", "CDU_DecTrabIlegais",
                                          "CDU_DecRespEstaleiro", "CDU_DecConhecimPSS", "Morada", "Localidade",
                                          "CodPostal", "CodPostalLocal", "EntidadeId", "id", "CDU_AnexoFinancas",
                                          "CDU_AnexoSegSocial", "CDU_FolhaPag" };

                // Iterando sobre as linhas dos dados
                for (int i = 0; i < dados.NumLinhas(); i++)
                {
                    // Preenchendo o dicionário com os valores de cada coluna
                    foreach (var coluna in colunas)
                    {
                        // Obtendo o valor de cada coluna para o tipo string e armazenando no dicionário
                        var valor = dados.DaValor<string>(coluna);
                        entidade[coluna] = valor; // Atribui o valor à chave correspondente
                    }

                    // Avançando para a próxima linha de dados
                    dados.Seguinte();
                }
            }
        }

        public void SetInfoEntidades(Dictionary<string, string> entidade)
        {
            _id = entidade["id"];
            _txtCodigo.Text = entidade["Codigo"];
            _txtNome.Text = entidade["Nome"];
            _txtContribuinte.Text = entidade["NIPC"];
            _txtAlvara.Text = entidade["AlvaraNumero"];

            // Tratamento da validade do alvará como um DateTimePicker
            if (!string.IsNullOrEmpty(entidade["AlvaraValidade"]))
            {
                try
                {
                    _txtAlvaraValidade.Value = Convert.ToDateTime(entidade["AlvaraValidade"]);
                    _txtAlvaraValidade.Checked = true;
                }
                catch
                {
                    _txtAlvaraValidade.Checked = false;
                }
            }
            else
            {
                _txtAlvaraValidade.Checked = false;
            }

            // Converter strings para DateTime para os campos de data
            if (!string.IsNullOrEmpty(entidade["CDU_NaoDivFinancas"]))
            {
                _txtNaoDivFinancas.Value = Convert.ToDateTime(entidade["CDU_NaoDivFinancas"]);
                _txtNaoDivFinancas.Checked = true;
            }
            else
            {
                _txtNaoDivFinancas.Checked = false;
            }

            if (!string.IsNullOrEmpty(entidade["CDU_NaoDivSegSocial"]))
            {
                _txtNaoDivSegSocial.Value = Convert.ToDateTime(entidade["CDU_NaoDivSegSocial"]);
                _txtNaoDivSegSocial.Checked = true;
            }
            else
            {
                _txtNaoDivSegSocial.Checked = false;
            }

            if (!string.IsNullOrEmpty(entidade["CDU_FolhaPagSegSocial"]))
            {
                _txtFolhaPagSegSocial.Value = Convert.ToDateTime(entidade["CDU_FolhaPagSegSocial"]);
                _txtFolhaPagSegSocial.Checked = true;
            }
            else
            {
                _txtFolhaPagSegSocial.Checked = false;
            }
            _txtReciboApoliceAT.Text = entidade["CDU_ReciboApoliceAT"];
            _txtReciboRC.Text = entidade["CDU_ReciboRC"];

            // Recupera o caminho da pasta
            _txtCaminhoPasta.Text = entidade["CDU_Caminho"];

            // Recupera os caminhos dos anexos específicos
            _caminhoAnexoFinancas = entidade["CDU_AnexoFinancas"] ?? "";
            _caminhoAnexoSegSocial = entidade["CDU_AnexoSegSocial"] ?? "";
            _caminhoAnexoFolhaPag = entidade["CDU_FolhaPag"] ?? "";

            // Atualiza os labels de anexos específicos
            AtualizarLabelsAnexos();

            // Recupera os valores do banco de dados
            string reciboPagSegSocial = entidade["CDU_ReciboPagSegSocial"];
            string apoliceAT = entidade["CDU_ApoliceAT"];
            string apoliceRC = entidade["CDU_ApoliceRC"];
            string horarioTrabalho = entidade["CDU_HorarioTrabalho"];
            string decTrabIlegais = entidade["CDU_DecTrabIlegais"];
            string decRespEstaleiro = entidade["CDU_DecRespEstaleiro"];
            string decConhecimPSS = entidade["CDU_DecConhecimPSS"];

            PreencherComboBox(_cbReciboPagSegSocial, reciboPagSegSocial);
            PreencherComboBox(_cbApoliceAT, apoliceAT);
            PreencherComboBox(_cbApoliceRC, apoliceRC);
            PreencherComboBox(_cbHorarioTrabalho, horarioTrabalho);
            PreencherComboBox(_cbDecTrabIlegais, decTrabIlegais);
            PreencherComboBox(_cbDecRespEstaleiro, decRespEstaleiro);
            PreencherComboBox(_cbDecConhecimPSS, decConhecimPSS);

            var moradaCompleta = $"{entidade["Morada"]}, {entidade["Localidade"]}, {entidade["CodPostal"]}, {entidade["CodPostalLocal"]}";

            if (moradaCompleta == ", , , ")
            {
                moradaCompleta = "";
            }
            else
            {
                _txtSede.Text = moradaCompleta;
            }

            CarregarObrasComboBox(entidade);
        }

        private void CarregarObrasComboBox(Dictionary<string, string> entidade)
        {
            var BDObras = GetObrasSumbempreiteiro(entidade["EntidadeId"]);

            _cbObras.Items.Clear(); // Limpa antes de adicionar novos itens

            while (!BDObras.NoFim())
            {
                string codigo = BDObras.DaValor<string>("Codigo");
                string descricao = BDObras.DaValor<string>("Descricao");

                // Adiciona um item ao ComboBox
                _cbObras.Items.Add(new KeyValuePair<string, string>(codigo, $"{codigo} - {descricao}"));

                BDObras.Seguinte();
            }

            _cbObras.DisplayMember = "Value"; // O que será exibido
            _cbObras.ValueMember = "Key"; // O valor interno
        }

        private StdBELista GetObrasSumbempreiteiro(string entidadeId)
        {
            var query = $@"SELECT * FROM COP_Obras 
                            WHERE Tipo = 'S' AND EntidadeIDA = '{entidadeId}'";
            var BDObras = _bso.Consulta(query);
            return BDObras;
        }

        public void GetEntidades(ref Dictionary<string, string> entidade)
        {
            string NomeLista = "Entidades";
            string Campos = "Codigo,Nome, NIPC, AlvaraNumero, AlvaraValidade, CDU_NaoDivFinancas, CDU_NaoDivSegSocial, CDU_FolhaPagSegSocial, CDU_ReciboApoliceAT, CDU_ReciboRC, CDU_Caminho, CDU_ReciboPagSegSocial, CDU_ApoliceAT, CDU_ApoliceRC, CDU_HorarioTrabalho, CDU_DecTrabIlegais, CDU_DecRespEstaleiro, CDU_DecConhecimPSS, Morada, Localidade ,CodPostal,CodPostalLocal,EntidadeId,id,CDU_AnexoFinancas,CDU_AnexoSegSocial,CDU_FolhaPag";
            string Tabela = "Geral_Entidade (NOLOCK)";
            string Where = "CDU_TrataSGS = 0";
            string CamposF4 = "Codigo,Nome, NIPC, AlvaraNumero, AlvaraValidade, CDU_NaoDivFinancas, CDU_NaoDivSegSocial, CDU_FolhaPagSegSocial, CDU_ReciboApoliceAT, CDU_ReciboRC, CDU_Caminho, CDU_ReciboPagSegSocial, CDU_ApoliceAT, CDU_ApoliceRC, CDU_HorarioTrabalho, CDU_DecTrabIlegais, CDU_DecRespEstaleiro, CDU_DecConhecimPSS, Morada, Localidade ,CodPostal,CodPostalLocal,EntidadeId,id,CDU_AnexoFinancas,CDU_AnexoSegSocial,CDU_FolhaPag";
            string orderby = "Codigo, Nome";

            List<string> ResQuery = new List<string>();

            OpenF4List(Campos, Tabela, Where, CamposF4, orderby, NomeLista, ref ResQuery);

            if (ResQuery.Count > 0)
            {
                string[] colunas = CamposF4.Split(',');
                for (int i = 0; i < colunas.Length; i++)
                {
                    if (i < ResQuery.Count)
                    {
                        entidade[colunas[i].Trim()] = ResQuery[i].ToString();
                    }
                }
            }
        }

        private void OpenF4List(string campos, string tabela, string where, string camposF4, string orderby, string nomeLista, ref List<string> resQuery)
        {
            string strSQL = "select distinct " + campos + " FROM " + tabela;

            if (where.Length > 0)
            {
                strSQL += " WHERE " + where;
            }

            strSQL += " Order by " + orderby;
            string result = Convert.ToString(_pso.Listas.GetF4SQL(nomeLista, strSQL, camposF4));

            if (!string.IsNullOrEmpty(result))
            {
                string[] itemQuery = result.Split('\t');
                resQuery.AddRange(itemQuery);
            }
        }

        private void PreencherComboBox(ComboBox comboBox, string valorBanco)
        {
            // Defina a coleção de opções que você deseja no ComboBox
            var options = new List<string> { "C", "N/C", "N/A" };

            // Define o DataSource do ComboBox
            comboBox.DataSource = options;

            // Verifica se o valor retornado é NULL ou vazio
            if (string.IsNullOrEmpty(valorBanco))
            {
                // Se for NULL ou vazio, seleciona a opção "N/A" (terceira opção)
                comboBox.SelectedItem = options[2]; // "N/A"
            }
            else
            {
                // Caso contrário, verifica se o valor está na lista
                if (options.Contains(valorBanco))
                {
                    comboBox.SelectedItem = valorBanco;
                }
                else
                {
                    // Se o valor não estiver na lista, pode-se definir um valor padrão (por exemplo, "N/A")
                    comboBox.SelectedItem = options[2]; // "N/A"
                }
            }
        }

        public void SalvarDados()
        {
            try
            {
                // Atualiza a tabela Geral_Entidade
                var querySalvar = $@"
            UPDATE Geral_Entidade
            SET 
                NIPC = '{_txtContribuinte.Text}', 
                AlvaraNumero = '{_txtAlvara.Text}', 
                AlvaraValidade = '{(_txtAlvaraValidade.Checked ? _txtAlvaraValidade.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_NaoDivFinancas = '{(_txtNaoDivFinancas.Checked ? _txtNaoDivFinancas.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_NaoDivSegSocial = '{(_txtNaoDivSegSocial.Checked ? _txtNaoDivSegSocial.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_FolhaPagSegSocial = '{(_txtFolhaPagSegSocial.Checked ? _txtFolhaPagSegSocial.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_ReciboApoliceAT = '{_txtReciboApoliceAT.Text}', 
                CDU_ReciboRC = '{_txtReciboRC.Text}', 
                CDU_Caminho = '{_txtCaminhoPasta.Text}',
                CDU_ReciboPagSegSocial = '{_cbReciboPagSegSocial.Text}', 
                CDU_ApoliceAT = '{_cbApoliceAT.Text}', 
                CDU_ApoliceRC = '{_cbApoliceRC.Text}', 
                CDU_HorarioTrabalho = '{_cbHorarioTrabalho.Text}', 
                CDU_DecTrabIlegais = '{_cbDecTrabIlegais.Text}', 
                CDU_DecRespEstaleiro = '{_cbDecRespEstaleiro.Text}', 
                CDU_DecConhecimPSS = '{_cbDecConhecimPSS.Text}',
                CDU_AnexoFinancas = '{_caminhoAnexoFinancas}',
                CDU_AnexoSegSocial = '{_caminhoAnexoSegSocial}',
                CDU_FolhaPag = '{_caminhoAnexoFolhaPag}'
            WHERE ID = '{_id}';
        ";

                _bso.DSO.ExecuteSQL(querySalvar);
                MessageBox.Show("Dados salvos com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar os dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SalvarObra()
        {
            try
            {
                // Verifica se há linhas no DataGridView
                if (_dataGridView.Rows.Count == 0)
                {
                    MessageBox.Show("Não há dados para salvar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Verifica se uma obra foi selecionada no ComboBox
                if (_cbObras.SelectedItem == null)
                {
                    MessageBox.Show("Selecione uma obra antes de salvar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Obtém o código da obra selecionada
                string codigoObraSelecionada = ((KeyValuePair<string, string>)_cbObras.SelectedItem).Key;

                // Percorre cada linha do DataGridView
                foreach (DataGridViewRow row in _dataGridView.Rows)
                {
                    // Garante que a linha não esteja vazia
                    if (row.Cells[0].Value != null)
                    {
                        string entradaObra = row.Cells[0].Value.ToString();
                        string saidaObra = row.Cells[1].Value.ToString();
                        string contratoSubempreitada = row.Cells[2].Value.ToString();
                        bool autorizacaoEntrada = Convert.ToBoolean(row.Cells[3].Value);
                        Guid id = Guid.NewGuid();

                        // Monta a query de inserção
                        string queryUpsert = $@"
    IF EXISTS (
        SELECT 1 FROM TDU_AD_Obras 
        WHERE CDU_Obra = '{codigoObraSelecionada}' 
        AND CDU_EntradaObra = '{entradaObra}'
        AND CDU_SaidaObra = '{saidaObra}'
        AND CDU_ContratoSubempreitada = '{contratoSubempreitada}'
    )
    BEGIN
        UPDATE TDU_AD_Obras 
        SET CDU_AutorizacaoEntrada = {(autorizacaoEntrada ? 1 : 0)}
        WHERE CDU_Obra = '{codigoObraSelecionada}' 
        AND CDU_EntradaObra = '{entradaObra}'
        AND CDU_SaidaObra = '{saidaObra}'
        AND CDU_ContratoSubempreitada = '{contratoSubempreitada}';
    END
    ELSE
    BEGIN
        INSERT INTO TDU_AD_Obras 
        (CDU_Codigo, CDU_Obra, CDU_EntradaObra, CDU_SaidaObra, CDU_ContratoSubempreitada, CDU_AutorizacaoEntrada) 
        VALUES 
        ('{id}', '{codigoObraSelecionada}', '{entradaObra}', '{saidaObra}', '{contratoSubempreitada}', {(autorizacaoEntrada ? 1 : 0)});
    END";

                        _bso.DSO.ExecuteSQL(queryUpsert);
                    }
                }

                MessageBox.Show("Registros adicionados com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar os dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CarregarObrasEmDataGridView(string codigoObraSelecionada)
        {
            string queryGetObras = $@"SELECT * FROM TDU_AD_Obras 
                                  WHERE CDU_Obra = '{codigoObraSelecionada}'";

            var DBObras = _bso.Consulta(queryGetObras);

            _dataGridView.Rows.Clear();

            if (DBObras.NumLinhas() > 0)
            {
                DBObras.Inicio();

                while (!DBObras.NoFim())
                {
                    _dataGridView.Rows.Add(
                        DBObras.DaValor<string>("CDU_EntradaObra"),
                        DBObras.DaValor<string>("CDU_SaidaObra"),
                        DBObras.DaValor<string>("CDU_ContratoSubempreitada"),
                        DBObras.DaValor<int>("CDU_AutorizacaoEntrada") == 1
                    );

                    int lastRowIndex = _dataGridView.Rows.Count - 1;
                    _dataGridView.Rows[lastRowIndex].DefaultCellStyle.BackColor = Color.LightYellow;

                    DBObras.Seguinte();
                }
            }
        }

        #region Gestão de Documentos
        // Method that was missing in EmpresaManager
        public void AnexarDocumento()
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(_txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro uma pasta para guardar os documentos.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png|Documentos Word|*.doc;*.docx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Selecionar Documentos para Anexar";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        foreach (string sourceFile in openFileDialog.FileNames)
                        {
                            string fileName = System.IO.Path.GetFileName(sourceFile);
                            string destFile = System.IO.Path.Combine(_txtCaminhoPasta.Text, fileName);

                            // Verifica se o arquivo já existe
                            if (System.IO.File.Exists(destFile))
                            {
                                DialogResult result = MessageBox.Show(
                                    $"O arquivo {fileName} já existe na pasta de destino. Deseja substituí-lo?",
                                    "Arquivo já existe",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question);

                                if (result == DialogResult.No)
                                    continue;
                            }

                            // Copia o arquivo para a pasta de destino
                            System.IO.File.Copy(sourceFile, destFile, true);
                        }

                        MessageBox.Show("Documentos anexados com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos
                        AtualizarListaDocumentos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documentos: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public void SelecionarPasta()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Selecione a pasta para os documentos fiscais";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    _txtCaminhoPasta.Text = folderDialog.SelectedPath;
                }
            }
        }

        public void AtualizarListaDocumentos()
        {
            // Verifica se o caminho da pasta existe
            if (string.IsNullOrEmpty(_txtCaminhoPasta.Text) || !System.IO.Directory.Exists(_txtCaminhoPasta.Text))
                return;

            try
            {
                // Obtém todos os arquivos da pasta
                string[] arquivos = System.IO.Directory.GetFiles(_txtCaminhoPasta.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao listar documentos: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AtualizarLabelsAnexos()
        {
            // Atualiza o texto nos labels que mostram os anexos específicos
            _lblAnexoFinancas.Text = string.IsNullOrEmpty(_caminhoAnexoFinancas) ?
                "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoAnexoFinancas);

            _lblAnexoSegSocial.Text = string.IsNullOrEmpty(_caminhoAnexoSegSocial) ?
                "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoAnexoSegSocial);

            _lblFolhaPagSS.Text = string.IsNullOrEmpty(_caminhoAnexoFolhaPag) ?
                "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoAnexoFolhaPag);
        }

        public void AnexarDocumentoFinancas()
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(_txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro uma pasta para guardar os documentos.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;
                openFileDialog.Title = "Selecionar Documento da Certidão de Não Dívida às Finanças";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "NaoDivFinancas_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
                                          System.IO.Path.GetExtension(sourceFile);
                        string destFile = System.IO.Path.Combine(_txtCaminhoPasta.Text, fileName);

                        // Verifica se o arquivo já existe
                        if (System.IO.File.Exists(destFile))
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
                        System.IO.File.Copy(sourceFile, destFile, true);

                        // Atualiza o caminho do anexo
                        _caminhoAnexoFinancas = destFile;

                        // Atualiza o label
                        AtualizarLabelsAnexos();

                        MessageBox.Show("Documento anexado com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos
                        AtualizarListaDocumentos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public void AnexarDocumentoSegSocial()
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(_txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro uma pasta para guardar os documentos.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;
                openFileDialog.Title = "Selecionar Documento da Certidão de Não Dívida à Segurança Social";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "NaoDivSegSocial_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
                                          System.IO.Path.GetExtension(sourceFile);
                        string destFile = System.IO.Path.Combine(_txtCaminhoPasta.Text, fileName);

                        // Verifica se o arquivo já existe
                        if (System.IO.File.Exists(destFile))
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
                        System.IO.File.Copy(sourceFile, destFile, true);

                        // Atualiza o caminho do anexo
                        _caminhoAnexoSegSocial = destFile;

                        // Atualiza o label
                        AtualizarLabelsAnexos();

                        MessageBox.Show("Documento anexado com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos
                        AtualizarListaDocumentos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public void AnexarFolhaPag()
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(_txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro uma pasta para guardar os documentos.",
                    "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os arquivos|*.*|Documentos PDF|*.pdf|Imagens|*.jpg;*.jpeg;*.png";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;
                openFileDialog.Title = "Selecionar Documento da Folha Pag.";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "FolhaPagSS_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
                                          System.IO.Path.GetExtension(sourceFile);
                        string destFile = System.IO.Path.Combine(_txtCaminhoPasta.Text, fileName);

                        // Verifica se o arquivo já existe
                        if (System.IO.File.Exists(destFile))
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
                        System.IO.File.Copy(sourceFile, destFile, true);

                        // Atualiza o caminho do anexo
                        _caminhoAnexoFolhaPag = destFile;

                        // Atualiza o label
                        AtualizarLabelsAnexos();

                        MessageBox.Show("Documento anexado com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos
                        AtualizarListaDocumentos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public void VisualizarAnexoFinancas()
        {
            if (string.IsNullOrEmpty(_caminhoAnexoFinancas) || !System.IO.File.Exists(_caminhoAnexoFinancas))
            {
                MessageBox.Show("Não existe anexo para a certidão de não dívida às Finanças.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoAnexoFinancas);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void VisualizarAnexoSegSocial()
        {
            if (string.IsNullOrEmpty(_caminhoAnexoSegSocial) || !System.IO.File.Exists(_caminhoAnexoSegSocial))
            {
                MessageBox.Show("Não existe anexo para a certidão de não dívida à Segurança Social.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoAnexoSegSocial);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void VisualizarFolhaPag()
        {
            if (string.IsNullOrEmpty(_caminhoAnexoFolhaPag) || !System.IO.File.Exists(_caminhoAnexoFolhaPag))
            {
                MessageBox.Show("Não existe anexo para a Folha Pag.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoAnexoFolhaPag);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}
