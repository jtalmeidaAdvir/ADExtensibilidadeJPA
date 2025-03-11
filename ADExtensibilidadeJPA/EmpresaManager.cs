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
                                      "CDU_AnexoSegSocial", "CDU_FolhaPag", "CDU_AnexoApoliceAT",
                                      "CDU_AnexoApoliceRC", "CDU_AnexoHorarioTrabalho",
                                      "CDU_AnexoD", "CDU_DecTrabEmigr", "CDU_InscricaoSS",
                                      "CDU_AnexoDStatus", "CDU_DecTrabEmigrStatus", "CDU_InscricaoSSStatus" };

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
            _caminhoAnexoApoliceAT = entidade["CDU_AnexoApoliceAT"] ?? "";
            _caminhoAnexoApoliceRC = entidade["CDU_AnexoApoliceRC"] ?? "";
            _caminhoAnexoHorarioTrabalho = entidade["CDU_AnexoHorarioTrabalho"] ?? "";
            _caminhoAnexoD = entidade["CDU_AnexoD"] ?? "";
            _caminhoDecTrabEmigr = entidade["CDU_DecTrabEmigr"] ?? "";
            _caminhoInscricaoSS = entidade["CDU_InscricaoSS"] ?? "";

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
                CDU_FolhaPag = '{_caminhoAnexoFolhaPag}',
                CDU_AnexoApoliceAT = '{_caminhoAnexoApoliceAT}',
                CDU_AnexoApoliceRC = '{_caminhoAnexoApoliceRC}',
                CDU_AnexoHorarioTrabalho = '{_caminhoAnexoHorarioTrabalho}',
                CDU_AnexoD = '{_caminhoAnexoD}',
                CDU_DecTrabEmigr = '{_caminhoDecTrabEmigr}',
                CDU_InscricaoSS = '{_caminhoInscricaoSS}'
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

            // Atualiza os novos labels adicionados
            Control lblAnexoApoliceAT = _lblFolhaPagSS.Parent.Controls.Find("lblAnexoApoliceAT", true).FirstOrDefault();
            if (lblAnexoApoliceAT != null)
            {
                lblAnexoApoliceAT.Text = string.IsNullOrEmpty(_caminhoAnexoApoliceAT) ?
                    "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoAnexoApoliceAT);
            }

            Control lblAnexoApoliceRC = _lblFolhaPagSS.Parent.Controls.Find("lblAnexoApoliceRC", true).FirstOrDefault();
            if (lblAnexoApoliceRC != null)
            {
                lblAnexoApoliceRC.Text = string.IsNullOrEmpty(_caminhoAnexoApoliceRC) ?
                    "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoAnexoApoliceRC);
            }

            Control lblAnexoHorarioTrabalho = _lblFolhaPagSS.Parent.Controls.Find("lblAnexoHorarioTrabalho", true).FirstOrDefault();
            if (lblAnexoHorarioTrabalho != null)
            {
                lblAnexoHorarioTrabalho.Text = string.IsNullOrEmpty(_caminhoAnexoHorarioTrabalho) ?
                    "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoAnexoHorarioTrabalho);
            }

            Control lblAnexoD = _lblFolhaPagSS.Parent.Controls.Find("lblAnexoD", true).FirstOrDefault();
            if (lblAnexoD != null)
            {
                lblAnexoD.Text = string.IsNullOrEmpty(_caminhoAnexoD) ?
                    "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoAnexoD);
            }

            Control lblDecTrabEmigr = _lblFolhaPagSS.Parent.Controls.Find("lblDecTrabEmigr", true).FirstOrDefault();
            if (lblDecTrabEmigr != null)
            {
                lblDecTrabEmigr.Text = string.IsNullOrEmpty(_caminhoDecTrabEmigr) ?
                    "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoDecTrabEmigr);
            }

            Control lblInscricaoSS = _lblFolhaPagSS.Parent.Controls.Find("lblInscricaoSS", true).FirstOrDefault();
            if (lblInscricaoSS != null)
            {
                lblInscricaoSS.Text = string.IsNullOrEmpty(_caminhoInscricaoSS) ?
                    "Nenhum anexo" : System.IO.Path.GetFileName(_caminhoInscricaoSS);
            }
        }

        public void AnexarDocumentoFinancas(string responsavel = "")
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(_txtCaminhoPasta.Text) || !System.IO.Directory.Exists(_txtCaminhoPasta.Text))
            {
                MessageBox.Show("Por favor, selecione primeiro o caminho da pasta para armazenar os anexos.",
                    "Caminho não selecionado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Armazenar o responsável (poderia ser salvo em um campo adicional no banco de dados)
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

        public void AnexarDocumentoSegSocial(string responsavel = "")
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

        public void AnexarFolhaPag(string responsavel = "")
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

        // Variáveis para armazenar os caminhos dos anexos adicionais
        private string _caminhoAnexoApoliceAT = "";
        private string _caminhoAnexoApoliceRC = "";
        private string _caminhoAnexoHorarioTrabalho = "";
        private string _caminhoAnexoD = "";
        private string _caminhoDecTrabEmigr = "";
        private string _caminhoInscricaoSS = "";

        public void AnexarDocumentoApoliceAT(string responsavel = "")
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
                openFileDialog.Title = "Selecionar Documento da Apólice de Seguro de Acidentes de Trabalho";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "ApoliceAT_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
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
                        _caminhoAnexoApoliceAT = destFile;

                        // Atualiza o sistema
                        AtualizarStatusAnexos();

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

        public void AnexarDocumentoApoliceRC(string responsavel = "")
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
                openFileDialog.Title = "Selecionar Documento da Apólice de Responsabilidade Civil";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "ApoliceRC_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
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
                        _caminhoAnexoApoliceRC = destFile;

                        // Atualiza o sistema
                        AtualizarStatusAnexos();

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

        public void AnexarHorarioTrabalho(string responsavel = "", DateTime? validade = null)
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
                openFileDialog.Title = "Selecionar Documento de Horário de Trabalho";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "HorarioTrabalho_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
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
                        _caminhoAnexoHorarioTrabalho = destFile;

                        // Atualiza o sistema
                        AtualizarStatusAnexos();

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

        public void AnexarAnexoD(string responsavel = "", DateTime? validade = null)
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
                openFileDialog.Title = "Selecionar Anexo D RU 2023";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "AnexoD_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
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
                        _caminhoAnexoD = destFile;

                        // Atualiza o sistema
                        AtualizarStatusAnexos();

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

        public void AnexarDecTrabEmigr(string responsavel = "", DateTime? validade = null)
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
                openFileDialog.Title = "Selecionar Declaração de Trabalhadores Emigrantes";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "DecTrabEmigr_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
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
                        _caminhoDecTrabEmigr = destFile;

                        // Atualiza o sistema
                        AtualizarStatusAnexos();

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

        public void AnexarInscricaoSS(string responsavel = "", DateTime? validade = null)
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
                openFileDialog.Title = "Selecionar Documento de Inscrição na Segurança Social";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "InscricaoSS_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
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
                        _caminhoInscricaoSS = destFile;

                        // Atualiza o sistema
                        AtualizarStatusAnexos();

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

        public void VisualizarApoliceAT()
        {
            if (string.IsNullOrEmpty(_caminhoAnexoApoliceAT) || !System.IO.File.Exists(_caminhoAnexoApoliceAT))
            {
                MessageBox.Show("Não existe anexo para a Apólice de Seguro de Acidentes de Trabalho.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoAnexoApoliceAT);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void VisualizarApoliceRC()
        {
            if (string.IsNullOrEmpty(_caminhoAnexoApoliceRC) || !System.IO.File.Exists(_caminhoAnexoApoliceRC))
            {
                MessageBox.Show("Não existe anexo para a Apólice de Responsabilidade Civil.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoAnexoApoliceRC);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void VisualizarHorarioTrabalho()
        {
            if (string.IsNullOrEmpty(_caminhoAnexoHorarioTrabalho) || !System.IO.File.Exists(_caminhoAnexoHorarioTrabalho))
            {
                MessageBox.Show("Não existe anexo para o Horário de Trabalho.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoAnexoHorarioTrabalho);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void VisualizarAnexoD()
        {
            if (string.IsNullOrEmpty(_caminhoAnexoD) || !System.IO.File.Exists(_caminhoAnexoD))
            {
                MessageBox.Show("Não existe anexo para o Anexo D RU 2023.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoAnexoD);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void VisualizarDecTrabEmigr()
        {
            if (string.IsNullOrEmpty(_caminhoDecTrabEmigr) || !System.IO.File.Exists(_caminhoDecTrabEmigr))
            {
                MessageBox.Show("Não existe anexo para a Declaração de Trabalhadores Emigrantes.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoDecTrabEmigr);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void VisualizarInscricaoSS()
        {
            if (string.IsNullOrEmpty(_caminhoInscricaoSS) || !System.IO.File.Exists(_caminhoInscricaoSS))
            {
                MessageBox.Show("Não existe anexo para Inscrição na Segurança Social.",
                    "Anexo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(_caminhoInscricaoSS);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o anexo: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AtualizarStatusAnexos()
        {
            // Atualizar labels para todos os anexos
            AtualizarLabelsAnexos();

            // Atualizar base de dados com novos caminhos
            try
            {
                var query = $@"
                    UPDATE Geral_Entidade
                    SET 
                        CDU_AnexoApoliceAT = '{_caminhoAnexoApoliceAT}',
                        CDU_AnexoApoliceRC = '{_caminhoAnexoApoliceRC}',
                        CDU_AnexoHorarioTrabalho = '{_caminhoAnexoHorarioTrabalho}',
                        CDU_AnexoD = '{_caminhoAnexoD}',
                        CDU_DecTrabEmigr = '{_caminhoDecTrabEmigr}',
                        CDU_InscricaoSS = '{_caminhoInscricaoSS}'
                    WHERE ID = '{_id}';
                ";

                _bso.DSO.ExecuteSQL(query);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar a base de dados: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool[] GetDocumentosAnexados()
        {
            // Array para armazenar status de cada documento (true = anexado, false = não anexado)
            bool[] documentosAnexados = new bool[9];

            // Verificar cada documento
            documentosAnexados[0] = !string.IsNullOrEmpty(_caminhoAnexoFinancas);
            documentosAnexados[1] = !string.IsNullOrEmpty(_caminhoAnexoSegSocial);
            documentosAnexados[2] = !string.IsNullOrEmpty(_caminhoAnexoFolhaPag);
            documentosAnexados[3] = !string.IsNullOrEmpty(_caminhoAnexoApoliceAT);
            documentosAnexados[4] = !string.IsNullOrEmpty(_caminhoAnexoApoliceRC);
            documentosAnexados[5] = !string.IsNullOrEmpty(_caminhoAnexoHorarioTrabalho);
            documentosAnexados[6] = !string.IsNullOrEmpty(_caminhoAnexoD);
            documentosAnexados[7] = !string.IsNullOrEmpty(_caminhoDecTrabEmigr);
            documentosAnexados[8] = !string.IsNullOrEmpty(_caminhoInscricaoSS);

            return documentosAnexados;
        }

        public void AbrirPastaAnexos()
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(_txtCaminhoPasta.Text) || !System.IO.Directory.Exists(_txtCaminhoPasta.Text))
            {
                MessageBox.Show("Pasta de anexos não definida ou não existente. Por favor, selecione uma pasta válida primeiro.",
                    "Pasta não encontrada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Abre a pasta no explorador de arquivos
                System.Diagnostics.Process.Start("explorer.exe", _txtCaminhoPasta.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir a pasta: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void AnexarDocumento(string responsavel = "", DateTime? validade = null)
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
                openFileDialog.Title = "Selecionar Documento";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string sourceFile = openFileDialog.FileName;
                        string fileName = "Documento_" + _txtCodigo.Text + "_" + DateTime.Now.ToString("yyyyMMdd") +
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

                        // Exibe mensagem de sucesso
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

        public void VerificarDocumentosFaltantes()
        {
            List<string> documentosFaltantes = new List<string>();

            // Verificar todos os documentos obrigatórios
            if (string.IsNullOrEmpty(_caminhoAnexoFinancas))
                documentosFaltantes.Add("Certidão de Não Dívida às Finanças");

            if (string.IsNullOrEmpty(_caminhoAnexoSegSocial))
                documentosFaltantes.Add("Certidão de Não Dívida à Segurança Social");

            if (string.IsNullOrEmpty(_caminhoAnexoFolhaPag))
                documentosFaltantes.Add("Folha de Pagamento da Segurança Social");

            if (string.IsNullOrEmpty(_caminhoAnexoApoliceAT))
                documentosFaltantes.Add("Apólice de Seguro de Acidentes de Trabalho");

            if (string.IsNullOrEmpty(_caminhoAnexoApoliceRC))
                documentosFaltantes.Add("Apólice de Responsabilidade Civil");

            if (string.IsNullOrEmpty(_caminhoAnexoHorarioTrabalho))
                documentosFaltantes.Add("Horário de Trabalho");

            if (string.IsNullOrEmpty(_caminhoAnexoD))
                documentosFaltantes.Add("Anexo D RU 2023");

            if (string.IsNullOrEmpty(_caminhoDecTrabEmigr))
                documentosFaltantes.Add("Declaração de Trabalhadores Emigrantes");

            if (string.IsNullOrEmpty(_caminhoInscricaoSS))
                documentosFaltantes.Add("Inscrição na Segurança Social");

            // Verificar status dos comboboxes quando estão como "N/C" (Não Conforme)
            if (_cbReciboPagSegSocial.Text == "N/C")
                documentosFaltantes.Add("Recibo de Pagamento à Segurança Social");

            if (_cbApoliceAT.Text == "N/C")
                documentosFaltantes.Add("Conformidade da Apólice AT");

            if (_cbApoliceRC.Text == "N/C")
                documentosFaltantes.Add("Conformidade da Apólice RC");

            if (_cbHorarioTrabalho.Text == "N/C")
                documentosFaltantes.Add("Conformidade do Horário de Trabalho");

            if (_cbDecTrabIlegais.Text == "N/C")
                documentosFaltantes.Add("Declaração de Trabalhadores Ilegais");

            if (_cbDecRespEstaleiro.Text == "N/C")
                documentosFaltantes.Add("Declaração de Responsabilidade de Estaleiro");

            if (_cbDecConhecimPSS.Text == "N/C")
                documentosFaltantes.Add("Declaração de Conhecimento PSS");

            // Exibir resultado
            if (documentosFaltantes.Count > 0)
            {
                string mensagem = "Documentos em falta:\n\n";
                foreach (var doc in documentosFaltantes)
                {
                    mensagem += "• " + doc + "\n";
                }

                // Criar um formulário personalizado para exibir a lista
                Form frmDocumentosFaltantes = new Form
                {
                    Text = "Documentos em Falta - " + _txtNome.Text,
                    Width = 450,
                    Height = 400,
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false
                };

                // Adicionar um TextBox para exibir a lista
                TextBox txtLista = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    BackColor = Color.White,
                    ForeColor = Color.Red,
                    Font = new Font("Calibri", 10),
                    Text = mensagem,
                    Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Vertical
                };

                // Adicionar um painel de título
                Panel panelTitulo = new Panel
                {
                    Height = 40,
                    Dock = DockStyle.Top,
                    BackColor = Color.FromArgb(59, 89, 152)
                };

                Label lblTitulo = new Label
                {
                    Text = "Lista de Documentos em Falta",
                    ForeColor = Color.White,
                    Font = new Font("Calibri", 12, FontStyle.Bold),
                    AutoSize = true,
                    Location = new Point(10, 10)
                };
                panelTitulo.Controls.Add(lblTitulo);

                // Adicionar botão OK
                Button btnOK = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Width = 100,
                    Height = 30,
                    Dock = DockStyle.Bottom,
                    BackColor = Color.LightSteelBlue,
                    FlatStyle = FlatStyle.Flat
                };

                // Adicionar controles ao formulário
                frmDocumentosFaltantes.Controls.Add(txtLista);
                frmDocumentosFaltantes.Controls.Add(btnOK);
                frmDocumentosFaltantes.Controls.Add(panelTitulo);

                // Exibir o formulário
                frmDocumentosFaltantes.ShowDialog();
            }
            else
            {
                MessageBox.Show("Todos os documentos obrigatórios estão presentes e conformes.",
                    "Verificação de Documentos", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion
    }
}