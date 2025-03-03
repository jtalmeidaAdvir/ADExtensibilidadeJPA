using Primavera.Extensibility.BusinessEntities;
using Primavera.Extensibility.CustomForm;
using StdBE100;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class Menu : CustomForm
    {
        public string _ID;

        public Menu()
        {
            InitializeComponent();
            ConfigurarEstiloControles();
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

        private void BTF4_Click(object sender, EventArgs e)
        {
            CarregarDadosEntidade();
        }

        private void CarregarDadosEntidade()
        {
            Dictionary<string, string> entidade = new Dictionary<string, string>();
            GetEntidades(ref entidade);

            if (entidade.Count > 0)
            {
                SetInfoEntidades(entidade);
            }
        }

        private void SetInfoEntidades(Dictionary<string, string> entidade)
        {
            _ID = entidade["id"];
            TXT_Codigo.Text = entidade["Codigo"];
            TXT_Nome.Text = entidade["Nome"];
            TXT_Contribuinte.Text = entidade["NIPC"];
            TXT_Alvara.Text = entidade["AlvaraNumero"];

            // Tratamento da validade do alvará como um DateTimePicker
            if (!string.IsNullOrEmpty(entidade["AlvaraValidade"]))
            {
                try
                {
                    TXT_AlvaraValidade.Value = Convert.ToDateTime(entidade["AlvaraValidade"]);
                    TXT_AlvaraValidade.Checked = true;
                }
                catch
                {
                    TXT_AlvaraValidade.Checked = false;
                }
            }
            else
            {
                TXT_AlvaraValidade.Checked = false;
            }

            // Converter strings para DateTime para os campos de data
            if (!string.IsNullOrEmpty(entidade["CDU_NaoDivFinancas"]))
            {
                TXT_NaoDivFinancas.Value = Convert.ToDateTime(entidade["CDU_NaoDivFinancas"]);
                TXT_NaoDivFinancas.Checked = true;
            }
            else
            {
                TXT_NaoDivFinancas.Checked = false;
            }

            if (!string.IsNullOrEmpty(entidade["CDU_NaoDivSegSocial"]))
            {
                TXT_NaoDivSegSocial.Value = Convert.ToDateTime(entidade["CDU_NaoDivSegSocial"]);
                TXT_NaoDivSegSocial.Checked = true;
            }
            else
            {
                TXT_NaoDivSegSocial.Checked = false;
            }

            if (!string.IsNullOrEmpty(entidade["CDU_FolhaPagSegSocial"]))
            {
                TXT_FolhaPagSegSocial.Value = Convert.ToDateTime(entidade["CDU_FolhaPagSegSocial"]);
                TXT_FolhaPagSegSocial.Checked = true;
            }
            else
            {
                TXT_FolhaPagSegSocial.Checked = false;
            }
            TXT_ReciboApoliceAT.Text = entidade["CDU_ReciboApoliceAT"];
            TXT_ReciboRC.Text = entidade["CDU_ReciboRC"];

            // Recupera o caminho da pasta
            txtCaminhoPasta.Text = entidade["CDU_Caminho"];

            // Atualiza a lista de documentos ao carregar a entidade
            AtualizarListaDocumentos();

            // Recupera os valores do banco de dados
            string reciboPagSegSocial = entidade["CDU_ReciboPagSegSocial"];
            string apoliceAT = entidade["CDU_ApoliceAT"];
            string apoliceRC = entidade["CDU_ApoliceRC"];
            string horarioTrabalho = entidade["CDU_HorarioTrabalho"];
            string decTrabIlegais = entidade["CDU_DecTrabIlegais"];
            string decRespEstaleiro = entidade["CDU_DecRespEstaleiro"];
            string decConhecimPSS = entidade["CDU_DecConhecimPSS"];

            PreencherComboBox(cb_ReciboPagSegSocial, reciboPagSegSocial);
            PreencherComboBox(cb_ApoliceAT, apoliceAT);
            PreencherComboBox(cb_ApoliceRC, apoliceRC);
            PreencherComboBox(cb_HorarioTrabalho, horarioTrabalho);
            PreencherComboBox(cb_DecTrabIlegais, decTrabIlegais);
            PreencherComboBox(cb_DecRespEstaleiro, decRespEstaleiro);
            PreencherComboBox(cb_DecConhecimPSS, decConhecimPSS);

            var moradaCompleta = $"{entidade["Morada"]}, {entidade["Localidade"]}, {entidade["CodPostal"]}, {entidade["CodPostalLocal"]}";

            if (moradaCompleta == ", , , ")
            {
                moradaCompleta = "";
            }
            else
            {
                TXT_Sede.Text = moradaCompleta;
            }

            CarregarObrasComboBox(entidade);
        }

        private void cb_Obras_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Obras.SelectedItem is KeyValuePair<string, string> obraSelecionada)
            {
                string codigoObraSelecionada = obraSelecionada.Key;
                string queryGetObras = $@"SELECT * FROM TDU_AD_Obras 
                                  WHERE CDU_Obra = '{codigoObraSelecionada}'";

                var DBObras = BSO.Consulta(queryGetObras);

                dataGridView1.Rows.Clear();

                if (DBObras.NumLinhas() > 0)
                {
                    DBObras.Inicio();

                    while (!DBObras.NoFim())
                    {
                        dataGridView1.Rows.Add(
                            DBObras.DaValor<string>("CDU_EntradaObra"),
                            DBObras.DaValor<string>("CDU_SaidaObra"),
                            DBObras.DaValor<string>("CDU_ContratoSubempreitada"),
                            DBObras.DaValor<int>("CDU_AutorizacaoEntrada") == 1
                        );

                        int lastRowIndex = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[lastRowIndex].DefaultCellStyle.BackColor = Color.LightYellow;

                        DBObras.Seguinte();
                    }
                }
            }
        }

        private void CarregarObrasComboBox(Dictionary<string, string> entidade)
        {
            var BDObras = GetObrasSumbempreiteiro(entidade["EntidadeId"]);

            cb_Obras.Items.Clear(); // Limpa antes de adicionar novos itens

            while (!BDObras.NoFim())
            {
                string codigo = BDObras.DaValor<string>("Codigo");
                string descricao = BDObras.DaValor<string>("Descricao");

                // Adiciona um item ao ComboBox
                cb_Obras.Items.Add(new KeyValuePair<string, string>(codigo, $"{codigo} - {descricao}"));

                BDObras.Seguinte();
            }

            cb_Obras.DisplayMember = "Value"; // O que será exibido
            cb_Obras.ValueMember = "Key"; // O valor interno
        }

        private StdBELista GetObrasSumbempreiteiro(string entidadeId)
        {
            var query = $@"SELECT * FROM COP_Obras 
                            WHERE Tipo = 'S' AND EntidadeIDA = '{entidadeId}'";
            var BDObras = BSO.Consulta(query);
            return BDObras;
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

        private void btnSelecionarPasta_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Selecione a pasta para os documentos fiscais";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtCaminhoPasta.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void btnAnexarDocumento_Click(object sender, EventArgs e)
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(txtCaminhoPasta.Text))
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
                            string destFile = System.IO.Path.Combine(txtCaminhoPasta.Text, fileName);

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

                        // Atualiza a lista de documentos (se implementada)
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

        private void AtualizarListaDocumentos()
        {
            // Limpa a lista atual
            listBoxDocumentos.Items.Clear();

            // Verifica se o caminho da pasta existe
            if (string.IsNullOrEmpty(txtCaminhoPasta.Text) || !System.IO.Directory.Exists(txtCaminhoPasta.Text))
                return;

            try
            {
                // Obtém todos os arquivos da pasta
                string[] arquivos = System.IO.Directory.GetFiles(txtCaminhoPasta.Text);

                // Adiciona cada arquivo à lista
                foreach (string arquivo in arquivos)
                {
                    listBoxDocumentos.Items.Add(System.IO.Path.GetFileName(arquivo));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao listar documentos: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVisualizarDocumento_Click(object sender, EventArgs e)
        {
            // Verifica se há um documento selecionado
            if (listBoxDocumentos.SelectedItem == null)
            {
                MessageBox.Show("Por favor, selecione um documento para visualizar.",
                    "Nenhum documento selecionado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Obtém o caminho completo do arquivo
                string nomeArquivo = listBoxDocumentos.SelectedItem.ToString();
                string caminhoCompleto = System.IO.Path.Combine(txtCaminhoPasta.Text, nomeArquivo);

                // Verifica se o arquivo existe
                if (!System.IO.File.Exists(caminhoCompleto))
                {
                    MessageBox.Show("O arquivo selecionado não existe mais na pasta.",
                        "Arquivo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Abre o arquivo com o programa padrão do sistema
                System.Diagnostics.Process.Start(caminhoCompleto);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRemoverDocumento_Click(object sender, EventArgs e)
        {
            // Verifica se há um documento selecionado
            if (listBoxDocumentos.SelectedItem == null)
            {
                MessageBox.Show("Por favor, selecione um documento para remover.",
                    "Nenhum documento selecionado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Obtém o caminho completo do arquivo
                string nomeArquivo = listBoxDocumentos.SelectedItem.ToString();
                string caminhoCompleto = System.IO.Path.Combine(txtCaminhoPasta.Text, nomeArquivo);

                // Confirma a remoção
                DialogResult result = MessageBox.Show(
                    $"Tem certeza que deseja remover o documento '{nomeArquivo}'?",
                    "Confirmar remoção",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Verifica se o arquivo existe
                    if (System.IO.File.Exists(caminhoCompleto))
                    {
                        System.IO.File.Delete(caminhoCompleto);
                        MessageBox.Show("Documento removido com sucesso!",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Atualiza a lista de documentos
                        AtualizarListaDocumentos();
                    }
                    else
                    {
                        MessageBox.Show("O arquivo selecionado não existe mais na pasta.",
                            "Arquivo não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao remover o documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetEntidades(ref Dictionary<string, string> entidade)
        {
            string NomeLista = "Entidades";
            string Campos = "Codigo,Nome, NIPC, AlvaraNumero, AlvaraValidade, CDU_NaoDivFinancas, CDU_NaoDivSegSocial, CDU_FolhaPagSegSocial, CDU_ReciboApoliceAT, CDU_ReciboRC, CDU_Caminho, CDU_ReciboPagSegSocial, CDU_ApoliceAT, CDU_ApoliceRC, CDU_HorarioTrabalho, CDU_DecTrabIlegais, CDU_DecRespEstaleiro, CDU_DecConhecimPSS, Morada, Localidade ,CodPostal,CodPostalLocal,EntidadeId,id";
            string Tabela = "Geral_Entidade (NOLOCK)";
            string Where = "CDU_TrataSGS = 0";
            string CamposF4 = "Codigo,Nome, NIPC, AlvaraNumero, AlvaraValidade, CDU_NaoDivFinancas, CDU_NaoDivSegSocial, CDU_FolhaPagSegSocial, CDU_ReciboApoliceAT, CDU_ReciboRC, CDU_Caminho, CDU_ReciboPagSegSocial, CDU_ApoliceAT, CDU_ApoliceRC, CDU_HorarioTrabalho, CDU_DecTrabIlegais, CDU_DecRespEstaleiro, CDU_DecConhecimPSS, Morada, Localidade ,CodPostal,CodPostalLocal,EntidadeId,id";
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
            string result = Convert.ToString(PSO.Listas.GetF4SQL(nomeLista, strSQL, camposF4));

            if (!string.IsNullOrEmpty(result))
            {
                string[] itemQuery = result.Split('\t');
                resQuery.AddRange(itemQuery);
            }
        }

        private void BT_Salvar_Click_Click(object sender, EventArgs e)
        {
            try
            {
                // Atualiza a tabela Geral_Entidade
                var querySalvar = $@"
            UPDATE Geral_Entidade
            SET 
                NIPC = '{TXT_Contribuinte.Text}', 
                AlvaraNumero = '{TXT_Alvara.Text}', 
                AlvaraValidade = '{(TXT_AlvaraValidade.Checked ? TXT_AlvaraValidade.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_NaoDivFinancas = '{(TXT_NaoDivFinancas.Checked ? TXT_NaoDivFinancas.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_NaoDivSegSocial = '{(TXT_NaoDivSegSocial.Checked ? TXT_NaoDivSegSocial.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_FolhaPagSegSocial = '{(TXT_FolhaPagSegSocial.Checked ? TXT_FolhaPagSegSocial.Value.ToString("yyyy-MM-dd") : "")}', 
                CDU_ReciboApoliceAT = '{TXT_ReciboApoliceAT.Text}', 
                CDU_ReciboRC = '{TXT_ReciboRC.Text}', 
                CDU_Caminho = '{txtCaminhoPasta.Text}',
                CDU_ReciboPagSegSocial = '{cb_ReciboPagSegSocial.Text}', 
                CDU_ApoliceAT = '{cb_ApoliceAT.Text}', 
                CDU_ApoliceRC = '{cb_ApoliceRC.Text}', 
                CDU_HorarioTrabalho = '{cb_HorarioTrabalho.Text}', 
                CDU_DecTrabIlegais = '{cb_DecTrabIlegais.Text}', 
                CDU_DecRespEstaleiro = '{cb_DecRespEstaleiro.Text}', 
                CDU_DecConhecimPSS = '{cb_DecConhecimPSS.Text}'
            WHERE ID = '{_ID}';
        ";

                BSO.DSO.ExecuteSQL(querySalvar);
                MessageBox.Show("Dados salvos com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar os dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Verifica se há linhas no DataGridView
                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Não há dados para salvar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Verifica se uma obra foi selecionada no ComboBox
                if (cb_Obras.SelectedItem == null)
                {
                    MessageBox.Show("Selecione uma obra antes de salvar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Obtém o código da obra selecionada
                string codigoObraSelecionada = ((KeyValuePair<string, string>)cb_Obras.SelectedItem).Key;

                // Percorre cada linha do DataGridView
                foreach (DataGridViewRow row in dataGridView1.Rows)
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

                        BSO.DSO.ExecuteSQL(queryUpsert);
                    }
                }

                MessageBox.Show("Registros adicionados com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar os dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
