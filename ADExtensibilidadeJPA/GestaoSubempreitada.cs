using ErpBS100;
using StdBE100;
using StdPlatBS100;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class GestaoSubempreitada : Form
    {
        private readonly ErpBS _BSO;
        private readonly StdBSInterfPub _PSO;
        private readonly string _idSelecionado;
        public GestaoSubempreitada(ErpBS BSO, StdBSInterfPub PSO, string idSelecionado)
        {
            InitializeComponent();
            _BSO = BSO;
            _PSO = PSO;
            _idSelecionado = idSelecionado;
            CarregarDados();
            InitializeButtonEvents();
        }


        public void CarregarDados()
        {
            Dictionary<string, string> entidade = new Dictionary<string, string>();
            GetEntidadesID(ref entidade);
            if (entidade.Count > 0)
            {
                SetInfoEntidades(entidade);
                CarregarStatusDocumentos();
            }
        }

        private void CarregarStatusDocumentos()
        {
            try
            {
                // Primeiro verificar e criar colunas se não existirem
                VerificarECriarColunas();

                // Consulta para obter os campos de documentos anexados
                string query = $@"SELECT 
                    CDU_AnexoFinancas, CDU_ValidadeFinancas,
                    CDU_AnexoSegSocial, CDU_ValidadeSegSocial,
                    CDU_AnexoFolhaPag, CDU_ValidadeFolhaPag,
                    CDU_AnexoComprovativoPagamento, CDU_ValidadeComprovativoPagamento,
                    CDU_AnexoReciboSeguroAT, CDU_ValidadeReciboSeguroAT,
                    CDU_AnexoSeguroRC, CDU_ValidadeSeguroRC,
                    CDU_AnexoHorarioTrabalho, CDU_ValidadeHorarioTrabalho,
                    CDU_AnexoSeguroAT, CDU_ValidadeSeguroAT,
                    CDU_AnexoAlvara, CDU_ValidadeAlvara,
                    CDU_AnexoCertidaoPermanente, CDU_ValidadeCertidaoPermanente,
                    CDU_AnexoContrato, CDU_ValidadeContrato,
                    CDU_AnexoDeclaracaoPSS, CDU_ValidadeDeclaracaoPSS,
                    CDU_AnexoResponsavelEstaleiro, CDU_ValidadeResponsavelEstaleiro
                    FROM Geral_Entidade WHERE id = '{_idSelecionado}'";

                var dados = _BSO.Consulta(query);

                if (dados.NumLinhas() > 0)
                {
                    dados.Inicio();

                    try
                    {
                        // Atualizar checkboxes com base nos valores do banco de dados
                        SeguroUpdateCheckboxFromDB(checkBox1, dados, "CDU_AnexoFinancas", "Finanças", "CDU_ValidadeFinancas");
                        SeguroUpdateCheckboxFromDB(checkBox2, dados, "CDU_AnexoSegSocial", "Segurança Social", "CDU_ValidadeSegSocial");
                        SeguroUpdateCheckboxFromDB(checkBox3, dados, "CDU_AnexoFolhaPag", "Folha Pagamento", "CDU_ValidadeFolhaPag");
                        SeguroUpdateCheckboxFromDB(checkBox4, dados, "CDU_AnexoComprovativoPagamento", "Comprovativo Pagamento", "CDU_ValidadeComprovativoPagamento");
                        SeguroUpdateCheckboxFromDB(checkBox5, dados, "CDU_AnexoReciboSeguroAT", "Seguro AT", "CDU_ValidadeReciboSeguroAT");
                        SeguroUpdateCheckboxFromDB(checkBox6, dados, "CDU_AnexoSeguroRC", "Seguro RC", "CDU_ValidadeSeguroRC");
                        SeguroUpdateCheckboxFromDB(checkBox7, dados, "CDU_AnexoHorarioTrabalho", "Horário Trabalho", "CDU_ValidadeHorarioTrabalho");
                        SeguroUpdateCheckboxFromDB(checkBox8, dados, "CDU_AnexoSeguroAT", "Condições Seguro AT", "CDU_ValidadeSeguroAT");
                        SeguroUpdateCheckboxFromDB(checkBox9, dados, "CDU_AnexoAlvara", "Alvará", "CDU_ValidadeAlvara");
                        SeguroUpdateCheckboxFromDB(checkBox10, dados, "CDU_AnexoCertidaoPermanente", "Certidão Permanente", "CDU_ValidadeCertidaoPermanente");
                        SeguroUpdateCheckboxFromDB(checkBox11, dados, "CDU_AnexoContrato", "Contrato", "CDU_ValidadeContrato");
                        SeguroUpdateCheckboxFromDB(checkBox12, dados, "CDU_AnexoDeclaracaoPSS", "Declaração PSS", "CDU_ValidadeDeclaracaoPSS");
                        SeguroUpdateCheckboxFromDB(checkBox13, dados, "CDU_AnexoResponsavelEstaleiro", "Responsável Estaleiro", "CDU_ValidadeResponsavelEstaleiro");
                    }
                    catch (FormatException fex)
                    {
                        MessageBox.Show($"Erro ao carregar status dos documentos: Cadeia de caracteres de entrada com formato incorreto. Detalhes: {fex.Message}",
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao carregar status dos documentos: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateCheckboxFromDB(CheckBox checkBox, int anexado, string tipoDoc, DateTime? validade)
        {
            if (anexado == 1)
            {
                checkBox.Checked = true;
                checkBox.Enabled = true;
                checkBox.Text = validade.HasValue ?
                    $"{tipoDoc} (Válido até: {validade.Value.ToShortDateString()})" :
                    $"{tipoDoc} (Anexado)";
                checkBox.AutoSize = true;
            }
        }

        private void SeguroUpdateCheckboxFromDB(CheckBox checkBox, StdBELista dados, string colunaNome, string tipoDoc, string colunaValidade)
        {
            try
            {
                // Verifica se a coluna existe
                // Verifica se a coluna existe de outra forma
                try
                {
                    // Tentar acessar a coluna para verificar se existe
                    var testeColuna = dados.Valor(colunaNome);
                }
                catch
                {
                    // Se lançar exceção, a coluna não existe
                    string query = $@"
                    IF NOT EXISTS (
                        SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = '{colunaNome}'
                    )
                    BEGIN
                        ALTER TABLE Geral_Entidade ADD {colunaNome} INT DEFAULT 0
                    END";
                    _BSO.DSO.ExecuteSQL(query);
                    return;
                }

                // Verifica se a coluna de validade existe, e se não, criá-la
                try
                {
                    var testeColuna = dados.Valor(colunaValidade);
                }
                catch
                {
                    string query = $@"
                    IF NOT EXISTS (
                        SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = '{colunaValidade}'
                    )
                    BEGIN
                        ALTER TABLE Geral_Entidade ADD {colunaValidade} DATE NULL
                    END";
                    _BSO.DSO.ExecuteSQL(query);
                }

                // Obtém o valor de anexo seguramente (se for nulo ou não puder ser convertido, assume 0)
                int anexado = 0;
                try
                {
                    // Tenta obter o valor como string
                    string valorString = dados.Valor(colunaNome) as string;

                    // Verifica se é nulo ou vazio
                    if (!string.IsNullOrEmpty(valorString))
                    {
                        // Tenta converter para int usando TryParse
                        int valorInt;
                        if (int.TryParse(valorString, out valorInt))
                        {
                            anexado = valorInt;
                        }
                    }
                }
                catch
                {
                    anexado = 0;
                }

                if (anexado == 1)
                {
                    checkBox.Checked = true;
                    checkBox.Enabled = true;

                    // Tenta obter a data de validade com segurança
                    DateTime? validade = null;
                    try
                    {
                        // Tenta obter o valor como string
                        string valorString = dados.Valor(colunaValidade) as string;

                        // Verifica se é nulo ou vazio
                        if (!string.IsNullOrEmpty(valorString))
                        {
                            // Tenta converter para DateTime usando TryParse
                            DateTime dataValidade;
                            if (DateTime.TryParse(valorString, out dataValidade))
                            {
                                validade = dataValidade;
                                Console.WriteLine($"Data válida encontrada para {colunaNome}: {validade}");
                            }
                        }

                        // Verificar também se já existe uma data de validade na tabela
                        string queryValidade = $"SELECT {colunaValidade} FROM Geral_Entidade WHERE Id = '{_idSelecionado}'";
                        var dadosValidade = _BSO.Consulta(queryValidade);

                        if (dadosValidade != null && dadosValidade.NumLinhas() > 0)
                        {
                            dadosValidade.Inicio();

                            try
                            {
                                // Tentar obter o valor diretamente como DateTime
                                object valorObj = dadosValidade.Valor(colunaValidade);

                                if (valorObj != null && valorObj != DBNull.Value)
                                {
                                    // Tentar converter o valor para DateTime de várias maneiras
                                    if (valorObj is DateTime dataValor)
                                    {
                                        validade = dataValor;
                                        Console.WriteLine($"Data encontrada como DateTime para {colunaNome}: {validade}");
                                    }
                                    else if (DateTime.TryParse(valorObj.ToString(), out DateTime dataParsed))
                                    {
                                        validade = dataParsed;
                                        Console.WriteLine($"Data convertida de string para {colunaNome}: {validade}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Erro ao converter data do banco: {ex.Message}");

                                // Tentar novamente com outro método se o anterior falhar
                                try
                                {
                                    string valorString2 = dadosValidade.Valor(colunaValidade)?.ToString();
                                    if (!string.IsNullOrEmpty(valorString2))
                                    {
                                        DateTime dataConvertida;
                                        if (DateTime.TryParse(valorString2, out dataConvertida))
                                        {
                                            validade = dataConvertida;
                                            Console.WriteLine($"Data recuperada com método alternativo para {colunaNome}: {validade}");
                                        }
                                    }
                                }
                                catch (Exception ex2)
                                {
                                    Console.WriteLine($"Segundo erro ao converter data: {ex2.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao obter validade: {ex.Message}");
                        validade = null;
                    }

                    // SEMPRE mostrar com formato completo, mesmo que não tenha validade
                    if (validade.HasValue)
                    {
                        checkBox.Text = $"{tipoDoc} (Válido até: {validade.Value.ToShortDateString()})";
                        Console.WriteLine($"Checkbox {tipoDoc} atualizado com data: {validade.Value.ToShortDateString()}");
                    }
                    else
                    {
                        // Tenta obter a data diretamente da tabela como uma última tentativa
                        try
                        {
                            string queryValidade = $"SELECT {colunaValidade} FROM Geral_Entidade WHERE Id = '{_idSelecionado}'";
                            var dadosValidade = _BSO.Consulta(queryValidade);

                            if (dadosValidade != null && dadosValidade.NumLinhas() > 0)
                            {
                                dadosValidade.Inicio();
                                var dataDB = dadosValidade.Valor(colunaValidade);

                                if (dataDB != null && dataDB != DBNull.Value)
                                {
                                    DateTime dataParsed;
                                    if (dataDB is DateTime dt)
                                    {
                                        checkBox.Text = $"{tipoDoc} (Válido até: {dt.ToShortDateString()})";
                                        Console.WriteLine($"Checkbox {tipoDoc} atualizado com data direta: {dt.ToShortDateString()}");
                                    }
                                    else if (DateTime.TryParse(dataDB.ToString(), out dataParsed))
                                    {
                                        checkBox.Text = $"{tipoDoc} (Válido até: {dataParsed.ToShortDateString()})";
                                        Console.WriteLine($"Checkbox {tipoDoc} atualizado com data convertida: {dataParsed.ToShortDateString()}");
                                    }
                                    else
                                    {
                                        checkBox.Text = $"{tipoDoc} (Válido até: não definida)";
                                        Console.WriteLine($"Checkbox {tipoDoc} sem data de validade definida - não foi possível converter");
                                    }
                                }
                                else
                                {
                                    checkBox.Text = $"{tipoDoc} (Válido até: não definida)";
                                    Console.WriteLine($"Checkbox {tipoDoc} sem data de validade definida - valor nulo no banco");
                                }
                            }
                            else
                            {
                                checkBox.Text = $"{tipoDoc} (Válido até: não definida)";
                                Console.WriteLine($"Checkbox {tipoDoc} sem data de validade definida - sem dados no banco");
                            }
                        }
                        catch (Exception ex)
                        {
                            checkBox.Text = $"{tipoDoc} (Válido até: não definida)";
                            Console.WriteLine($"Erro ao tentar obter data para {tipoDoc}: {ex.Message}");
                        }
                    }
                    checkBox.AutoSize = true;
                }
                else
                {
                    checkBox.Text = tipoDoc;
                    checkBox.Checked = false;
                }
            }
            catch (Exception ex)
            {
                // Log do erro sem interromper o processo
                System.Diagnostics.Debug.WriteLine($"Erro ao atualizar checkbox {colunaNome}: {ex.Message}");
            }
        }

        private void SetInfoEntidades(Dictionary<string, string> entidade)
        {
            TXT_Codigo.Text = entidade["Codigo"];
            TXT_Nome.Text = entidade["Nome"];
            TXT_nome2.Text = entidade["Nome"];
            TXT_Contribuinte.Text = entidade["NIPC"];

            var moradaCompleta = $"{entidade["Morada"]}, {entidade["Localidade"]}, {entidade["CodPostal"]}, {entidade["CodPostalLocal"]}";

            if (moradaCompleta == ", , , ")
            {
                moradaCompleta = "";
            }
            else
            {
                TXT_Sede.Text = moradaCompleta;
            }
        }
        private void GetEntidadesID(ref Dictionary<string, string> entidade)
        {
            // Consulta SQL para pegar os dados
            var query = $@"SELECT * FROM Geral_Entidade WHERE CDU_TrataSGS = 0 AND Id='{_idSelecionado}'";
            var dados = _BSO.Consulta(query);

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

        private void btnSelecionarPasta_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Selecione a pasta para os documentos";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtCaminhoPasta.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void InitializeButtonEvents()
        {
            // Associar eventos de click aos botões
            button1.Click += (sender, e) => AnexarDocumento("Financas");
            button2.Click += (sender, e) => AnexarDocumento("SegSocial");
            button3.Click += (sender, e) => AnexarDocumento("FolhaPagamento");
            button4.Click += (sender, e) => AnexarDocumento("ComprovativoPagamento");
            button5.Click += (sender, e) => AnexarDocumento("ReciboSeguroAT");
            button6.Click += (sender, e) => AnexarDocumento("SeguroRC");
            button7.Click += (sender, e) => AnexarDocumento("HorarioTrabalho");
            button8.Click += (sender, e) => AnexarDocumento("SeguroAT");
            button9.Click += (sender, e) => AnexarDocumento("Alvara");
            button10.Click += (sender, e) => AnexarDocumento("CertidaoPermanente");
            button11.Click += (sender, e) => AnexarDocumento("Contrato");
            button12.Click += (sender, e) => AnexarDocumento("DeclaracaoPSS");
            button13.Click += (sender, e) => AnexarDocumento("ResponsavelEstaleiro");
        }

        private void AnexarDocumento(string tipoDocumento)
        {
            try
            {
                // Limpar as mensagens de console anteriores
                Console.WriteLine($"==== Anexando documento {tipoDocumento} ====");

                // Verifica se o caminho da pasta foi definido
                if (string.IsNullOrEmpty(txtCaminhoPasta.Text) || !System.IO.Directory.Exists(txtCaminhoPasta.Text))
                {
                    MessageBox.Show("Por favor, selecione uma pasta válida para os anexos primeiro.",
                        "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Solicitar data de validade
                DateTime dataValidade;
                using (Form formValidade = new Form())
                {
                    formValidade.Text = "Data de Validade";
                    formValidade.StartPosition = FormStartPosition.CenterParent;
                    formValidade.Width = 320;
                    formValidade.Height = 150;
                    formValidade.FormBorderStyle = FormBorderStyle.FixedDialog;
                    formValidade.MaximizeBox = false;
                    formValidade.MinimizeBox = false;

                    Label lblInfo = new Label();
                    lblInfo.Text = "Informe a data de validade do documento:";
                    lblInfo.Left = 20;
                    lblInfo.Top = 20;
                    lblInfo.Width = 250;

                    DateTimePicker dtpValidade = new DateTimePicker();
                    dtpValidade.Left = 20;
                    dtpValidade.Top = 50;
                    dtpValidade.Width = 250;
                    dtpValidade.Format = DateTimePickerFormat.Short;
                    dtpValidade.Value = DateTime.Now.AddMonths(1); // Um mês à frente como padrão

                    Button btnOk = new Button();
                    btnOk.Text = "OK";
                    btnOk.DialogResult = DialogResult.OK;
                    btnOk.Left = 110;
                    btnOk.Top = 80;

                    formValidade.Controls.Add(lblInfo);
                    formValidade.Controls.Add(dtpValidade);
                    formValidade.Controls.Add(btnOk);
                    formValidade.AcceptButton = btnOk;

                    if (formValidade.ShowDialog() != DialogResult.OK)
                    {
                        return; // Usuário cancelou
                    }

                    dataValidade = dtpValidade.Value;
                }

                // Abre o diálogo para selecionar o arquivo
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Title = $"Selecionar {tipoDocumento}";
                    openFileDialog.Filter = "Todos os arquivos (*.*)|*.*|Documentos PDF (*.pdf)|*.pdf|Documentos Word (*.doc;*.docx)|*.doc;*.docx|Imagens (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string sourceFile = openFileDialog.FileName;
                        string nomeArquivo = string.IsNullOrEmpty(TXT_Nome.Text)
                            ? "Sem_Nome"
                            : TXT_Nome.Text.Replace(" ", "_");

                        string fileName = $"{tipoDocumento.Replace(" ", "_")}_{nomeArquivo}_{DateTime.Now.ToString("yyyyMMdd")}{System.IO.Path.GetExtension(sourceFile)}";
                        string destFile = System.IO.Path.Combine(txtCaminhoPasta.Text, fileName);

                        // Verificar se o arquivo já existe
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

                        // Atualizar o banco de dados ou alguma propriedade para indicar que o documento foi anexado
                        AtualizarStatusDocumento(tipoDocumento, destFile, dataValidade);

                        // Atualizar o checkbox correspondente
                        AtualizarCheckbox(tipoDocumento, System.IO.Path.GetFileName(sourceFile), dataValidade);

                        // Recarregar os dados para garantir exibição correta
                        CarregarStatusDocumentos();

                        MessageBox.Show($"Documento '{tipoDocumento}' anexado com sucesso!\nValidade: {dataValidade.ToShortDateString()}",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Método para atualizar o checkbox correspondente ao documento
        private void AtualizarCheckbox(string tipoDocumento, string nomeArquivo, DateTime dataValidade)
        {
            CheckBox checkBox = null;
            string nomeDocumento = "";

            // Identificar qual checkbox deve ser atualizado com base no tipo de documento
            switch (tipoDocumento)
            {
                case "Financas":
                    checkBox = checkBox1;
                    nomeDocumento = "Finanças";
                    break;
                case "SegSocial":
                    checkBox = checkBox2;
                    nomeDocumento = "Segurança Social";
                    break;
                case "FolhaPagamento":
                    checkBox = checkBox3;
                    nomeDocumento = "Folha Pagamento";
                    break;
                case "ComprovativoPagamento":
                    checkBox = checkBox4;
                    nomeDocumento = "Comprovativo Pagamento";
                    break;
                case "ReciboSeguroAT":
                    checkBox = checkBox5;
                    nomeDocumento = "Seguro AT";
                    break;
                case "SeguroRC":
                    checkBox = checkBox6;
                    nomeDocumento = "Seguro RC";
                    break;
                case "HorarioTrabalho":
                    checkBox = checkBox7;
                    nomeDocumento = "Horário Trabalho";
                    break;
                case "SeguroAT":
                    checkBox = checkBox8;
                    nomeDocumento = "Condições Seguro AT";
                    break;
                case "Alvara":
                    checkBox = checkBox9;
                    nomeDocumento = "Alvará";
                    break;
                case "CertidaoPermanente":
                    checkBox = checkBox10;
                    nomeDocumento = "Certidão Permanente";
                    break;
                case "Contrato":
                    checkBox = checkBox11;
                    nomeDocumento = "Contrato";
                    break;
                case "DeclaracaoPSS":
                    checkBox = checkBox12;
                    nomeDocumento = "Declaração PSS";
                    break;
                case "ResponsavelEstaleiro":
                    checkBox = checkBox13;
                    nomeDocumento = "Responsável Estaleiro";
                    break;
            }

            // Se encontrou o checkbox, atualiza seu estado e texto
            if (checkBox != null)
            {
                checkBox.Checked = true;
                checkBox.Enabled = true;
                checkBox.Text = $"{nomeDocumento} (Válido até: {dataValidade.ToShortDateString()})";

                // Ajustar a largura do checkbox para mostrar o texto completo
                checkBox.AutoSize = true;
            }
        }

        private void AtualizarStatusDocumento(string tipoDocumento, string caminho, DateTime dataValidade)
        {
            try
            {
                // Atualizar a tabela Geral_Entidade com o caminho do documento e sua validade
                string colunaCaminho = "CDU_Caminho";
                string colunaAnexo = $"CDU_Anexo{tipoDocumento}";
                string colunaValidade = $"CDU_Validade{tipoDocumento}";

                // Caso especial para FolhaPagamento -> FolhaPag
                if (tipoDocumento == "FolhaPagamento")
                {
                    colunaAnexo = "CDU_AnexoFolhaPag";
                    colunaValidade = "CDU_ValidadeFolhaPag";
                }

                // Primeiro verificar se as colunas existem, e se não, criá-las
                string queryVerificarColunaCaminho = $@"
                    IF NOT EXISTS (
                        SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = '{colunaCaminho}'
                    )
                    BEGIN
                        ALTER TABLE Geral_Entidade ADD {colunaCaminho} NVARCHAR(500) NULL
                    END";
                _BSO.DSO.ExecuteSQL(queryVerificarColunaCaminho);

                string queryVerificarColunaAnexo = $@"
                    IF NOT EXISTS (
                        SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = '{colunaAnexo}'
                    )
                    BEGIN
                        ALTER TABLE Geral_Entidade ADD {colunaAnexo} INT DEFAULT 0
                    END";
                _BSO.DSO.ExecuteSQL(queryVerificarColunaAnexo);

                string queryVerificarColunaValidade = $@"
                    IF NOT EXISTS (
                        SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = '{colunaValidade}'
                    )
                    BEGIN
                        ALTER TABLE Geral_Entidade ADD {colunaValidade} DATE NULL
                    END";
                _BSO.DSO.ExecuteSQL(queryVerificarColunaValidade);

                // Sanitizar o caminho do arquivo para evitar problemas com aspas
                string caminhoSanitizado = caminho.Replace("'", "''");

                // Agora, atualizar os dados
                string query = $@"UPDATE Geral_Entidade SET 
                                {colunaCaminho} = '{caminhoSanitizado}',
                                {colunaAnexo} = 1,
                                {colunaValidade} = '{dataValidade.ToString("yyyy-MM-dd")}'
                                WHERE Id = '{_idSelecionado}'";
                _BSO.DSO.ExecuteSQL(query);

                // Verificar se os dados foram atualizados corretamente
                string queryVerificar = $"SELECT {colunaValidade} FROM Geral_Entidade WHERE Id = '{_idSelecionado}'";
                var dadosVerificar = _BSO.Consulta(queryVerificar);

                if (dadosVerificar != null && dadosVerificar.NumLinhas() > 0)
                {
                    dadosVerificar.Inicio();
                    var valorData = dadosVerificar.Valor(colunaValidade);
                    Console.WriteLine($"Verificação após salvar: Valor de {colunaValidade} no banco = {valorData}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar status do documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool VerificarDocumentoAnexado(string tipoDocumento)
        {
            try
            {
                // Nome da coluna baseado no tipo do documento
                string coluna = $"CDU_Anexo{tipoDocumento.Replace(" ", "")}";

                // Verificar se a coluna existe, e se não, criá-la
                string queryVerificarColuna = $@"
                    IF NOT EXISTS (
                        SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                        WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = '{coluna}'
                    )
                    BEGIN
                        ALTER TABLE Geral_Entidade ADD {coluna} INT DEFAULT 0
                    END";
                _BSO.DSO.ExecuteSQL(queryVerificarColuna);

                // Consulta SQL para verificar se o documento está anexado
                string query = $@"SELECT {coluna} FROM Geral_Entidade WHERE Id = '{_idSelecionado}'";
                var dados = _BSO.Consulta(query);

                dados.Inicio();
                if (dados.NumLinhas() > 0)
                {
                    return dados.DaValor<int>(coluna) == 1;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao verificar documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // Método para abrir a pasta de anexos
        public void AbrirPastaAnexos()
        {
            // Verifica se o caminho da pasta foi definido
            if (string.IsNullOrEmpty(txtCaminhoPasta.Text) || !System.IO.Directory.Exists(txtCaminhoPasta.Text))
            {
                MessageBox.Show("Pasta de anexos não definida ou não existente.",
                    "Pasta não encontrada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Abre a pasta no explorador de arquivos
            System.Diagnostics.Process.Start("explorer.exe", txtCaminhoPasta.Text);
        }

        // Método para abrir um documento específico
        private void AbrirDocumento(string tipoDocumento)
        {
            try
            {
                // Consulta SQL para obter o caminho do documento
                string query = $@"SELECT CDU_Caminho FROM Geral_Entidade WHERE Id = '{_idSelecionado}'";
                var dados = _BSO.Consulta(query);

                dados.Inicio();
                if (dados.NumLinhas() > 0)
                {
                    string caminho = dados.DaValor<string>("CDU_Caminho");

                    if (!string.IsNullOrEmpty(caminho) && System.IO.File.Exists(caminho))
                    {
                        // Abre o documento com o programa padrão
                        System.Diagnostics.Process.Start(caminho);
                    }
                    else
                    {
                        MessageBox.Show("O documento não foi encontrado no caminho especificado.",
                            "Documento não encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Não foi possível encontrar informações do documento.",
                        "Informação não encontrada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void VerificarECriarColunas()
        {
            try
            {
                // Lista de colunas para verificar e criar se necessário
                Dictionary<string, string> colunas = new Dictionary<string, string>
                {
                    { "CDU_AnexoFinancas", "INT DEFAULT 0" },
                    { "CDU_ValidadeFinancas", "DATE NULL" },
                    { "CDU_AnexoSegSocial", "INT DEFAULT 0" },
                    { "CDU_ValidadeSegSocial", "DATE NULL" },
                    { "CDU_AnexoFolhaPag", "INT DEFAULT 0" },
                    { "CDU_ValidadeFolhaPag", "DATE NULL" },
                    { "CDU_AnexoComprovativoPagamento", "INT DEFAULT 0" },
                    { "CDU_ValidadeComprovativoPagamento", "DATE NULL" },
                    { "CDU_AnexoReciboSeguroAT", "INT DEFAULT 0" },
                    { "CDU_ValidadeReciboSeguroAT", "DATE NULL" },
                    { "CDU_AnexoSeguroRC", "INT DEFAULT 0" },
                    { "CDU_ValidadeSeguroRC", "DATE NULL" },
                    { "CDU_AnexoHorarioTrabalho", "INT DEFAULT 0" },
                    { "CDU_ValidadeHorarioTrabalho", "DATE NULL" },
                    { "CDU_AnexoSeguroAT", "INT DEFAULT 0" },
                    { "CDU_ValidadeSeguroAT", "DATE NULL" },
                    { "CDU_AnexoAlvara", "INT DEFAULT 0" },
                    { "CDU_ValidadeAlvara", "DATE NULL" },
                    { "CDU_AnexoCertidaoPermanente", "INT DEFAULT 0" },
                    { "CDU_ValidadeCertidaoPermanente", "DATE NULL" },
                    { "CDU_AnexoContrato", "INT DEFAULT 0" },
                    { "CDU_ValidadeContrato", "DATE NULL" },
                    { "CDU_AnexoDeclaracaoPSS", "INT DEFAULT 0" },
                    { "CDU_ValidadeDeclaracaoPSS", "DATE NULL" },
                    { "CDU_AnexoResponsavelEstaleiro", "INT DEFAULT 0" },
                    { "CDU_ValidadeResponsavelEstaleiro", "DATE NULL" },
                    { "CDU_Caminho", "NVARCHAR(500) NULL" }
                };

                // Verifica todas as colunas em batch para reduzir o número de consultas
                string listaColunasVerificar = string.Join(", ", colunas.Keys.Select(c => $"'{c}'"));
                string queryVerificarTodas = $@"
                    SELECT COLUMN_NAME 
                    FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE TABLE_NAME = 'Geral_Entidade' 
                    AND COLUMN_NAME IN ({listaColunasVerificar})";

                var colunasExistentes = new List<string>();
                var dadosVerificar = _BSO.Consulta(queryVerificarTodas);

                if (dadosVerificar != null && dadosVerificar.NumLinhas() > 0)
                {
                    dadosVerificar.Inicio();
                    for (int i = 0; i < dadosVerificar.NumLinhas(); i++)
                    {
                        colunasExistentes.Add(dadosVerificar.DaValor<string>("COLUMN_NAME"));
                        dadosVerificar.Seguinte();
                    }
                }

                // Para cada coluna não existente, criar
                foreach (var coluna in colunas)
                {
                    if (!colunasExistentes.Contains(coluna.Key))
                    {
                        try
                        {
                            string queryAdicionar = $@"
                                ALTER TABLE Geral_Entidade ADD {coluna.Key} {coluna.Value}";
                            _BSO.DSO.ExecuteSQL(queryAdicionar);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Erro ao adicionar coluna {coluna.Key}: {ex.Message}");
                            // Continua para as próximas colunas mesmo se uma falhar
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao verificar/criar colunas: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}