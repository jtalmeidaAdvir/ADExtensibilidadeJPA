using ErpBS100;
using StdBE100;
using StdPlatBS100;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ADExtensibilidadeJPA
{
    public partial class GestaoSubempreitada : Form
    {

        private string Edit = "0";
        private string EditEqui = "0";
        private string EditAut = "0";
        private string Caminhotrab = "";
        private string Caminhoequi = "";
        private string Caminhoauto = "";
        private string SerieEqui = "";
        private readonly ErpBS _BSO;
        private readonly StdBSInterfPub _PSO;
        private readonly string _idSelecionado;

        public string Obratexto { get; private set; }
        public string NovoCodigoSelecionado { get; private set; }
        public string LinkNuvem { get; set; }

        public GestaoSubempreitada(ErpBS BSO, StdBSInterfPub PSO, string idSelecionado)
        {
            InitializeComponent();
            _BSO = BSO;
            _PSO = PSO;
            _idSelecionado = idSelecionado;
            CarregarDados();

            InitializeButtonEvents();
            ObterObras();

            GetValoresAutorizarObras();
            _ = InicializarAsync();
        }
        private async Task InicializarAsync()
        {
            Task tarefaTrabalhadores = Task.Run(() => CarregarTrabalhadores());
            Task tarefaEquipamentos = Task.Run(() => CarregarEquipamentos());
            await Task.WhenAll(tarefaTrabalhadores, tarefaEquipamentos);
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
                    CDU_AnexoResponsavelEstaleiro, CDU_ValidadeResponsavelEstaleiro,CDU_AnexoSeguroResposabilidadeCivil,CDU_AnexoAnexoD,
                    CDU_ValidadeSeguroResposabilidadeCivil,ID
                    FROM Geral_Entidade WHERE id = '{_idSelecionado}'";

                var dados = _BSO.Consulta(query);
                if (dados.NumLinhas() > 0)
                {
                    dados.Inicio();

                    try
                    {
                        // Atualizar checkboxes com base nos valores do banco de dados
                        SeguroUpdateCheckboxFromDB(checkBox1, dados, "CDU_AnexoFinancas", "DND-Finanças", "CDU_ValidadeFinancas");
                        SeguroUpdateCheckboxFromDB(checkBox2, dados, "CDU_AnexoSegSocial", "DND-Segurança-Social", "CDU_ValidadeSegSocial");
                        SeguroUpdateCheckboxFromDB(checkBox3, dados, "CDU_AnexoFolhaPag", "Mapa de Rem. – SS", "CDU_ValidadeFolhaPag");
                        SeguroUpdateCheckboxFromDB(checkBox4, dados, "CDU_AnexoComprovativoPagamento", "TSU", "CDU_ValidadeComprovativoPagamento");
                        SeguroUpdateCheckboxFromDB(checkBox5, dados, "CDU_AnexoReciboSeguroAT", "Seguro AT", "CDU_ValidadeReciboSeguroAT");
                        SeguroUpdateCheckboxFromDB(checkBox6, dados, "CDU_AnexoSeguroRC", "Seguro RC", "CDU_ValidadeSeguroRC");
                        SeguroUpdateCheckboxFromDB(checkBox8, dados, "CDU_AnexoSeguroAT", "Condições Seguro AT", "CDU_ValidadeSeguroAT");
                        SeguroUpdateCheckboxFromDB(checkBox9, dados, "CDU_AnexoAlvara", "Alvará", "CDU_ValidadeAlvara");
                        SeguroUpdateCheckboxFromDB(checkBox10, dados, "CDU_AnexoCertidaoPermanente", "Certidão Permanente", "CDU_ValidadeCertidaoPermanente");
                        SeguroUpdateCheckboxFromDB(checkBox13, dados, "CDU_AnexoSeguroResposabilidadeCivil", "Condições Seguro RC", "CDU_ValidadeSeguroResposabilidadeCivil");
                        SeguroUpdateCheckboxFromDB(checkBox28, dados, "CDU_AnexoAnexoD", "AnexoD", "CDU_ValidadeAnexoD");
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
                    // Tenta obter o valor como objeto (pode ser bool ou int)
                    var valor = dados.Valor(colunaNome);

                    var teste = dados.DaValor<string>("ID");
                    //MessageBox.Show(teste);

                    // Verifica se o valor é do tipo 'bit' (normalmente um tipo booleano ou 1/0)
                    if (valor is bool valorBool)
                    {
                        // Converte booleano para int (1 para true e 0 para false)
                        anexado = valorBool ? 1 : 0;

                    }
                    else if (valor is int valorInt)
                    {
                        // Caso o valor já seja inteiro, usa ele diretamente
                        anexado = valorInt;

                    }
                    else if (valor is string valorByte)
                    {
                        // Caso o valor seja byte (também poderia ser 1 ou 0), converte
                        anexado = int.Parse("1");

                    }

                }
                catch (Exception ex)
                {
                    // Captura a exceção e exibe a mensagem de erro
                    MessageBox.Show($"Erro: {ex.Message}");
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
                                    }
                                    else if (DateTime.TryParse(valorObj.ToString(), out DateTime dataParsed))
                                    {
                                        validade = dataParsed;
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
                        validade = null;
                    }

                    bool dataExpirada = false;

                    // SEMPRE mostrar com formato completo, mesmo que não tenha validade
                    if (validade.HasValue)
                    {
                        // Verificar se a data está expirada
                        dataExpirada = validade.Value < DateTime.Today;

                        checkBox.Text = $"{tipoDoc} (Válido até: {validade.Value.ToShortDateString()})";

                        // Atualiza a cor do texto baseado na validade
                        if (dataExpirada)
                        {
                            checkBox.ForeColor = Color.Red;
                        }
                        else
                        {
                            checkBox.ForeColor = SystemColors.ControlText; // Cor de texto padrão
                        }
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
                                        dataExpirada = dt < DateTime.Today;
                                        checkBox.Text = $"{tipoDoc} (Válido até: {dt.ToShortDateString()})";
                                    }
                                    else if (DateTime.TryParse(dataDB.ToString(), out dataParsed))
                                    {
                                        dataExpirada = dataParsed < DateTime.Today;
                                        checkBox.Text = $"{tipoDoc} (Válido até: {dataParsed.ToShortDateString()})";
                                    }
                                    else
                                    {
                                        checkBox.Text = $"{tipoDoc}";
                                    }

                                    // Atualiza a cor do texto baseado na validade
                                    if (dataExpirada)
                                    {
                                        checkBox.ForeColor = Color.Red;
                                    }
                                    else
                                    {
                                        checkBox.ForeColor = SystemColors.ControlText; // Cor de texto padrão
                                    }
                                }
                                else
                                {
                                    checkBox.Text = $"{tipoDoc}";
                                    checkBox.ForeColor = SystemColors.ControlText;
                                }
                            }
                            else
                            {
                                checkBox.Text = $"{tipoDoc}";
                                checkBox.ForeColor = SystemColors.ControlText;
                            }
                        }
                        catch (Exception ex)
                        {
                            checkBox.Text = $"{tipoDoc}";
                            checkBox.ForeColor = SystemColors.ControlText;
                        }
                    }
                    checkBox.AutoSize = true;
                }
                else
                {
                    checkBox.Text = tipoDoc;
                    checkBox.Checked = false;
                    checkBox.ForeColor = SystemColors.ControlText; // Cor de texto padrão
                }
            }
            catch (Exception ex)
            {
                // Log do erro sem interromper o processo
                System.Diagnostics.Debug.WriteLine($"Erro ao atualizar checkbox {colunaNome}: {ex.Message}");
            }
        }
        private void CaminhoEmpresa(Dictionary<string, string> entidade)
        {
            // Verificar se a chave "CDU_Caminho" existe no dicionário, e se for nula, atribuir uma string vazia
            string caminhoBase = entidade.ContainsKey("CDU_Caminho") && entidade["CDU_Caminho"] != null
                                 ? entidade["CDU_Caminho"]
                                 : "";

            // Obter o nome do usuário logado no computador atual
            string usuarioAtual = Environment.UserName;

            // Verificar se o caminho contém "C:\Users\<nome do usuário>" e substituí-lo
            string caminhoBaseUsuario = @"C:\Users\";

            // Verificar se o caminho já contém a pasta de usuário, como "C:\Users\<nome>"
            if (!string.IsNullOrEmpty(caminhoBase) && caminhoBase.Contains(caminhoBaseUsuario))
            {
                // Encontrar a parte do caminho onde começa o nome do usuário
                int inicioUsuario = caminhoBase.IndexOf(caminhoBaseUsuario) + caminhoBaseUsuario.Length;
                int fimUsuario = caminhoBase.IndexOf('\\', inicioUsuario);

                // Extrair o nome do usuário do caminho
                string nomeUsuarioOriginal = caminhoBase.Substring(inicioUsuario, fimUsuario - inicioUsuario);

                // Substituir o nome do usuário antigo pelo nome do usuário atual
                string caminhoAjustado = caminhoBase.Replace($@"C:\Users\{nomeUsuarioOriginal}", $@"C:\Users\{usuarioAtual}");

                // Definir o caminho ajustado no campo de texto
                txtCaminhoPasta.Text = caminhoAjustado;
            }
            else
            {
                // Se o caminho não contém "C:\Users\<nome do usuário>", mantemos o caminho original ou uma string vazia
                txtCaminhoPasta.Text = caminhoBase;
            }
        }

        private void SetInfoEntidades(Dictionary<string, string> entidade)
        {
            TXT_Codigo.Text = entidade["Codigo"];
            TXT_Nome.Text = entidade["Nome"];
            TXT_nome2.Text = entidade["Nome"];
            TXT_Contribuinte.Text = entidade["NIPC"];
            LinkNuvem = entidade["CDU_LinkNuvem"];
            TXTOP_linknuvem.Text = LinkNuvem;
            CBOP_Enviado.Checked = entidade["CDU_EmailEnviado"]?.ToString().Trim().ToLower() == "true"
                       || entidade["CDU_EmailEnviado"]?.ToString().Trim() == "1";

            CBOP_SGS.Checked = entidade["CDU_TrataSGS"].ToString() == "True";
            string query = $@"
                    SELECT ec.Email, ge.CDU_Link
                    FROM Geral_Entidade ge
                    LEFT JOIN Geral_Entidade_Contactos ec ON CAST(ge.id AS uniqueidentifier) = ec.EntidadeID
                    WHERE ge.id = '{_idSelecionado}'";
            var dados = _BSO.Consulta(query);
            if (dados.NumLinhas() > 0)
            {
                dados.Inicio();
                string email = dados.DaValor<string>("Email");
                if (!string.IsNullOrEmpty(email))
                {
                    txt_email.Text = email;
                }
            }


            if (DateTime.TryParse(entidade["CDU_DataEnvio"]?.ToString(), out DateTime dataEnvio))
            {
                DTPOP_DataEnvio.Value = dataEnvio;
            }
            else
            {
                DTPOP_DataEnvio.Visible = false;
                datavalor.Visible = true;
            }
            check_AlertaCaducados.Checked = entidade["CDU_IgnoraAlerta"]?.ToString().Trim().ToLower() == "true"
                       || entidade["CDU_IgnoraAlerta"]?.ToString().Trim() == "1";

            CaminhoEmpresa(entidade);

            // CaminhoTrabalhadores(entidade); 

            //txtCaminhoPasta.Text = entidade["CDU_Caminho"];

            txt_caminhoequi.Text = entidade["CDU_CaminhoEqui"];




            txt_link.Text = entidade["CDU_Link"];
            if (entidade.ContainsKey("CDU_CaminhoTRab"))
            {
                txt_caminhotrab.Text = entidade["CDU_CaminhoTRab"]?.ToString() ?? "";
            }
            else
            {
                txt_caminhotrab.Text = ""; // Define como string vazia se a chave não existir
            }

            //txt_caminhotrab.Text = entidade["CDU_CaminhoTRab"]?.ToString() ?? "";

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

        private void CaminhoTrabalhadores(Dictionary<string, string> entidade)
        {
            // Carregar o caminho da base de dados
            string caminhoBase = entidade["CDU_CaminhoTRab"].ToString();

            // Obter o nome do usuário logado no computador atual
            string usuarioAtual = Environment.UserName;

            // Verificar se o caminho contém "C:\Users\<nome do usuário>" e substituí-lo
            string caminhoBaseUsuario = @"C:\Users\";

            // Verificar se o caminho já contém a pasta de usuário, como "C:\Users\<nome>"
            if (caminhoBase.Contains(caminhoBaseUsuario))
            {
                // Encontrar a parte do caminho onde começa o nome do usuário
                int inicioUsuario = caminhoBase.IndexOf(caminhoBaseUsuario) + caminhoBaseUsuario.Length;
                int fimUsuario = caminhoBase.IndexOf('\\', inicioUsuario);

                // Extrair o nome do usuário do caminho
                string nomeUsuarioOriginal = caminhoBase.Substring(inicioUsuario, fimUsuario - inicioUsuario);

                // Substituir o nome do usuário antigo pelo nome do usuário atual
                string caminhoAjustado = caminhoBase.Replace($@"C:\Users\{nomeUsuarioOriginal}", $@"C:\Users\{usuarioAtual}");

                // Definir o caminho ajustado no campo de texto
                txt_caminhotrab.Text = caminhoAjustado;
            }
            else
            {
                // Se o caminho não contém "C:\Users\<nome do usuário>", mantemos o caminho original
                txt_caminhotrab.Text = caminhoBase;
            }
        }

        private void GetEntidadesID(ref Dictionary<string, string> entidade)
        {
            // Consulta SQL para pegar os dados
            var query = $@"SELECT * FROM Geral_Entidade WHERE CDU_TrataSGS = 1 AND Id='{_idSelecionado}'";
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
                                      "CDU_AnexoDStatus", "CDU_DecTrabEmigrStatus", "CDU_InscricaoSSStatus","CDU_CaminhoTRab",
                                      "CDU_CaminhoEqui","CDU_Link", "CDU_LinkNuvem", "CDU_EmailEnviado", "CDU_TrataSGS", "CDU_DataEnvio",
                                        "CDU_IgnoraAlerta"};

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
                    var update = $@"UPDATE Geral_Entidade
                                set CDU_Caminho = '{txtCaminhoPasta.Text}'
                                WHERE ID = '{_idSelecionado}'";
                    _BSO.DSO.ExecuteSQL(update);

                }
            }
        }

        private void InitializeButtonEvents()
        {
            // Associar eventos de click aos botões
            button1.Click += (sender, e) => AnexarDocumento("DND-Finanças");
            button2.Click += (sender, e) => AnexarDocumento("DND-Segurança-Social");
            button3.Click += (sender, e) => AnexarDocumento("FolhaPagamento");
            button4.Click += (sender, e) => AnexarDocumento("ComprovativoPagamento");
            button5.Click += (sender, e) => AnexarDocumento("ReciboSeguroAT");
            button6.Click += (sender, e) => AnexarDocumento("SeguroRC");
            button8.Click += (sender, e) => AnexarDocumento("SeguroAT");
            button9.Click += (sender, e) => AnexarDocumento("Alvara");
            button10.Click += (sender, e) => AnexarDocumento("CertidaoPermanente");

            button12.Click += (sender, e) => AnexarDocumento("SeguroResposabilidadeCivil"); // NOVO
            button13.Click += (sender, e) => AnexarDocumento("AnexoD"); // NOVO



            button11.Click += (sender, e) => AnexarDocumentoAutorizar("Trabalhadores_Imigrantes"); // NOVO

            button14.Click += (sender, e) => AnexarDocumentoTrabalhador("CartaoCidadao");
            button15.Click += (sender, e) => AnexarDocumentoTrabalhador("FichaAptidao");
            button16.Click += (sender, e) => AnexarDocumentoTrabalhador("FormacaoProfissional");
            button17.Click += (sender, e) => AnexarDocumentoTrabalhador("Trabalhosespecializados");
            button18.Click += (sender, e) => AnexarDocumentoTrabalhador("FichaDistribuicao");
            button20.Click += (sender, e) => AnexarDocumentoEquipamento("ConformidadeCE");
            button21.Click += (sender, e) => AnexarDocumentoEquipamento("Decreto-Lei");
            button22.Click += (sender, e) => AnexarDocumentoEquipamento("RegistoManutencao");
            button23.Click += (sender, e) => AnexarDocumentoEquipamento("ManualInstrucoes");
            button24.Click += (sender, e) => AnexarDocumentoEquipamento("seguro");
            button35.Click += (sender, e) => AnexarDocumentoAutorizar("contrato");
            button31.Click += (sender, e) => AnexarDocumentoAutorizar("Horario");
            button32.Click += (sender, e) => AnexarDocumentoAutorizar("Declaracao_PSS");
            button7.Click += (sender, e) => AnexarDocumentoAutorizar("Declaracao_Estaleiro");

        }
        private void AnexarDocumentoAutorizar(string tipoDocumento)
        {


            try
            {




                // Verifica se o caminho da pasta foi definido
                if (string.IsNullOrEmpty(txtcaminhoAuto.Text) || !System.IO.Directory.Exists(txtcaminhoAuto.Text))
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
                            : txt_serie.Text.Replace(" ", "_");

                        string nometrab = txt_marca.Text.Replace(" ", "_");

                        string nomeEmpresa = string.IsNullOrEmpty(TXT_Nome.Text) ? "Sem_Nome" : TXT_Nome.Text.Replace(" ", "_");
                        string nomeObra = cb_obras.Text.Replace(" ", "_");

                        // Create company folder if it doesn't exist
                        string companyFolder = System.IO.Path.Combine(txtcaminhoAuto.Text, nomeEmpresa);
                        if (!System.IO.Directory.Exists(companyFolder))
                        {
                            System.IO.Directory.CreateDirectory(companyFolder);
                        }

                        // Create Autorizacoes subfolder
                        string autorizacoesFolder = System.IO.Path.Combine(companyFolder, "Autorizacoes");
                        if (!System.IO.Directory.Exists(autorizacoesFolder))
                        {
                            System.IO.Directory.CreateDirectory(autorizacoesFolder);
                        }

                        // Create work specific folder
                        string workFolder = System.IO.Path.Combine(autorizacoesFolder, nomeObra);
                        if (!System.IO.Directory.Exists(workFolder))
                        {
                            System.IO.Directory.CreateDirectory(workFolder);
                        }

                        string fileName = $"{tipoDocumento.Replace(" ", "_")}_{DateTime.Now.ToString("yyyyMMdd")}{System.IO.Path.GetExtension(sourceFile)}";
                        string destFile = System.IO.Path.Combine(workFolder, fileName);
                        Caminhoauto = destFile;
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
                        // AtualizarStatusDocumentotrabalhdor(tipoDocumento, destFile, dataValidade);

                        // Atualizar o checkbox correspondente
                        AtualizarCheckboxautorizacoes(tipoDocumento, System.IO.Path.GetFileName(sourceFile), dataValidade);

                        // Recarregar os dados para garantir exibição correta
                        // CarregarStatusDocumentos();

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
        private void AnexarDocumentoEquipamento(string tipoDocumento)
        {
            try
            {



                // Verifica se o caminho da pasta foi definido
                if (string.IsNullOrEmpty(txt_caminhoequi.Text) || !System.IO.Directory.Exists(txt_caminhoequi.Text))
                {
                    MessageBox.Show("Por favor, selecione uma pasta válida para os anexos primeiro.",
                        "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Solicitar data de validade
                DateTime dataValidade;

                if (tipoDocumento == "ConformidadeCE" || tipoDocumento == "Decreto-Lei" || tipoDocumento == "RegistoManutencao" || tipoDocumento == "ManualInstrucoes")
                {
                    dataValidade = DateTime.Today;
                }
                else
                {
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
                            : txt_serie.Text.Replace(" ", "_");

                        string nometrab = txt_marca.Text.Replace(" ", "_");

                        string nomeEmpresa = string.IsNullOrEmpty(TXT_Nome.Text) ? "Sem_Nome" : TXT_Nome.Text.Replace(" ", "_");
                        string nomeEquipamento = txt_marca.Text.Replace(" ", "_");

                        // Create company folder if it doesn't exist
                        string companyFolder = System.IO.Path.Combine(txt_caminhoequi.Text, nomeEmpresa);
                        if (!System.IO.Directory.Exists(companyFolder))
                        {
                            System.IO.Directory.CreateDirectory(companyFolder);
                        }

                        // Create Equipamentos subfolder
                        string equipamentosFolder = System.IO.Path.Combine(companyFolder, "Equipamentos");
                        if (!System.IO.Directory.Exists(equipamentosFolder))
                        {
                            System.IO.Directory.CreateDirectory(equipamentosFolder);
                        }

                        // Create equipment specific folder
                        string equipmentFolder = System.IO.Path.Combine(equipamentosFolder, nomeEquipamento);
                        if (!System.IO.Directory.Exists(equipmentFolder))
                        {
                            System.IO.Directory.CreateDirectory(equipmentFolder);
                        }

                        string fileName = $"{tipoDocumento.Replace(" ", "_")}_{DateTime.Now.ToString("yyyyMMdd")}{System.IO.Path.GetExtension(sourceFile)}";
                        string destFile = System.IO.Path.Combine(equipmentFolder, fileName);
                        Caminhoequi = destFile;
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
                        // AtualizarStatusDocumentotrabalhdor(tipoDocumento, destFile, dataValidade);

                        // Atualizar o checkbox correspondente
                        AtualizarCheckboxequipamento(tipoDocumento, System.IO.Path.GetFileName(sourceFile), dataValidade);

                        // Recarregar os dados para garantir exibição correta
                        // CarregarStatusDocumentos();
                        if (tipoDocumento == "ConformidadeCE" || tipoDocumento == "Decreto-Lei" || tipoDocumento == "RegistoManutencao" || tipoDocumento == "ManualInstrucoes")
                        {
                            MessageBox.Show($"Documento '{tipoDocumento}' anexado com sucesso!",
                   "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show($"Documento '{tipoDocumento}' anexado com sucesso!\nValidade: {dataValidade.ToShortDateString()}",
                   "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AnexarDocumentoTrabalhador(string tipoDocumento)
        {
            try
            {
                if (!string.IsNullOrEmpty(txt_contribuintetrab.Text))
                {



                    // Verifica se o caminho da pasta foi definido
                    if (string.IsNullOrEmpty(txt_caminhotrab.Text) || !System.IO.Directory.Exists(txt_caminhotrab.Text))
                    {
                        MessageBox.Show("Por favor, selecione uma pasta válida para os anexos primeiro.",
                            "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Solicitar data de validade e número de CC (se for Cartão de Cidadão)
                    DateTime dataValidade;
                    string numeroCC = "";

                    using (Form formValidade = new Form())
                    {
                        formValidade.Text = tipoDocumento == "CartaoCidadao" ? "Dados do Cartão de Cidadão" : "Data de Validade";
                        formValidade.StartPosition = FormStartPosition.CenterParent;
                        formValidade.Width = 320;
                        formValidade.Height = tipoDocumento == "CartaoCidadao" ? 230 : 170;
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

                        formValidade.Controls.Add(lblInfo);
                        formValidade.Controls.Add(dtpValidade);

                        TextBox txtNumeroCC = null;
                        if (tipoDocumento == "CartaoCidadao")
                        {
                            Label lblNumeroCC = new Label();
                            lblNumeroCC.Text = "Número do Cartão de Cidadão:";
                            lblNumeroCC.Left = 20;
                            lblNumeroCC.Top = 90;
                            lblNumeroCC.Width = 250;

                            txtNumeroCC = new TextBox();
                            txtNumeroCC.Left = 20;
                            txtNumeroCC.Top = 110;
                            txtNumeroCC.Width = 250;

                            formValidade.Controls.Add(lblNumeroCC);
                            formValidade.Controls.Add(txtNumeroCC);
                        }

                        Button btnOk = new Button();
                        btnOk.Text = "OK";
                        btnOk.DialogResult = DialogResult.OK;
                        btnOk.Left = 110;
                        btnOk.Top = tipoDocumento == "CartaoCidadao" ? 140 : 80;

                        formValidade.Controls.Add(btnOk);
                        formValidade.AcceptButton = btnOk;

                        if (formValidade.ShowDialog() != DialogResult.OK)
                        {
                            return; // Usuário cancelou
                        }

                        dataValidade = dtpValidade.Value;
                        if (tipoDocumento == "CartaoCidadao" && txtNumeroCC != null)
                        {
                            numeroCC = txtNumeroCC.Text;
                        }
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
                            string nomeEmpresa = string.IsNullOrEmpty(TXT_Nome.Text) ? "Sem_Nome" : TXT_Nome.Text.Replace(" ", "_");
                            string nomeTrabalhador = txt_nometrab.Text.Replace(" ", "_");

                            // Create company folder if it doesn't exist
                            string companyFolder = System.IO.Path.Combine(txt_caminhotrab.Text, nomeEmpresa);
                            if (!System.IO.Directory.Exists(companyFolder))
                            {
                                System.IO.Directory.CreateDirectory(companyFolder);
                            }

                            // Create Trabalhadores subfolder
                            string trabalhadoresFolder = System.IO.Path.Combine(companyFolder, "Trabalhadores");
                            if (!System.IO.Directory.Exists(trabalhadoresFolder))
                            {
                                System.IO.Directory.CreateDirectory(trabalhadoresFolder);
                            }

                            // Create worker specific folder
                            string workerFolder = System.IO.Path.Combine(trabalhadoresFolder, nomeTrabalhador);
                            if (!System.IO.Directory.Exists(workerFolder))
                            {
                                System.IO.Directory.CreateDirectory(workerFolder);
                            }

                            string fileName = $"{tipoDocumento.Replace(" ", "_")}_{DateTime.Now.ToString("yyyyMMdd")}{System.IO.Path.GetExtension(sourceFile)}";
                            string destFile = System.IO.Path.Combine(workerFolder, fileName);
                            Caminhotrab = destFile;
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

                            // Se for Cartão de Cidadão, atualizar o número de CC na base de dados
                            if (tipoDocumento == "CartaoCidadao" && !string.IsNullOrEmpty(numeroCC))
                            {
                                // Verificar e criar a coluna Nm_CC se não existir
                                string checkColumnQuery = @"
                                    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                                                   WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'Nm_CC')
                                    BEGIN
                                        ALTER TABLE TDU_AD_Trabalhadores ADD Nm_CC NVARCHAR(50) NULL
                                    END";
                                _BSO.DSO.ExecuteSQL(checkColumnQuery);

                                // Atualizar o número de CC
                                string updateNumCCQuery = $@"
                                    UPDATE TDU_AD_Trabalhadores 
                                    SET Nm_CC = '{numeroCC}'
                                    WHERE contribuinte = '{txt_contribuintetrab.Text}' 
                                    AND id_empresa = '{_idSelecionado}'";
                                _BSO.DSO.ExecuteSQL(updateNumCCQuery);
                            }

                            // Atualizar o checkbox correspondente
                            AtualizarCheckboxtrabalhador(tipoDocumento, System.IO.Path.GetFileName(sourceFile), dataValidade);

                            MessageBox.Show($"Documento '{tipoDocumento}' anexado com sucesso!\nValidade: {dataValidade.ToShortDateString()}",
                                "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("O campo 'Contribuinte' não pode estar vazio.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao anexar documento: {ex.Message}",
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AtualizarCheckboxautorizacoes(string tipoDocumento, string nomeArquivo, DateTime dataValidade)
        {
            CheckBox checkBox = null;
            string nomeDocumento = "";
            // Identificar qual checkbox deve ser atualizado com base no tipo de documento
            switch (tipoDocumento)
            {
                case "contrato":
                    checkBox = checkBox27;
                    nomeDocumento = "contrato";
                    break;
                case "Horario":
                    checkBox = checkBox25;
                    nomeDocumento = "Horario";
                    break;
                case "Declaracao_PSS":
                    checkBox = checkBox26;
                    nomeDocumento = "Declaracao_PSS";
                    break;
                case "Declaracao_Estaleiro":
                    checkBox = checkBox7;
                    nomeDocumento = "Declaracao_Estaleiro";
                    break;

                case "Trabalhadores_Imigrantes":
                    checkBox = checkBox12;
                    nomeDocumento = "Trabalhadores_Imigrantes";
                    break;
            }

            // Se encontrou o checkbox, atualiza seu estado e texto
            if (checkBox != null)
            {
                checkBox.Enabled = true;
                checkBox.Checked = true;
                checkBox.Text = $"{nomeDocumento} (Válido até: {dataValidade.ToShortDateString()})";

                // Verificar se a data está expirada
                bool dataExpirada = dataValidade < DateTime.Today;

                // Atualizar a cor do texto baseado na validade
                if (dataExpirada)
                {
                    checkBox.ForeColor = Color.Red;
                }
                else
                {
                    checkBox.ForeColor = SystemColors.ControlText; // Cor de texto padrão
                }

                // Ajustar a largura do checkbox para mostrar o texto completo
                checkBox.AutoSize = true;
            }
        }
        private void AtualizarCheckboxequipamento(string tipoDocumento, string nomeArquivo, DateTime dataValidade)
        {
            CheckBox checkBox = null;
            string nomeDocumento = "";
            // Identificar qual checkbox deve ser atualizado com base no tipo de documento
            switch (tipoDocumento)
            {
                case "ConformidadeCE":
                    checkBox = checkBox19;
                    nomeDocumento = "ConformidadeCE";
                    break;
                case "Decreto-Lei":
                    checkBox = checkBox20;
                    nomeDocumento = "Decreto-Lei";
                    break;
                case "RegistoManutencao":
                    checkBox = checkBox21;
                    nomeDocumento = "RegistoManutencao";
                    break;
                case "ManualInstrucoes":
                    checkBox = checkBox22;
                    nomeDocumento = "ManualInstrucoes";
                    break;
                case "seguro":
                    checkBox = checkBox23;
                    nomeDocumento = "seguro";
                    break;
            }


            if (tipoDocumento == "ConformidadeCE" || tipoDocumento == "Decreto-Lei" || tipoDocumento == "RegistoManutencao" || tipoDocumento == "ManualInstrucoes")
            {
                checkBox.Enabled = true;
                checkBox.Checked = true;
                checkBox.Text = $"{nomeDocumento}";
            }
            else
            {
                if (checkBox != null)
                {
                    checkBox.Enabled = true;
                    checkBox.Checked = true;
                    checkBox.Text = $"{nomeDocumento} (Válido até: {dataValidade.ToShortDateString()})";

                    // Verificar se a data está expirada
                    bool dataExpirada = dataValidade < DateTime.Today;

                    // Atualizar a cor do texto baseado na validade
                    if (dataExpirada)
                    {
                        checkBox.ForeColor = Color.Red;
                    }
                    else
                    {
                        checkBox.ForeColor = SystemColors.ControlText; // Cor de texto padrão
                    }

                    // Ajustar a largura do checkbox para mostrar o texto completo
                    checkBox.AutoSize = true;
                }
            }

            // Se encontrou o checkbox, atualiza seu estado e texto

        }
        private void AtualizarCheckboxtrabalhador(string tipoDocumento, string nomeArquivo, DateTime dataValidade)
        {
            CheckBox checkBox = null;
            string nomeDocumento = "";
            // Identificar qual checkbox deve ser atualizado com base no tipo de documento
            switch (tipoDocumento)
            {
                case "CartaoCidadao":
                    checkBox = checkBox14;
                    nomeDocumento = "CartaoCidadao";
                    break;
                case "FichaAptidao":
                    checkBox = checkBox15;
                    nomeDocumento = "FichaAptidao";
                    break;
                case "FormacaoProfissional":
                    checkBox = checkBox16;
                    nomeDocumento = "FormacaoProfissional";
                    break;
                case "Trabalhosespecializados":
                    checkBox = checkBox17;
                    nomeDocumento = "Trabalhosespecializados";
                    break;
                case "FichaDistribuicao":
                    checkBox = checkBox18;
                    nomeDocumento = "FichaDistribuicao";
                    break;
            }

            // Se encontrou o checkbox, atualiza seu estado e texto
            if (checkBox != null)
            {
                checkBox.Enabled = true;
                checkBox.Checked = true;
                checkBox.Text = $"{nomeDocumento} (Válido até: {dataValidade.ToShortDateString()})";

                // Verificar se a data está expirada
                bool dataExpirada = dataValidade < DateTime.Today;

                // Atualizar a cor do texto baseado na validade
                if (dataExpirada)
                {
                    checkBox.ForeColor = Color.Red;
                }
                else
                {
                    checkBox.ForeColor = SystemColors.ControlText; // Cor de texto padrão
                }

                // Ajustar a largura do checkbox para mostrar o texto completo
                checkBox.AutoSize = true;
            }
        }


        private void AnexarDocumento(string tipoDocumento)
        {
            try
            {

                // Verifica se o caminho da pasta foi definido
                if (string.IsNullOrEmpty(txtCaminhoPasta.Text) || !System.IO.Directory.Exists(txtCaminhoPasta.Text))
                {
                    MessageBox.Show("Por favor, selecione uma pasta válida para os anexos primeiro.",
                        "Pasta não definida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                DateTime dataValidade;
                string numeroApoliceAt = "";
                string numeroApoliceRc = "";
                if (tipoDocumento == "AnexoD")
                {
                    dataValidade = DateTime.Today;
                }
                else if (tipoDocumento == "SeguroResposabilidadeCivil")
                {
                    dataValidade = DateTime.Today;
                    using (Form formApolice = new Form())
                    {
                        formApolice.Text = "Número da Apólice RC";
                        formApolice.StartPosition = FormStartPosition.CenterParent;
                        formApolice.Width = 320;
                        formApolice.Height = 150;
                        formApolice.FormBorderStyle = FormBorderStyle.FixedDialog;
                        formApolice.MaximizeBox = false;
                        formApolice.MinimizeBox = false;

                        Label lblNumeroApolice = new Label();
                        lblNumeroApolice.Text = "Número da Apólice RC:";
                        lblNumeroApolice.Left = 20;
                        lblNumeroApolice.Top = 20;
                        lblNumeroApolice.Width = 250;

                        TextBox txtNumeroApolice = new TextBox();
                        txtNumeroApolice.Left = 20;
                        txtNumeroApolice.Top = 50;
                        txtNumeroApolice.Width = 250;

                        Button btnOk = new Button();
                        btnOk.Text = "OK";
                        btnOk.DialogResult = DialogResult.OK;
                        btnOk.Left = 110;
                        btnOk.Top = 80;

                        formApolice.Controls.Add(lblNumeroApolice);
                        formApolice.Controls.Add(txtNumeroApolice);
                        formApolice.Controls.Add(btnOk);
                        formApolice.AcceptButton = btnOk;

                        if (formApolice.ShowDialog() != DialogResult.OK)
                        {
                            return; // Usuário cancelou
                        }

                        numeroApoliceRc = txtNumeroApolice.Text;
                    }
                }
                else if (tipoDocumento == "SeguroAT")
                {
                    dataValidade = DateTime.Today;
                    using (Form formApolice = new Form())
                    {
                        formApolice.Text = "Número da Apólice AT";
                        formApolice.StartPosition = FormStartPosition.CenterParent;
                        formApolice.Width = 320;
                        formApolice.Height = 150;
                        formApolice.FormBorderStyle = FormBorderStyle.FixedDialog;
                        formApolice.MaximizeBox = false;
                        formApolice.MinimizeBox = false;

                        Label lblNumeroApolice = new Label();
                        lblNumeroApolice.Text = "Número da Apólice AT:";
                        lblNumeroApolice.Left = 20;
                        lblNumeroApolice.Top = 20;
                        lblNumeroApolice.Width = 250;

                        TextBox txtNumeroApolice = new TextBox();
                        txtNumeroApolice.Left = 20;
                        txtNumeroApolice.Top = 50;
                        txtNumeroApolice.Width = 250;

                        Button btnOk = new Button();
                        btnOk.Text = "OK";
                        btnOk.DialogResult = DialogResult.OK;
                        btnOk.Left = 110;
                        btnOk.Top = 80;

                        formApolice.Controls.Add(lblNumeroApolice);
                        formApolice.Controls.Add(txtNumeroApolice);
                        formApolice.Controls.Add(btnOk);
                        formApolice.AcceptButton = btnOk;

                        if (formApolice.ShowDialog() != DialogResult.OK)
                        {
                            return; // Usuário cancelou
                        }

                        numeroApoliceAt = txtNumeroApolice.Text;
                    }
                }
                else
                {
                    using (Form formValidade = new Form())
                    {
                        formValidade.Text = "Data de Validade";
                        formValidade.StartPosition = FormStartPosition.CenterParent;
                        formValidade.Width = 320;
                        formValidade.Height = 170;
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

                        formValidade.Controls.Add(lblInfo);
                        formValidade.Controls.Add(dtpValidade);



                        Button btnOk = new Button();
                        btnOk.Text = "OK";
                        btnOk.DialogResult = DialogResult.OK;
                        btnOk.Left = 110;
                        btnOk.Top = 80;

                        formValidade.Controls.Add(btnOk);
                        formValidade.AcceptButton = btnOk;

                        if (formValidade.ShowDialog() != DialogResult.OK)
                        {
                            return; // Usuário cancelou
                        }

                        dataValidade = dtpValidade.Value;

                    }
                }
                // Solicitar data de validade


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

                        // Create main company folder
                        string companyFolder = System.IO.Path.Combine(txtCaminhoPasta.Text, nomeArquivo);
                        if (!System.IO.Directory.Exists(companyFolder))
                        {
                            System.IO.Directory.CreateDirectory(companyFolder);
                        }

                        // Create EMPRESA subfolder
                        string empresaFolder = System.IO.Path.Combine(companyFolder, "EMPRESA");
                        if (!System.IO.Directory.Exists(empresaFolder))
                        {
                            System.IO.Directory.CreateDirectory(empresaFolder);
                        }

                        string fileName = $"{tipoDocumento.Replace(" ", "_")}_{nomeArquivo}_{DateTime.Now.ToString("yyyyMMdd")}{System.IO.Path.GetExtension(sourceFile)}";
                        string destFile = System.IO.Path.Combine(empresaFolder, fileName);

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

                        // Se for Seguro Responsabilidade Civil, atualizar o número de Apólice RC na base de dados
                        if (tipoDocumento == "SeguroResposabilidadeCivil" && !string.IsNullOrEmpty(numeroApoliceRc))
                        {
                            // Verificar e criar a coluna CDU_NumApoliceRc se não existir
                            string checkColumnQuery = @"
                                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                                               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_NumApoliceRc')
                                BEGIN
                                    ALTER TABLE Geral_Entidade ADD CDU_NumApoliceRc NVARCHAR(50) NULL
                                END";
                            _BSO.DSO.ExecuteSQL(checkColumnQuery);

                            // Atualizar o número de Apólice RC
                            string updateNumApoliceQuery = $@"
                                UPDATE Geral_Entidade 
                                SET CDU_NumApoliceRc = '{numeroApoliceRc}'
                                WHERE ID = '{_idSelecionado}'";
                            _BSO.DSO.ExecuteSQL(updateNumApoliceQuery);
                        }

                        // Se for Seguro AT, atualizar o número de Apólice AT na base de dados
                        if (tipoDocumento == "SeguroAT" && !string.IsNullOrEmpty(numeroApoliceAt))
                        {
                            // Verificar e criar a coluna CDU_NumApoliceAt se não existir
                            string checkColumnQuery = @"
                                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                                               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_NumApoliceAt')
                                BEGIN
                                    ALTER TABLE Geral_Entidade ADD CDU_NumApoliceAt NVARCHAR(50) NULL
                                END";
                            _BSO.DSO.ExecuteSQL(checkColumnQuery);

                            // Atualizar o número de Apólice AT
                            string updateNumApoliceQuery = $@"
                                UPDATE Geral_Entidade 
                                SET CDU_NumApoliceAt = '{numeroApoliceAt}'
                                WHERE ID = '{_idSelecionado}'";
                            _BSO.DSO.ExecuteSQL(updateNumApoliceQuery);
                        }

                        // Atualizar o banco de dados ou alguma propriedade para indicar que o documento foi anexado
                        AtualizarStatusDocumento(tipoDocumento, destFile, dataValidade);

                        // Atualizar o checkbox correspondente
                        AtualizarCheckbox(tipoDocumento, System.IO.Path.GetFileName(sourceFile), dataValidade);

                        // Recarregar os dados para garantir exibição correta
                        // CarregarStatusDocumentos();

                        if (tipoDocumento == "SeguroAT" || tipoDocumento == "SeguroResposabilidadeCivil" || tipoDocumento == "AnexoD")
                        {
                            MessageBox.Show($"Documento '{tipoDocumento}' anexado com sucesso!",
                                "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show($"Documento '{tipoDocumento}' anexado com sucesso!\nValidade: {dataValidade.ToShortDateString()}",
                                "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
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
                case "DND-Finanças":
                    checkBox = checkBox1;
                    nomeDocumento = "DND-Finanças";
                    break;
                case "DND-Segurança-Social":
                    checkBox = checkBox2;
                    nomeDocumento = "DND-Segurança-Social";
                    break;
                case "Mapa de Rem. – SS":
                    checkBox = checkBox3;
                    nomeDocumento = "Mapa de Rem. – SS";
                    break;
                case "TSU":
                    checkBox = checkBox4;
                    nomeDocumento = "TSU";
                    break;
                case "ReciboSeguroAT":
                    checkBox = checkBox5;
                    nomeDocumento = "Seguro AT";
                    break;
                case "SeguroRC":
                    checkBox = checkBox6;
                    nomeDocumento = "Seguro RC";
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
                case "SeguroResposabilidadeCivil":
                    checkBox = checkBox13;
                    nomeDocumento = "Condições Seguro RC";
                    break;
                case "AnexoD":
                    checkBox = checkBox28;
                    nomeDocumento = "AnexoD";
                    break;
            }
            if (tipoDocumento == "SeguroAT" || tipoDocumento == "SeguroResposabilidadeCivil" || tipoDocumento == "AnexoD")
            {
                checkBox.Enabled = true;
                checkBox.Checked = true;
                checkBox.Text = $"{nomeDocumento}";
            }
            else
            {
                // Se encontrou o checkbox, atualiza seu estado e texto
                if (checkBox != null)
                {
                    checkBox.Enabled = true;
                    checkBox.Checked = true;
                    checkBox.Text = $"{nomeDocumento} (Válido até: {dataValidade.ToShortDateString()})";

                    // Verificar se a data está expirada
                    bool dataExpirada = dataValidade < DateTime.Today;

                    // Atualizar a cor do texto baseado na validade
                    if (dataExpirada)
                    {
                        checkBox.ForeColor = Color.Red;
                    }
                    else
                    {
                        checkBox.ForeColor = SystemColors.ControlText; // Cor de texto padrão
                    }

                    // Ajustar a largura do checkbox para mostrar o texto completo
                    checkBox.AutoSize = true;
                }
            }

        }

        private void AtualizarStatusDocumento(string tipoDocumento, string caminho, DateTime dataValidade)
        {
            try
            {
                // Atualizar a tabela Geral_Entidade com o caminho do documento e sua validade
                string colunaCaminho = "CDU_Caminho";
                string colunaAnexo;
                string colunaValidade;
                // Mapear nomes de documentos para nomes de colunas
                switch (tipoDocumento)
                {
                    case "DND-Finanças":
                        colunaAnexo = "CDU_AnexoFinancas";
                        colunaValidade = "CDU_ValidadeFinancas";
                        break;
                    case "DND-Segurança-Social":
                        colunaAnexo = "CDU_AnexoSegSocial";
                        colunaValidade = "CDU_ValidadeSegSocial";
                        break;
                    case "FolhaPagamento":
                        colunaAnexo = "CDU_AnexoFolhaPag";
                        colunaValidade = "CDU_ValidadeFolhaPag";
                        break;
                    case "TSU":
                        colunaAnexo = "CDU_AnexoComprovativoPagamento";
                        colunaValidade = "CDU_ValidadeComprovativoPagamento";
                        break;
                    case "ReciboSeguroAT":
                        colunaAnexo = "CDU_AnexoReciboSeguroAT";
                        colunaValidade = "CDU_ValidadeReciboSeguroAT";
                        break;
                    case "SeguroRC":
                        colunaAnexo = "CDU_AnexoSeguroRC";
                        colunaValidade = "CDU_ValidadeSeguroRC";
                        break;
                    case "HorarioTrabalho":
                        colunaAnexo = "CDU_AnexoHorarioTrabalho";
                        colunaValidade = "CDU_ValidadeHorarioTrabalho";
                        break;
                    case "SeguroAT":
                        colunaAnexo = "CDU_AnexoSeguroAT";
                        colunaValidade = "CDU_ValidadeSeguroAT";
                        break;
                    case "Alvara":
                        colunaAnexo = "CDU_AnexoAlvara";
                        colunaValidade = "CDU_ValidadeAlvara";
                        break;
                    case "CertidaoPermanente":
                        colunaAnexo = "CDU_AnexoCertidaoPermanente";
                        colunaValidade = "CDU_ValidadeCertidaoPermanente";
                        break;
                    case "Contrato":
                        colunaAnexo = "CDU_AnexoContrato";
                        colunaValidade = "CDU_ValidadeContrato";
                        break;
                    case "DeclaracaoPSS":
                        colunaAnexo = "CDU_AnexoDeclaracaoPSS";
                        colunaValidade = "CDU_ValidadeDeclaracaoPSS";
                        break;
                    case "ResponsavelEstaleiro":
                        colunaAnexo = "CDU_AnexoResponsavelEstaleiro";
                        colunaValidade = "CDU_ValidadeResponsavelEstaleiro";
                        break;
                    case "SeguroResposabilidadeCivil":
                        colunaAnexo = "CDU_AnexoSeguroResposabilidadeCivil";
                        colunaValidade = "CDU_ValidadeSeguroResposabilidadeCivil";
                        break;
                    case "AnexoD":
                        colunaAnexo = "CDU_AnexoAnexoD";
                        colunaValidade = "CDU_ValidadeAnexoD";
                        break;
                    default:
                        // Caso não mapeado, usar o nome do tipo como parte do nome da coluna
                        colunaAnexo = $"CDU_Anexo{tipoDocumento}";
                        colunaValidade = $"CDU_Validade{tipoDocumento}";
                        break;
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
                if (tipoDocumento == "SeguroAT" || tipoDocumento == "SeguroResposabilidadeCivil")
                {
                    string query2 = $@"UPDATE Geral_Entidade SET 
                                {colunaAnexo} = 1
                                WHERE Id = '{_idSelecionado}'";
                    _BSO.DSO.ExecuteSQL(query2);
                }
                else
                {
                    string query = $@"UPDATE Geral_Entidade SET 
                                {colunaAnexo} = 1,
                                {colunaValidade} = '{dataValidade.ToString("yyyy-MM-dd")}'
                                WHERE Id = '{_idSelecionado}'";
                    _BSO.DSO.ExecuteSQL(query);
                }
                // Agora, atualizar os dados


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

        private void bt_AbrirPasta_Click(object sender, EventArgs e)
        {
            string caminhoPasta = txtCaminhoPasta.Text;

            // Verificar se o caminho da pasta existe
            if (Directory.Exists(caminhoPasta))
            {
                // Abrir a pasta no explorador de arquivos
                Process.Start("explorer.exe", caminhoPasta);
            }
            else
            {
                MessageBox.Show("O caminho da pasta não é válido.");
            }
        }

        private void vt_adcionar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_contribuintetrab.Text))
            {
                // Ação quando o TextBox estiver vazio ou null
                MessageBox.Show("O campo 'Contribuinte' não pode estar vazio.");
            }
            else
            {
                if (Edit == "0")
                {

                    InsereTrabalhador();
                }
                else
                {
                    AtualizaTrabalhador();
                }
                LimpaCampos();
            }


        }


        private void LimpaCampos()
        {
            txt_nometrab.Text = "";
            txt_categoriatrab.Text = "";
            txt_contribuintetrab.Text = "";
            txt_segurancasocialtrab.Text = "";
            rxt_emailTrabalhador.Text = "";

            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;

            checkBox14.Text = "";
            checkBox15.Text = "";
            checkBox16.Text = "";
            checkBox17.Text = "";
            checkBox18.Text = "";

            cb_FormacaoProfissional.SelectedIndex = 0;
            cb_especializados.SelectedIndex = 0;
            txt_contribuintetrab.Enabled = true;
            dtpnascimento.Value = DateTime.Now;
            bt_remover.Visible = false;
            button28.Visible = false;
            Edit = "0";

        }



        // Função para tratar e "sanitizar" a string
        private string SanitizeString(string input)
        {
            // Substitui caracteres problemáticos que possam interferir na consulta SQL
            return input.Replace("'", "''").Replace(":", "&#58;").Replace("(", "&#40;").Replace(")", "&#41;");
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Verifica se o usuário clicou em uma linha válida
            if (e.RowIndex >= 0)
            {
                // Obtém a linha selecionada
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                // Preenche os campos de texto com os valores da linha
                txt_nometrab.Text = row.Cells["Nome"].Value.ToString();
                txt_categoriatrab.Text = row.Cells["Categoria"].Value.ToString();
                txt_contribuintetrab.Text = row.Cells["Contribuinte"].Value.ToString();

                txt_contribuintetrab.Enabled = false;
                txt_segurancasocialtrab.Text = row.Cells["SSocial"].Value.ToString();

                // Buscar email da base de dados
                string contribuinte = row.Cells["Contribuinte"].Value.ToString();
                string queryEmail = $@"SELECT email FROM TDU_AD_Trabalhadores 
                                      WHERE id_empresa = '{_idSelecionado}' AND contribuinte = '{contribuinte}'";
                var dadosEmail = _BSO.Consulta(queryEmail);
                if (dadosEmail.NumLinhas() > 0)
                {
                    dadosEmail.Inicio();
                    rxt_emailTrabalhador.Text = dadosEmail.DaValor<string>("email") ?? "";
                }
                else
                {
                    rxt_emailTrabalhador.Text = "";
                }

                checkBox14.Checked = ConvertToBool(row.Cells["AnexoCC"].Value);
                checkBox15.Checked = ConvertToBool(row.Cells["AnexoFM"].Value);
                checkBox16.Checked = ConvertToBool(row.Cells["AnexoCT"].Value);
                checkBox17.Checked = ConvertToBool(row.Cells["AnexoTE"].Value);
                checkBox18.Checked = ConvertToBool(row.Cells["AnexoEPI"].Value);

                VerificarEColorirCheckBox(checkBox14, row.Cells["caminho1"].Value);
                VerificarEColorirCheckBox(checkBox15, row.Cells["caminho2"].Value);
                VerificarEColorirCheckBox(checkBox16, row.Cells["caminho3"].Value);
                VerificarEColorirCheckBox(checkBox17, row.Cells["caminho4"].Value);
                VerificarEColorirCheckBox(checkBox18, row.Cells["caminho5"].Value);

                cb_FormacaoProfissional.Text = row.Cells["CBFormacaoProfissional"].Value.ToString();
                cb_especializados.Text = row.Cells["CBespecializados"].Value.ToString();

                var datanascimento = row.Cells["DataNasc"].Value.ToString();
                DateTime dataNasc;
                if (DateTime.TryParse(row.Cells["DataNasc"].Value.ToString(), out dataNasc))
                {
                    dtpnascimento.Enabled = true;
                    dtpnascimento.Visible = true;
                    dtpnascimento.CustomFormat = "dd/MM/yyyy";
                    dtpnascimento.Value = dataNasc;
                    checkBox11.Checked = true;
                }
                else
                {
                    checkBox11.Checked = false;
                    dtpnascimento.Enabled = false;
                    dtpnascimento.Visible = false;
                    dtpnascimento.CustomFormat = " "; // Deixa a data em branco
                    dtpnascimento.Value = new DateTime(1753, 1, 1);// ou qualquer 
                }


                Match match = Regex.Match(checkBox14.Text, @"\d{2}/\d{2}/\d{4}");


                bt_remover.Visible = true;
                button28.Visible = true;

                Edit = "1";

            }
        }

        private static void VerificarEColorirCheckBox(CheckBox checkBox, object cellValue)
        {
            if (cellValue == null) return; // Ignora se for null

            string text = cellValue.ToString();
            checkBox.Text = text; // Atribui o texto ao CheckBox

            Match match = Regex.Match(text, @"\d{2}/\d{2}/\d{4}");

            if (!match.Success) return; // Se não encontrar data, sai da função

            string dataStr = match.Value;
            DateTime dataExtraida = DateTime.ParseExact(dataStr, "dd/MM/yyyy", null);
            DateTime hoje = DateTime.Today;

            if (dataExtraida < hoje)
            {
                checkBox.ForeColor = System.Drawing.Color.Red; // Pinta de vermelho se a data for antiga
            }
            else
            {
                checkBox.ForeColor = System.Drawing.Color.Black;
            }
        }

        private bool ConvertToBool(object value)
        {
            if (value == null) return false; // Caso o valor seja null, retorna false

            // Se for um booleano, simplesmente retorna o valor
            if (value is bool)
            {
                return (bool)value;
            }

            // Se for uma string ("true" ou "false"), tenta converter
            if (value is string stringValue)
            {
                bool result;
                if (bool.TryParse(stringValue, out result))
                {
                    return result;
                }
            }

            // Se for um número inteiro (0 ou 1), considera 0 como false e 1 como true
            if (value is int intValue)
            {
                return intValue == 1; // 1 -> true, 0 -> false
            }

            // Se o valor não for de tipo esperado, retorna false
            return false;
        }



        private string RestoreSanitizedString(string input)
        {
            // Restaurar os caracteres escapados para os seus valores originais
            return input.Replace("&#40;", "(")
                        .Replace("&#41;", ")")
                        .Replace("&#58;", ":");
        }

        private void btnSelecionarPastaTrab_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Selecione a pasta para os documentos";
                folderDialog.ShowNewFolderButton = true;


                string checkColumnQuery = @"
    SELECT * 
    FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME = 'Geral_Entidade' 
    AND COLUMN_NAME = 'CDU_CaminhoTRab'";
                var columnExists = _BSO.Consulta(checkColumnQuery);


                if (columnExists.NumLinhas() > 0)
                {
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        txt_caminhotrab.Text = folderDialog.SelectedPath;
                        var update = $@"UPDATE Geral_Entidade
                                set CDU_CaminhoTRab = '{txt_caminhotrab.Text}'
                                WHERE ID = '{_idSelecionado}'";
                        _BSO.DSO.ExecuteSQL(update);

                    }
                }
                else
                {
                    // Cria a coluna se não existir
                    string alterTableQuery = @"
                    ALTER TABLE Geral_Entidade 
                    ADD CDU_CaminhoTRab NVARCHAR(500)"; // Ajuste o tipo de dado conforme necessário

                    _BSO.DSO.ExecuteSQL(alterTableQuery);
                }


            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string caminhoPasta = txt_caminhotrab.Text;

            // Verificar se o caminho da pasta existe
            if (Directory.Exists(caminhoPasta))
            {
                // Abrir a pasta no explorador de arquivos
                Process.Start("explorer.exe", caminhoPasta);
            }
            else
            {
                MessageBox.Show("O caminho da pasta não é válido.");
            }
        }

        private void btnSelecionarPastaEqui_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Selecione a pasta para os documentos";
                folderDialog.ShowNewFolderButton = true;


                string checkColumnQuery = @"
    SELECT * 
    FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME = 'Geral_Entidade' 
    AND COLUMN_NAME = 'CDU_CaminhoEqui'";
                var columnExists = _BSO.Consulta(checkColumnQuery);


                if (columnExists.NumLinhas() > 0)
                {
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        txt_caminhoequi.Text = folderDialog.SelectedPath;
                        var update = $@"UPDATE Geral_Entidade
                                set CDU_CaminhoEqui = '{txt_caminhoequi.Text}'
                                WHERE ID = '{_idSelecionado}'";
                        _BSO.DSO.ExecuteSQL(update);

                    }
                }
                else
                {
                    // Cria a coluna se não existir
                    string alterTableQuery = @"
                    ALTER TABLE Geral_Entidade 
                    ADD CDU_CaminhoEqui NVARCHAR(500)"; // Ajuste o tipo de dado conforme necessário

                    _BSO.DSO.ExecuteSQL(alterTableQuery);
                }


            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            string caminhoPasta = txt_caminhoequi.Text;

            // Verificar se o caminho da pasta existe
            if (Directory.Exists(caminhoPasta))
            {
                // Abrir a pasta no explorador de arquivos
                Process.Start("explorer.exe", caminhoPasta);
            }
            else
            {
                MessageBox.Show("O caminho da pasta não é válido.");
            }
        }

        private void bt_adcionarequi_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_serie.Text))
            {
                MessageBox.Show("O campo 'Série' é obrigatório. Por favor, preencha-o.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;  // Interrompe a execução se a série estiver vazia
            }

            if (EditEqui == "0")
            {

                InsereEquipamento();
            }
            else
            {
                AtualizaEquipamento();
            }

        }
        private void AtualizaTrabalhador()
        {
            // Obtém os dados a serem atualizados
            string nome = txt_nometrab.Text;
            string categoriatrab = txt_categoriatrab.Text;
            string contribuintetrab = txt_contribuintetrab.Text;
            string segurancasocialtrab = txt_segurancasocialtrab.Text;
            string emailTrab = rxt_emailTrabalhador.Text;
            int anexo1 = checkBox14.Checked ? 1 : 0;
            int anexo2 = checkBox15.Checked ? 1 : 0;
            int anexo3 = checkBox16.Checked ? 1 : 0;
            int anexo4 = checkBox17.Checked ? 1 : 0;
            int anexo5 = checkBox18.Checked ? 1 : 0;

            var cBFormacaoProfissional = cb_FormacaoProfissional.Text;
            var cBespecializados = cb_especializados.Text;

            string dtpnascimento2 = dtpnascimento.Value == DateTime.MinValue ? "NULL" : $"'{dtpnascimento.Value:yyyy-MM-dd HH:mm:ss}'";
            var datanasci = $"";

            if (dtpnascimento2 == "'1753-01-01 00:00:00'")
            {
                datanasci = $"";
            }
            else
            {
                datanasci = $"{dtpnascimento.Value:yyyy-MM-dd HH:mm:ss}";
            }
            // Encontre a linha selecionada no DataGridView para atualização, usando o 'contribuinte' como filtro
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["contribuinte"].Value != null && row.Cells["contribuinte"].Value.ToString() == contribuintetrab) // Verifica o contribuinte
                {
                    // Atualiza os valores na linha
                    row.Cells["nome"].Value = nome;
                    row.Cells["categoria"].Value = categoriatrab;
                    row.Cells["SSocial"].Value = segurancasocialtrab;

                    row.Cells["AnexoCC"].Value = anexo1;
                    row.Cells["AnexoFM"].Value = anexo2;
                    row.Cells["AnexoCT"].Value = anexo3;
                    row.Cells["AnexoTE"].Value = anexo4;
                    row.Cells["AnexoEPI"].Value = anexo5;

                    // Atualiza as labels de texto no DataGridView
                    row.Cells["caminho1"].Value = checkBox14.Text;
                    row.Cells["caminho2"].Value = checkBox15.Text;
                    row.Cells["caminho3"].Value = checkBox16.Text;
                    row.Cells["caminho4"].Value = checkBox17.Text;
                    row.Cells["caminho5"].Value = checkBox18.Text;

                    row.Cells["CBFormacaoProfissional"].Value = cBFormacaoProfissional;
                    row.Cells["CBespecializados"].Value = cBespecializados;

                    row.Cells["DataNasc"].Value = datanasci;

                    break; // Encontre e atualize a primeira linha correspondente
                }
            }
            string caminho1 = SanitizeString(checkBox14.Text);
            string caminho2 = SanitizeString(checkBox15.Text);
            string caminho3 = SanitizeString(checkBox16.Text);
            string caminho4 = SanitizeString(checkBox17.Text);
            string caminho5 = SanitizeString(checkBox18.Text);
            // Atualiza os dados na base de dados com o filtro no "contribuinte"
            string queryUpdate = $@"
        UPDATE TDU_AD_Trabalhadores
        SET nome = '{nome}',
            categoria = '{categoriatrab}', 
            contribuinte = '{contribuintetrab}', 
            seguranca_social = '{segurancasocialtrab}',
            email = '{emailTrab}', 
            anexo1 = {anexo1}, 
            anexo2 = {anexo2}, 
            anexo3 = {anexo3}, 
            anexo4 = {anexo4}, 
            anexo5 = {anexo5},
            caminho1 = '{caminho1}',
            caminho2 = '{caminho2}',
            caminho3 = '{caminho3}',
            caminho4 = '{caminho4}',
            caminho5 = '{caminho5}',

            cBFormacaoProfissional = '{cBFormacaoProfissional}', 
            cBespecializados = '{cBespecializados}', 

            data_nascimento = '{datanasci}'
        WHERE id_empresa = '{_idSelecionado}' AND contribuinte = '{contribuintetrab}';
    ";

            // Executa a query de atualização no banco de dados
            _BSO.DSO.ExecuteSQL(queryUpdate);

            // Mostra uma mensagem de confirmação
            MessageBox.Show("Trabalhador atualizado com sucesso.");

            // Limpa os campos de entrada após a atualização
            LimpaCampos();

            // Opcional: Retorna o foco para o primeiro campo
            txt_nometrab.Focus();
        }

        private void AtualizaEquipamento()
        {
            // Obtém os dados a serem atualizados
            string marca = txt_marca.Text;
            string tipo = txt_tipo.Text;
            string serie = txt_serie.Text;
            int anexo1 = checkBox19.Checked ? 1 : 0;
            int anexo2 = checkBox20.Checked ? 1 : 0;
            int anexo3 = checkBox21.Checked ? 1 : 0;
            int anexo4 = checkBox22.Checked ? 1 : 0;
            int anexo5 = checkBox23.Checked ? 1 : 0;
            var cBSeguro = cb_seguro.Text;
            var cBManualInstrucoes = cb_ManualInstrucoes.Text;
            var cBManutencao = cb_Manutencao.Text;
            var cBDecreto_Lei = cb_Decreto_Lei.Text;
            var cBConformidadeCE = cb_ConformidadeCE.Text;


            // Encontre a linha selecionada no DataGridView para atualização, usando o 'contribuinte' como filtro
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                var sss = row.Cells["Serieeq"].Value.ToString();

                if (sss == serie) // Verifica o contribuinte
                {

                    // Atualiza os valores na linha
                    row.Cells["marca"].Value = marca;
                    row.Cells["tipo"].Value = tipo;
                    row.Cells["Anexo1"].Value = anexo1;
                    row.Cells["Anexo2"].Value = anexo2;
                    row.Cells["Anexo3"].Value = anexo3;
                    row.Cells["Anexo4"].Value = anexo4;
                    row.Cells["Anexo5"].Value = anexo5;

                    // Atualiza as labels de texto no DataGridView
                    row.Cells["caminho6"].Value = checkBox19.Text;
                    row.Cells["caminho7"].Value = checkBox20.Text;
                    row.Cells["caminho8"].Value = checkBox21.Text;
                    row.Cells["caminho9"].Value = checkBox22.Text;
                    row.Cells["caminho10"].Value = checkBox23.Text;
                    row.Cells["CBSeguro"].Value = cBSeguro;
                    row.Cells["CBManualInstrucoes"].Value = cBManualInstrucoes;
                    row.Cells["CBManutencao"].Value = cBManutencao;
                    row.Cells["CBDecreto_Lei"].Value = cBDecreto_Lei;
                    row.Cells["CBConformidadeCE"].Value = cBConformidadeCE;

                    break; // Encontre e atualize a primeira linha correspondente
                }
            }
            string caminho1 = SanitizeString(checkBox19.Text);
            string caminho2 = SanitizeString(checkBox20.Text);
            string caminho3 = SanitizeString(checkBox21.Text);
            string caminho4 = SanitizeString(checkBox22.Text);
            string caminho5 = SanitizeString(checkBox23.Text);
            // Atualiza os dados na base de dados com o filtro no "contribuinte"
            string queryUpdate = $@"
        UPDATE TDU_AD_Equipamentos
        SET marca = '{marca}',
            tipo = '{tipo}', 
            anexo1 = {anexo1}, 
            anexo2 = {anexo2}, 
            anexo3 = {anexo3}, 
            anexo4 = {anexo4}, 
            anexo5 = {anexo5},
            caminho1 = '{caminho1}',
            caminho2 = '{caminho2}',
            caminho3 = '{caminho3}',
            caminho4 = '{caminho4}',
            caminho5 = '{caminho5}',
            cBSeguro = '{cBSeguro}',
            cBManualInstrucoes = '{cBManualInstrucoes}',
            cBManutencao = '{cBManutencao}',
            cBDecreto_Lei = '{cBDecreto_Lei}',
            cBConformidadeCE = '{cBConformidadeCE}'
        WHERE id_empresa = '{_idSelecionado}' AND serie = '{serie}';
    ";

            // Executa a query de atualização no banco de dados
            _BSO.DSO.ExecuteSQL(queryUpdate);

            // Mostra uma mensagem de confirmação
            MessageBox.Show("Trabalhador atualizado com sucesso.");

            // Limpa os campos de entrada após a atualização
            LimpaCamposEqui();

            // Opcional: Retorna o foco para o primeiro campo
            txt_marca.Focus();
        }

        private void InsereTrabalhador()
        {
            string nome = txt_nometrab.Text;  // Supondo que tenha um TextBox chamado txtNome
            string categoriatrab = txt_categoriatrab.Text;
            string contribuintetrab = txt_contribuintetrab.Text;
            string segurancasocialtrab = txt_segurancasocialtrab.Text;
            string emailTrab = rxt_emailTrabalhador.Text;
            int anexo1 = checkBox14.Checked ? 1 : 0;
            int anexo2 = checkBox15.Checked ? 1 : 0;
            int anexo3 = checkBox16.Checked ? 1 : 0;
            int anexo4 = checkBox17.Checked ? 1 : 0;
            int anexo5 = checkBox18.Checked ? 1 : 0;

            var cBFormacaoProfissional = cb_FormacaoProfissional.Text;
            var cBespecializados = cb_especializados.Text;

            string dtpnascimento2 = dtpnascimento.Value == DateTime.MinValue ? "NULL" : $"'{dtpnascimento.Value:yyyy-MM-dd HH:mm:ss}'";
            var datanasci = $"";

            if (dtpnascimento2 == "'1753-01-01 00:00:00'")
            {
                datanasci = $"";
            }
            else
            {
                datanasci = $"{dtpnascimento.Value:yyyy-MM-dd HH:mm:ss}";
            }
            string checkContribuinteQuery = $@"
            SELECT * FROM TDU_AD_Trabalhadores 
            WHERE contribuinte = '{contribuintetrab}' AND id_empresa = '{_idSelecionado}'
        ";

            var contribuinteExistente = _BSO.Consulta(checkContribuinteQuery);

            if (contribuinteExistente.NumLinhas() > 0)
            {
                MessageBox.Show("O contribuinte já está registrado. A inserção não será realizada.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // Se já existe, não prossegue com a inserção
            }

            dataGridView1.Rows.Add(nome, categoriatrab, contribuintetrab, segurancasocialtrab, emailTrab, anexo1, anexo2, anexo3, anexo4, anexo5, checkBox14.Text, checkBox15.Text, checkBox16.Text, checkBox17.Text, checkBox18.Text, cBFormacaoProfissional, cBespecializados, datanasci);

            // Aqui, você pode ocultar a coluna do checkBox Text (opcionalmente)
            int lastColumnIndex = dataGridView1.Columns.Count - 1; // Última coluna (onde você adicionou checkBox14.Text)
            dataGridView1.Columns[lastColumnIndex].Visible = false;
            // adcionar no sql
            string checkAndCreateColumnsQuery = $@"
            -- Verificar e criar a coluna 'id_empresa'
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'id_empresa')
                ALTER TABLE TDU_AD_Trabalhadores ADD id_empresa NVARCHAR(255);

            -- Verificar e criar a coluna 'nome'
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'nome')
                ALTER TABLE TDU_AD_Trabalhadores ADD nome NVARCHAR(255);

            -- Verificar e criar a coluna 'data_nascimento'
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'data_nascimento')
                ALTER TABLE TDU_AD_Trabalhadores ADD data_nascimento NVARCHAR(255);

            -- Verificar e criar a coluna 'categoria'
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'categoria')
                ALTER TABLE TDU_AD_Trabalhadores ADD categoria NVARCHAR(255);

            -- Verificar e criar a coluna 'contribuinte'
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'contribuinte')
                ALTER TABLE TDU_AD_Trabalhadores ADD contribuinte NVARCHAR(255);

            -- Verificar e criar a coluna 'seguranca_social'
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'seguranca_social')
                ALTER TABLE TDU_AD_Trabalhadores ADD seguranca_social NVARCHAR(255);

            -- Verificar e criar a coluna 'email'
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'email')
                ALTER TABLE TDU_AD_Trabalhadores ADD email NVARCHAR(255);

            -- Verificar e criar a coluna 'anexo1' (booleano)
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'anexo1')
                ALTER TABLE TDU_AD_Trabalhadores ADD anexo1 BIT;

            -- Verificar e criar a coluna 'anexo2' (booleano)
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'anexo2')
                ALTER TABLE TDU_AD_Trabalhadores ADD anexo2 BIT;

            -- Verificar e criar a coluna 'anexo3' (booleano)
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'anexo3')
                ALTER TABLE TDU_AD_Trabalhadores ADD anexo3 BIT;

            -- Verificar e criar a coluna 'anexo4' (booleano)
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'anexo4')
                ALTER TABLE TDU_AD_Trabalhadores ADD anexo4 BIT;

            -- Verificar e criar a coluna 'anexo5' (booleano)
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'anexo5')
                ALTER TABLE TDU_AD_Trabalhadores ADD anexo5 BIT;

            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'caminho1')
                ALTER TABLE TDU_AD_Trabalhadores ADD caminho1 NVARCHAR(255);
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'caminho2')
                ALTER TABLE TDU_AD_Trabalhadores ADD caminho2 NVARCHAR(255);
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'caminho3')
                ALTER TABLE TDU_AD_Trabalhadores ADD caminho3 NVARCHAR(255);
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'caminho4')
                ALTER TABLE TDU_AD_Trabalhadores ADD caminho4 NVARCHAR(255);
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'caminho5')
                ALTER TABLE TDU_AD_Trabalhadores ADD caminho5 NVARCHAR(255);
            ";

            // Executa a query de verificação e criação das colunas
            _BSO.DSO.ExecuteSQL(checkAndCreateColumnsQuery);

            // Fazendo uma substituição para caracteres especiais
            string caminho1 = SanitizeString(checkBox14.Text);
            string caminho2 = SanitizeString(checkBox15.Text);
            string caminho3 = SanitizeString(checkBox16.Text);
            string caminho4 = SanitizeString(checkBox17.Text);
            string caminho5 = SanitizeString(checkBox18.Text);
            string query = $@"
                INSERT INTO TDU_AD_Trabalhadores 
            (id_empresa, nome, categoria, contribuinte, seguranca_social, email, anexo1, anexo2, anexo3, anexo4, anexo5,caminho1,caminho2,caminho3,caminho4,caminho5,cBFormacaoProfissional,cBespecializados,data_nascimento) 
            VALUES 
            ('{_idSelecionado}', '{nome}', '{categoriatrab}', '{contribuintetrab}', '{segurancasocialtrab}', '{emailTrab}', {anexo1}, {anexo2}, {anexo3}, {anexo4}, {anexo5}, '{caminho1}', '{caminho2}', '{caminho3}', '{caminho4}', '{caminho5}','{cBFormacaoProfissional}','{cBespecializados}','{datanasci}')
            ";

            _BSO.DSO.ExecuteSQL(query);
            // Limpa os campos após adicionar
            LimpaCampos();
            // Opcional: Retorna o foco para o primeiro campo
            txt_nometrab.Focus();
        }

        private void InsereEquipamento()
        {
            string marca = txt_marca.Text;
            string tipo = txt_tipo.Text;
            string serie = txt_serie.Text;
            int anexo1 = checkBox19.Checked ? 1 : 0;
            int anexo2 = checkBox20.Checked ? 1 : 0;
            int anexo3 = checkBox21.Checked ? 1 : 0;
            int anexo4 = checkBox22.Checked ? 1 : 0;
            int anexo5 = checkBox23.Checked ? 1 : 0;

            // Verifica se a tabela existe e a cria se necessário
            string createTableQuery = @"
    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'TDU_AD_Equipamentos')
    BEGIN
        CREATE TABLE TDU_AD_Equipamentos (
            id_empresa NVARCHAR(255),
            marca NVARCHAR(255),
            tipo NVARCHAR(255),
            serie NVARCHAR(255) PRIMARY KEY,
            anexo1 BIT,
            anexo2 BIT,
            anexo3 BIT,
            anexo4 BIT,
            anexo5 BIT,
            caminho1 NVARCHAR(255),
            caminho2 NVARCHAR(255),
            caminho3 NVARCHAR(255),
            caminho4 NVARCHAR(255),
            caminho5 NVARCHAR(255)
        );
    END";

            _BSO.DSO.ExecuteSQL(createTableQuery);

            // Verifica se o equipamento já existe na tabela
            string checkContribuinteQuery = $@"
    SELECT * FROM TDU_AD_Equipamentos WHERE serie = '{serie}'";
            var contribuinteExistente = _BSO.Consulta(checkContribuinteQuery);

            if (contribuinteExistente.NumLinhas() > 0)
            {
                MessageBox.Show("A serie já está registrado. A inserção não será realizada.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var cBSeguro = cb_seguro.Text;
            var cBManualInstrucoes = cb_ManualInstrucoes.Text;
            var cBManutencao = cb_Manutencao.Text;
            var cBDecreto_Lei = cb_Decreto_Lei.Text;
            var cBConformidadeCE = cb_ConformidadeCE.Text;
            // Adiciona os dados ao DataGridView
            dataGridView2.Rows.Add(marca, tipo, serie, anexo1, anexo2, anexo3, anexo4, anexo5,
                                    checkBox19.Text, checkBox20.Text, checkBox21.Text, checkBox22.Text, checkBox23.Text, cBSeguro, cBManualInstrucoes, cBManutencao, cBDecreto_Lei, cBConformidadeCE);

            // Oculta a última coluna (se necessário)
            int lastColumnIndex = dataGridView2.Columns.Count - 1;
            dataGridView2.Columns[lastColumnIndex].Visible = false;

            // Sanitiza os caminhos
            string caminho1 = SanitizeString(checkBox19.Text);
            string caminho2 = SanitizeString(checkBox20.Text);
            string caminho3 = SanitizeString(checkBox21.Text);
            string caminho4 = SanitizeString(checkBox22.Text);
            string caminho5 = SanitizeString(checkBox23.Text);


            string checkColumnsQuery = @"
IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME = 'TDU_AD_Equipamentos' 
    AND COLUMN_NAME = 'cBSeguro')
BEGIN
    ALTER TABLE TDU_AD_Equipamentos ADD cBSeguro NVARCHAR(255);
END

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME = 'TDU_AD_Equipamentos' 
    AND COLUMN_NAME = 'cBManualInstrucoes')
BEGIN
    ALTER TABLE TDU_AD_Equipamentos ADD cBManualInstrucoes NVARCHAR(255);
END

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME = 'TDU_AD_Equipamentos' 
    AND COLUMN_NAME = 'cBManutencao')
BEGIN
    ALTER TABLE TDU_AD_Equipamentos ADD cBManutencao NVARCHAR(255);
END

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME = 'TDU_AD_Equipamentos' 
    AND COLUMN_NAME = 'cBDecreto_Lei')
BEGIN
    ALTER TABLE TDU_AD_Equipamentos ADD cBDecreto_Lei NVARCHAR(255);
END

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS 
    WHERE TABLE_NAME = 'TDU_AD_Equipamentos' 
    AND COLUMN_NAME = 'cBConformidadeCE')
BEGIN
    ALTER TABLE TDU_AD_Equipamentos ADD cBConformidadeCE NVARCHAR(255);
END


";

            _BSO.DSO.ExecuteSQL(checkColumnsQuery);
            // Insere os dados na tabela
            string insertQuery = $@"
    INSERT INTO TDU_AD_Equipamentos 
    (id_empresa, marca, tipo, serie, anexo1, anexo2, anexo3, anexo4, anexo5, caminho1, caminho2, caminho3, caminho4, caminho5, cBSeguro, cBManualInstrucoes,cBManutencao,cBDecreto_Lei) 
    VALUES ('{_idSelecionado}', '{marca}', '{tipo}', '{serie}', {anexo1}, {anexo2}, {anexo3}, {anexo4}, {anexo5}, '{caminho1}', '{caminho2}', '{caminho3}', '{caminho4}', '{caminho5}', '{cBSeguro}', '{cBManualInstrucoes}', '{cBManutencao}', '{cBDecreto_Lei}')";

            _BSO.DSO.ExecuteSQL(insertQuery);

            // Limpa os campos
            LimpaCamposEqui();
            txt_marca.Focus();
        }
        private void CarregarTrabalhadores()
        {
            // Verificar e criar a coluna email se não existir
            string checkEmailColumn = @"
            IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TDU_AD_Trabalhadores' AND COLUMN_NAME = 'email')
                ALTER TABLE TDU_AD_Trabalhadores ADD email NVARCHAR(255);";
            _BSO.DSO.ExecuteSQL(checkEmailColumn);

            // Consulta para buscar os trabalhadores na base de dados
            var query = $@"
            SELECT 
                nome, 
                categoria, 
                contribuinte, 
                seguranca_social,
                email,
                anexo1, 
                anexo2, 
                anexo3, 
                anexo4, 
                anexo5,
                caminho1,
                caminho2,
                caminho3,
                caminho4,
                caminho5,
                cBFormacaoProfissional,
                cBEspecializados,
                data_nascimento
            FROM TDU_AD_Trabalhadores
            WHERE id_empresa = '{_idSelecionado}';
        ";

            // Execute a consulta e recupere os dados
            var trabalhadores = _BSO.Consulta(query);

            dataGridView1.Rows.Clear();


            var numtrabalhadores = trabalhadores.NumLinhas();
            trabalhadores.Inicio();
            for (int i = 0; i < numtrabalhadores; i++)
            {
                var nome = trabalhadores.DaValor<string>("nome");
                var categoriatrab = trabalhadores.DaValor<string>("categoria");
                var contribuintetrab = trabalhadores.DaValor<string>("contribuinte");
                var segurancasocialtrab = trabalhadores.DaValor<string>("seguranca_social");
                var email = trabalhadores.DaValor<string>("email"); // Captura o email
                var anexo1 = trabalhadores.DaValor<bool>("anexo1");
                var anexo2 = trabalhadores.DaValor<bool>("anexo2");
                var anexo3 = trabalhadores.DaValor<bool>("anexo3");
                var anexo4 = trabalhadores.DaValor<bool>("anexo4");
                var anexo5 = trabalhadores.DaValor<bool>("anexo5");
                var caminho1 = RestoreSanitizedString(trabalhadores.DaValor<string>("caminho1"));
                var caminho2 = RestoreSanitizedString(trabalhadores.DaValor<string>("caminho2"));
                var caminho3 = RestoreSanitizedString(trabalhadores.DaValor<string>("caminho3"));
                var caminho4 = RestoreSanitizedString(trabalhadores.DaValor<string>("caminho4"));
                var caminho5 = RestoreSanitizedString(trabalhadores.DaValor<string>("caminho5"));

                var cBFormacaoProfissional = trabalhadores.DaValor<string>("cBFormacaoProfissional");
                var cBespecializados = trabalhadores.DaValor<string>("cBespecializados");

                var datanascimento = trabalhadores.DaValor<string>("data_nascimento");
                if (trabalhadores.DaValor<string>("data_nascimento").ToString() == "01/01/1753 00:00:00")
                {
                    dataGridView1.Rows.Add(nome, categoriatrab, contribuintetrab, segurancasocialtrab, anexo1, anexo2, anexo3, anexo4, anexo5, caminho1, caminho2, caminho3, caminho4, caminho5, cBFormacaoProfissional, cBespecializados, "");
                }
                else
                {
                    dataGridView1.Rows.Add(nome, categoriatrab, contribuintetrab, segurancasocialtrab, anexo1, anexo2, anexo3, anexo4, anexo5, caminho1, caminho2, caminho3, caminho4, caminho5, cBFormacaoProfissional, cBespecializados, datanascimento);

                }



                //dataGridView1.Rows.Add(nome, categoriatrab, contribuintetrab, segurancasocialtrab, anexo1, anexo2, anexo3, anexo4, anexo5, caminho1, caminho2, caminho3, caminho4, caminho5);
                trabalhadores.Seguinte();
            }
            cb_FormacaoProfissional.SelectedIndex = 0;
            cb_especializados.SelectedIndex = 0;


        }

        private void LimpaCamposEqui()
        {
            txt_marca.Text = "";
            txt_tipo.Text = "";
            txt_serie.Text = "";

            checkBox19.Checked = false;
            checkBox20.Checked = false;
            checkBox21.Checked = false;
            checkBox22.Checked = false;
            checkBox23.Checked = false;

            checkBox19.Text = "";
            checkBox20.Text = "";
            checkBox21.Text = "";
            checkBox22.Text = "";
            checkBox23.Text = "";


            button27.Visible = false;
            button26.Visible = false;
            txt_serie.Enabled = true;
            cb_seguro.SelectedIndex = 0;
            cb_ManualInstrucoes.SelectedIndex = 0;
            cb_Manutencao.SelectedIndex = 0;
            cb_Decreto_Lei.SelectedIndex = 0;
            cb_ConformidadeCE.SelectedIndex = 0;
            EditEqui = "0";
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Verifica se o usuário clicou em uma linha válida
            if (e.RowIndex >= 0)
            {
                // Obtém a linha selecionada
                DataGridViewRow row = dataGridView2.Rows[e.RowIndex];

                // Preenche os campos de texto com os valores da linha
                txt_marca.Text = row.Cells["marca"].Value.ToString();
                txt_tipo.Text = row.Cells["tipo"].Value.ToString();
                txt_serie.Text = row.Cells["Serieeq"].Value.ToString();
                checkBox19.Checked = ConvertToBool(row.Cells["anexo1"].Value);
                checkBox20.Checked = ConvertToBool(row.Cells["anexo2"].Value);
                checkBox21.Checked = ConvertToBool(row.Cells["anexo3"].Value);
                checkBox22.Checked = ConvertToBool(row.Cells["anexo4"].Value);
                checkBox23.Checked = ConvertToBool(row.Cells["anexo5"].Value);

                VerificarEColorirCheckBox(checkBox19, row.Cells["caminho6"].Value);
                VerificarEColorirCheckBox(checkBox20, row.Cells["caminho7"].Value);
                VerificarEColorirCheckBox(checkBox21, row.Cells["caminho8"].Value);
                VerificarEColorirCheckBox(checkBox22, row.Cells["caminho9"].Value);
                VerificarEColorirCheckBox(checkBox23, row.Cells["caminho10"].Value);
                cb_seguro.Text = row.Cells["CBSeguro"].Value.ToString();
                cb_ManualInstrucoes.Text = row.Cells["CBManualInstrucoes"].Value.ToString();
                cb_Manutencao.Text = row.Cells["CBManutencao"].Value.ToString();
                cb_Decreto_Lei.Text = row.Cells["CBDecreto_Lei"].Value.ToString();
                cb_ConformidadeCE.Text = row.Cells["CBConformidadeCE"].Value.ToString();

                txt_serie.Enabled = false;
                button27.Visible = true;
                button26.Visible = true;
                EditEqui = "1";

            }
        }

        private void CarregarEquipamentos()
        {
            // Consulta para buscar os trabalhadores na base de dados
            string query = $@"
        SELECT marca, tipo, serie, anexo1, anexo2, anexo3, anexo4, anexo5, caminho1, caminho2, caminho3, caminho4, caminho5, cBSeguro, cBManualInstrucoes, cBManutencao, cBDecreto_Lei, cBConformidadeCE
        FROM TDU_AD_Equipamentos
        WHERE id_empresa = '{_idSelecionado}';
        ";

            // Execute a consulta e recupere os dados
            var equipamentos = _BSO.Consulta(query);

            dataGridView1.Rows.Clear();


            var numtrabalhadores = equipamentos.NumLinhas();
            equipamentos.Inicio();
            for (int i = 0; i < numtrabalhadores; i++)
            {
                var nome = equipamentos.DaValor<string>("marca");
                var categoriatrab = equipamentos.DaValor<string>("tipo");
                var contribuintetrab = equipamentos.DaValor<string>("serie");
                var anexo1 = equipamentos.DaValor<bool>("anexo1");
                var anexo2 = equipamentos.DaValor<bool>("anexo2");
                var anexo3 = equipamentos.DaValor<bool>("anexo3");
                var anexo4 = equipamentos.DaValor<bool>("anexo4");
                var anexo5 = equipamentos.DaValor<bool>("anexo5");
                var caminho1 = equipamentos.DaValor<string>("caminho1");
                var caminho2 = equipamentos.DaValor<string>("caminho2");
                var caminho3 = equipamentos.DaValor<string>("caminho3");
                var caminho4 = equipamentos.DaValor<string>("caminho4");
                var caminho5 = RestoreSanitizedString(equipamentos.DaValor<string>("caminho5"));
                var cBSeguro = equipamentos.DaValor<string>("cBSeguro");
                var cBManualInstrucoes = equipamentos.DaValor<string>("cBManualInstrucoes");
                var cBManutencao = equipamentos.DaValor<string>("cBManutencao");
                var cBDecreto_Lei = equipamentos.DaValor<string>("cBDecreto_Lei");
                var cBConformidadeCE = equipamentos.DaValor<string>("cBConformidadeCE");


                dataGridView2.Rows.Add(nome, categoriatrab, contribuintetrab, anexo1, anexo2, anexo3, anexo4, anexo5, caminho1, caminho2, caminho3, caminho4, caminho5, cBSeguro, cBManualInstrucoes, cBManutencao, cBDecreto_Lei, cBConformidadeCE);
                equipamentos.Seguinte();
            }
            cb_seguro.SelectedIndex = 0;
            cb_ManualInstrucoes.SelectedIndex = 0;
            cb_Manutencao.SelectedIndex = 0;
            cb_Decreto_Lei.SelectedIndex = 0;
            cb_ConformidadeCE.SelectedIndex = 0;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            LimpaCamposEqui();
            button27.Visible = false;
            button26.Visible = false;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            LimpaCampos();
            button28.Visible = false;
            bt_remover.Visible = false;
        }

        private void bt_remover_Click(object sender, EventArgs e)
        {
            DialogResult resultado = MessageBox.Show("Tem certeza que deseja remover este trabalhador?",
                                         "Confirmação",
                                         MessageBoxButtons.YesNo,
                                         MessageBoxIcon.Warning);

            if (resultado == DialogResult.Yes)
            {
                var contri = txt_contribuintetrab.Text;
                var nomeEmpresa = TXT_Nome.Text.Replace(" ", "_");
                var nomeTrabalhador = txt_nometrab.Text.Replace(" ", "_");
                string pastaTrabalhador = Path.Combine(txt_caminhotrab.Text, nomeEmpresa, "Trabalhadores", nomeTrabalhador);

                // Verifica se a pasta existe antes de perguntar
                if (Directory.Exists(pastaTrabalhador))
                {
                    DialogResult resultadoPasta = MessageBox.Show(
                        "Deseja também remover a pasta com os documentos do trabalhador?",
                        "Remover Documentos",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (resultadoPasta == DialogResult.Yes)
                    {
                        try
                        {
                            Directory.Delete(pastaTrabalhador, true);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Erro ao remover a pasta: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                var removequery = $@"DELETE FROM TDU_AD_Trabalhadores 
                                    WHERE contribuinte = '{contri}' AND id_empresa = '{_idSelecionado}';";
                _BSO.DSO.ExecuteSQL(removequery);
                RemoverDoDataGridtrab(contri);
                LimpaCampos();
                MessageBox.Show("Trabalhador removido com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button28.Visible = false;
                bt_remover.Visible = false;
            }
        }

        private void RemoverDoDataGridtrab(string contribuinte)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["contribuinte"].Value?.ToString() == contribuinte) // Verifica a coluna correta
                {
                    dataGridView1.Rows.Remove(row);
                    break; // Sai do loop depois de remover
                }
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            DialogResult resultado = MessageBox.Show("Tem certeza que deseja remover este trabalhador?",
                             "Confirmação",
                             MessageBoxButtons.YesNo,
                             MessageBoxIcon.Warning);

            if (resultado == DialogResult.Yes)
            {

                var serie = txt_serie.Text;
                var nomeEmpresa = TXT_Nome.Text.Replace(" ", "_");
                var nomeEquipamento = txt_marca.Text.Replace(" ", "_");
                string pastaEquipamento = Path.Combine(txt_caminhoequi.Text, nomeEmpresa, "Equipamentos", nomeEquipamento);

                // Verifica se a pasta existe antes de perguntar
                if (Directory.Exists(pastaEquipamento))
                {
                    DialogResult resultadoPasta = MessageBox.Show(
                        "Deseja também remover a pasta com os documentos do equipamento?",
                        "Remover Documentos",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (resultadoPasta == DialogResult.Yes)
                    {
                        try
                        {
                            Directory.Delete(pastaEquipamento, true);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Erro ao remover a pasta: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                var removequery = $@"DELETE FROM TDU_AD_Equipamentos 
                                    WHERE serie = '{serie}' AND id_empresa = '{_idSelecionado}';";
                _BSO.DSO.ExecuteSQL(removequery);
                RemoverDoDataGridequi(serie);
                LimpaCamposEqui();
                MessageBox.Show("Equipamento removido com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button27.Visible = false;
                button26.Visible = false;
            }
        }
        private void RemoverDoDataGridequi(string serie)
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.Cells["Serieeq"].Value?.ToString() == serie) // Verifica a coluna correta
                {
                    dataGridView2.Rows.Remove(row);
                    break; // Sai do loop depois de remover
                }
            }
        }

        private void ObterObras()
        {
            var query = $@"SELECT 
    o.EntidadeIDA,
    o.TipoEmp,
    o.Codigo,
    o.Estado, 
    o.DataCriacao,
    o.*,
    e.*
FROM 
    COP_Obras o
JOIN 
    Geral_Entidade e
ON 
    o.EntidadeIDA = e.EntidadeId
WHERE 
    o.Estado IN ('CONS', 'ADJU') AND e.id = '{_idSelecionado}'
ORDER BY 
    o.Codigo DESC;";

            var lista = _BSO.Consulta(query);

            var num = lista.NumLinhas();
            lista.Inicio();
            for (int i = 0; i < num; i++)
            {
                string codigo = lista.Valor("Codigo").ToString();
                string nome = lista.Valor("Nome").ToString();


                cb_obras.Items.Add(new KeyValuePair<string, string>(codigo, $"{codigo} - {nome}"));
                cb_obras.DisplayMember = "Value";
                cb_obras.ValueMember = "Key";
                lista.Seguinte();
            }
        }

        private void bt_autorizar_Click(object sender, EventArgs e)
        {
            // Verifica se uma obra foi selecionada
            if (cb_obras.SelectedText == null)
            {
                MessageBox.Show("Por favor, selecione uma obra para autorizar.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            if (EditAut == "1")
            {
                AtualizaAutorizacao();
            }
            else
            {
                CriaAutorizacao();

            }


        }

        private void AtualizaAutorizacao()
        {

            string dataEntrada = dtpEntrada.Value == DateTime.MinValue ? "NULL" : $"'{dtpEntrada.Value:yyyy-MM-dd HH:mm:ss}'";
            string dataSaida = dtpSaida.Value == DateTime.MinValue ? "NULL" : $"'{dtpSaida.Value:yyyy-MM-dd HH:mm:ss}'";
            var nulo = false;
            int anexo1 = checkBox27.Checked ? 1 : 0;
            int anexo2 = checkBox25.Checked ? 1 : 0;
            int anexo3 = checkBox26.Checked ? 1 : 0;
            int anexo4 = checkBox7.Checked ? 1 : 0;
            int anexo5 = checkBox12.Checked ? 1 : 0;
            string caminho1 = SanitizeString(checkBox27.Text);
            string caminho2 = SanitizeString(checkBox25.Text);
            string caminho3 = SanitizeString(checkBox26.Text);
            string caminho4 = SanitizeString(checkBox7.Text);
            string caminho5 = SanitizeString(checkBox12.Text);
            if (dataSaida == "'1753-01-01 00:00:00'")
            {
                dataSaida = $"''";
                nulo = true;
            }

            MessageBox.Show(cb_obras.Text);
            var update = $@"UPDATE TDU_AD_Autorizacoes
                            SET Data_Entrada = {dataEntrada},
                                Data_Saida = {dataSaida},
                                anexo1 = {anexo1}, 
                                anexo2 = {anexo2}, 
                                anexo3 = {anexo3}, 
                                anexo4 = {anexo4}, 
                                anexo5 = {anexo5}, 
                                caminho1 = '{caminho1}',
                                caminho2 = '{caminho2}',
                                caminho3 = '{caminho3}',
                                caminho4 = '{caminho4}',
                                caminho5 = '{caminho5}'     
                                WHERE ID_Entidade = '{_idSelecionado}' AND Nome_Obra = '{cb_obras.Text}'";

            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (row.Cells["Obra"].Value.ToString() == cb_obras.Text)
                {
                    row.Cells["DataEntrada"].Value = dtpEntrada.Value.ToString("yyyy-MM-dd HH:mm:ss");
                    if (nulo == false)
                    {
                        row.Cells["DataSaida"].Value = dtpSaida.Value.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    else
                    {
                        row.Cells["DataSaida"].Value = "";
                    }


                    row.Cells["AnexoC"].Value = anexo1;
                    row.Cells["AnexoHTE"].Value = anexo2;
                    row.Cells["AnexoAPSS"].Value = anexo3;
                    row.Cells["AnexoDRE"].Value = anexo4;
                    row.Cells["AnexoDRTI"].Value = anexo5;

                    // Atualiza as labels de texto no DataGridView
                    row.Cells["caminho11"].Value = checkBox27.Text;
                    row.Cells["caminho12"].Value = checkBox25.Text;
                    row.Cells["caminho13"].Value = checkBox26.Text;
                    row.Cells["caminho14"].Value = checkBox7.Text;
                    row.Cells["caminho15"].Value = checkBox12.Text;

                    break;  // Atualiza a linha e sai do loop
                }
            }
            _BSO.DSO.ExecuteSQL(update);
            Limparcamposautorizar();

        }

        private void CriaAutorizacao()
        {
            DialogResult result = MessageBox.Show("Você tem certeza que deseja autorizar esta obra?",
                                      "Confirmação de Autorização",
                                      MessageBoxButtons.YesNo,
                                      MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {

                    var querytotalverifica = @"DECLARE @tableName NVARCHAR(128) = 'TDU_AD_Autorizacoes';

-- Verifica e adiciona cada coluna, se necessário
IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'ID_Entidade')
    ALTER TABLE TDU_AD_Autorizacoes ADD ID_Entidade NVARCHAR(500) NOT NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'Nome_Obra')
    ALTER TABLE TDU_AD_Autorizacoes ADD Nome_Obra NVARCHAR(500) NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'Codigo_Obra')
    ALTER TABLE TDU_AD_Autorizacoes ADD Codigo_Obra NVARCHAR(500) NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'Data_Entrada')
    ALTER TABLE TDU_AD_Autorizacoes ADD Data_Entrada DATETIME NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'Data_Saida')
    ALTER TABLE TDU_AD_Autorizacoes ADD Data_Saida DATETIME NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'anexo1')
    ALTER TABLE TDU_AD_Autorizacoes ADD anexo1 BIT NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'anexo2')
    ALTER TABLE TDU_AD_Autorizacoes ADD anexo2 BIT NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'anexo3')
    ALTER TABLE TDU_AD_Autorizacoes ADD anexo3 BIT NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'anexo4')
    ALTER TABLE TDU_AD_Autorizacoes ADD anexo4 BIT NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'caminho1')
    ALTER TABLE TDU_AD_Autorizacoes ADD caminho1 NVARCHAR(255) NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'caminho2')
    ALTER TABLE TDU_AD_Autorizacoes ADD caminho2 NVARCHAR(255) NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'caminho3')
    ALTER TABLE TDU_AD_Autorizacoes ADD caminho3 NVARCHAR(255) NULL;

IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = 'caminho4')
    ALTER TABLE TDU_AD_Autorizacoes ADD caminho4 NVARCHAR(255) NULL;
";
                    _BSO.DSO.ExecuteSQL(querytotalverifica);

                    // 1. Criar a tabela caso não exista
                    var verificaTabela = @"
        IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'TDU_AD_Autorizacoes')
        BEGIN
            CREATE TABLE TDU_AD_Autorizacoes (
                ID INT IDENTITY(1,1) PRIMARY KEY,
                ID_Entidade NVARCHAR(500) NOT NULL,
                Nome_Obra NVARCHAR(500) NOT NULL,
                Codigo_Obra NVARCHAR(500) NOT NULL,
                Data_Entrada DATETIME NULL,
                Data_Saida DATETIME NULL,
                anexo1 BIT,
                anexo2 BIT,
                anexo3 BIT,
                anexo4 BIT,
                caminho1 NVARCHAR(255),
                caminho2 NVARCHAR(255),
                caminho3 NVARCHAR(255),
                caminho4 NVARCHAR(255)
            );
        END";
                    _BSO.DSO.ExecuteSQL(verificaTabela);


                    var verificaObraExistente = $@"
    SELECT * 
    FROM TDU_AD_Autorizacoes 
    WHERE Nome_Obra = '{cb_obras.SelectedItem.ToString()}' AND  ID_Entidade =  '{_idSelecionado}'";

                    var obraExistente = _BSO.Consulta(verificaObraExistente);

                    if (obraExistente.NumLinhas() > 0)
                    {
                        MessageBox.Show("Esta obra já foi autorizada anteriormente.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }





                    // Formatar datas corretamente para SQL Server (YYYY-MM-DD HH:MM:SS)
                    string dataEntrada = dtpEntrada.Value == DateTime.MinValue ? "NULL" : $"'{dtpEntrada.Value:yyyy-MM-dd HH:mm:ss}'";
                    string dataSaida = dtpSaida.Value == DateTime.MinValue ? "NULL" : $"'{dtpSaida.Value:yyyy-MM-dd HH:mm:ss}'";
                    var selectedItem = (KeyValuePair<string, string>)cb_obras.SelectedItem;
                    string key = selectedItem.Key;

                    var dataen = $"{dtpEntrada.Value:yyyy-MM-dd HH:mm:ss}";
                    var datasai = $"";

                    if (dataSaida == "'1753-01-01 00:00:00'")
                    {
                        datasai = $"";
                    }
                    else
                    {
                        datasai = $"{dtpSaida.Value:yyyy-MM-dd HH:mm:ss}";
                    }

                    // 2. Inserir uma nova autorização~~

                    int anexo1 = checkBox27.Checked ? 1 : 0;
                    int anexo2 = checkBox25.Checked ? 1 : 0;
                    int anexo3 = checkBox26.Checked ? 1 : 0;
                    int anexo4 = checkBox7.Checked ? 1 : 0;
                    int anexo5 = checkBox12.Checked ? 1 : 0;
                    string caminho1 = SanitizeString(checkBox27.Text);
                    string caminho2 = SanitizeString(checkBox25.Text);
                    string caminho3 = SanitizeString(checkBox26.Text);
                    string caminho4 = SanitizeString(checkBox7.Text);
                    string caminho5 = SanitizeString(checkBox12.Text);


                    var insertAutorizacao = $@"
        INSERT INTO TDU_AD_Autorizacoes (ID_Entidade, Nome_Obra, Data_Entrada, Data_Saida, Codigo_Obra,anexo1,anexo2,anexo3,anexo4,anexo5, caminho1, caminho2, caminho3, caminho4, caminho5)
        VALUES ('{_idSelecionado}', '{cb_obras.SelectedItem.ToString()}', {dataEntrada}, {dataSaida}, '{key}','{anexo1}', '{anexo2}', '{anexo3}', '{anexo4}', '{anexo5}', '{caminho1}',
                '{caminho2}', '{caminho3}', '{caminho4}', '{caminho5}' )";

                    dataGridView3.Rows.Add(cb_obras.SelectedItem.ToString(), dataen, datasai, anexo1, anexo2, anexo3, anexo4, anexo5, true, checkBox27.Text, checkBox25.Text, checkBox26.Text, checkBox7.Text, checkBox12.Text, key); // ou false


                    _BSO.DSO.ExecuteSQL(insertAutorizacao);

                    // Mensagem de sucesso
                    MessageBox.Show("Obra autorizada com sucesso!", "Autorização", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    Limparcamposautorizar();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro ao autorizar a obra: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Limparcamposautorizar()
        {
            dtpEntrada.Value = DateTime.Now;
            dtpSaida.Value = DateTime.Now;
            cb_obras.SelectedIndex = -1;

            cb_obras.Text = "";
            cb_obras.SelectedText = "";
            button30.Visible = false;
            button29.Visible = false;

            checkBox27.Text = "";
            checkBox25.Text = "";
            checkBox26.Text = "";
            checkBox7.Text = "";
            checkBox12.Text = "";
            checkBox24.Checked = true;
            checkBox27.Checked = false;
            checkBox25.Checked = false;
            checkBox26.Checked = false;
            checkBox7.Checked = false;
            checkBox12.Checked = false;
            dtpSaida.Enabled = true;
            dtpSaida.Visible = true;
            dtpSaida.CustomFormat = "dd/MM/yyyy"; // Ou o formato que preferir
            dtpSaida.Value = DateTime.Today;
            EditAut = "0";
        }

        private void GetValoresAutorizarObras()
        {

            var verificaTabela = @"
        IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'TDU_AD_Autorizacoes')
        BEGIN
            CREATE TABLE TDU_AD_Autorizacoes (
                ID INT IDENTITY(1,1) PRIMARY KEY,
                ID_Entidade NVARCHAR(500) NOT NULL,
                Nome_Obra NVARCHAR(500) NOT NULL,
                Data_Entrada DATETIME NULL,
                Data_Saida DATETIME NULL,
                Contrato NVARCHAR(255) NULL
            );
        END";
            _BSO.DSO.ExecuteSQL(verificaTabela);

            var query = $@"SELECT * FROM TDU_AD_Autorizacoes WHERE ID_Entidade = '{_idSelecionado}';";
            var lista = _BSO.Consulta(query);

            var num = lista.NumLinhas();
            lista.Inicio();
            for (int i = 0; i < num; i++)
            {
                var nomeObra = lista.DaValor<string>("Nome_Obra");
                var dataentrada = lista.DaValor<DateTime>("Data_Entrada");
                var datasaida = lista.DaValor<DateTime>("Data_Saida");
                var anexo1 = lista.DaValor<bool>("anexo1");
                var anexo2 = lista.DaValor<bool>("anexo2");
                var anexo3 = lista.DaValor<bool>("anexo3");
                var anexo4 = lista.DaValor<bool>("anexo4");
                var anexo5 = lista.DaValor<bool>("anexo5");

                var caminho1 = RestoreSanitizedString(lista.DaValor<string>("caminho1"));
                var caminho2 = RestoreSanitizedString(lista.DaValor<string>("caminho2"));
                var caminho3 = RestoreSanitizedString(lista.DaValor<string>("caminho3"));
                var caminho4 = RestoreSanitizedString(lista.DaValor<string>("caminho4"));
                var caminho5 = RestoreSanitizedString(lista.DaValor<string>("caminho5"));

                //var caminho1 = lista.DaValor<string>("caminho1");
                //var caminho2 = lista.DaValor<string>("caminho2");
                //var caminho3 = lista.DaValor<string>("caminho3");
                //var caminho4 = lista.DaValor<string>("caminho4");
                var codigoobra = lista.DaValor<string>("Codigo_Obra");

                if (lista.DaValor<DateTime>("Data_Saida").ToString() == "01/01/1753 00:00:00")
                {
                    dataGridView3.Rows.Add(nomeObra, dataentrada, "", anexo1, anexo2, anexo3, anexo4, anexo5, true, caminho1, caminho2, caminho3, caminho4, caminho5, codigoobra);
                }
                else
                {
                    dataGridView3.Rows.Add(nomeObra, dataentrada, datasaida, anexo1, anexo2, anexo3, anexo4, anexo5, true, caminho1, caminho2, caminho3, caminho4, caminho5, codigoobra);

                }






                lista.Seguinte();
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Obtém a linha selecionada
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];

                // Obtém os valores das células da linha selecionada
                string codigo = row.Cells[12].Value?.ToString();
                string valor2 = row.Cells[1].Value?.ToString(); // Coluna 2 (Data Entrada)
                string valor3 = row.Cells[2].Value?.ToString(); // Coluna 3 (Data Saída)
                string valor1 = row.Cells[0].Value?.ToString(); // Coluna 1
                                                                // Atribui valores ao ComboBox, DateTimePickers e TextBox
                Obratexto = valor1;
                cb_obras.SelectedText = valor1;
                cb_obras.Text = valor1;

                var getcaminhoobra = $"SELECT CDU_CaminhoAuto FROM COP_Obras  WHERE Codigo = '{codigo}'";
                var caminho = _BSO.Consulta(getcaminhoobra);
                if (caminho.NumLinhas() > 0)
                {
                    var caminhocompleto = caminho.DaValor<string>("CDU_CaminhoAuto");
                    if (!string.IsNullOrEmpty(caminhocompleto))
                    {
                        txtcaminhoAuto.Text = caminhocompleto;
                    }
                    // 
                }


                // Tentativa de conversão das datas, com valor padrão se falhar
                if (DateTime.TryParse(valor2, out DateTime dataEntrada))
                {
                    dtpEntrada.Value = dataEntrada; // Define a data no DateTimePicker de Entrada
                }
                else
                {
                    // Se falhar, define a data atual ou uma data padrão
                    dtpEntrada.Value = DateTime.Now; // ou qualquer data padrão que você prefira
                }

                if (DateTime.TryParse(valor3, out DateTime dataSaida))
                {
                    dtpSaida.Enabled = true;
                    dtpSaida.Visible = true;
                    dtpSaida.CustomFormat = "dd/MM/yyyy";
                    dtpSaida.Value = dataSaida; // Define a data no DateTimePicker de Saída
                }
                else
                {
                    // Se falhar, define a data atual ou uma data padrão
                    checkBox24.Checked = false;
                    dtpSaida.Enabled = false;
                    dtpSaida.Visible = false;
                    dtpSaida.CustomFormat = " "; // Deixa a data em branco
                    dtpSaida.Value = new DateTime(1753, 1, 1);// ou qualquer data padrão que você prefira
                }

                checkBox27.Checked = ConvertToBool(row.Cells["AnexoC"].Value);
                checkBox25.Checked = ConvertToBool(row.Cells["AnexoHTE"].Value);
                checkBox26.Checked = ConvertToBool(row.Cells["AnexoAPSS"].Value);
                checkBox7.Checked = ConvertToBool(row.Cells["AnexoDRE"].Value);
                checkBox12.Checked = ConvertToBool(row.Cells["AnexoDRTI"].Value);

                VerificarEColorirCheckBox(checkBox27, row.Cells["caminho11"].Value);
                VerificarEColorirCheckBox(checkBox25, row.Cells["caminho12"].Value);
                VerificarEColorirCheckBox(checkBox26, row.Cells["caminho13"].Value);
                VerificarEColorirCheckBox(checkBox7, row.Cells["caminho14"].Value);
                VerificarEColorirCheckBox(checkBox12, row.Cells["caminho15"].Value);


                button30.Visible = true;
                button29.Visible = true;
                EditAut = "1";
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            Limparcamposautorizar();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            // Pergunta ao utilizador se ele tem a certeza que deseja remover
            DialogResult result = MessageBox.Show(
                "Tem a certeza que deseja remover esta obra?", // Mensagem
                "Confirmação de Remoção",                     // Título
                MessageBoxButtons.YesNo,                      // Botões Sim e Não
                MessageBoxIcon.Question);                     // Ícone de pergunta

            if (result == DialogResult.Yes)
            {
                // DELETE no banco de dados
                var nomeEmpresa = TXT_Nome.Text.Replace(" ", "_");
                string pastaAutorizacao = Path.Combine(txtcaminhoAuto.Text, nomeEmpresa, "Autorizacoes", cb_obras.Text.Replace(" ", "_"));

                // Verifica se a pasta existe antes de perguntar
                if (Directory.Exists(pastaAutorizacao))
                {
                    DialogResult resultadoPasta = MessageBox.Show(
                        "Deseja também remover a pasta com os documentos da autorização?",
                        "Remover Documentos",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (resultadoPasta == DialogResult.Yes)
                    {
                        try
                        {
                            Directory.Delete(pastaAutorizacao, true);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Erro ao remover a pasta: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                var delete = $@"DELETE TDU_AD_Autorizacoes 
                        WHERE Nome_Obra = '{Obratexto}' 
                        AND ID_Entidade = '{_idSelecionado}'";

                _BSO.DSO.ExecuteSQL(delete);
                // Remover a linha correspondente do DataGridView
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    if (row.Cells["Obra"].Value.ToString() == cb_obras.Text.ToString())
                    {
                        dataGridView3.Rows.Remove(row); // Remove a linha do DataGridView
                        break;  // Interrompe o loop depois de remover a linha
                    }
                }
                Limparcamposautorizar();
                MessageBox.Show("Obra removida com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("A remoção foi cancelada.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btsaveLink_Click(object sender, EventArgs e)
        {
            // Verifica se o campo txt_link está vazio
            if (string.IsNullOrWhiteSpace(txt_link.Text))
            {
                // Exibe uma mensagem de erro caso o campo esteja vazio
                MessageBox.Show("O link introduzido é inválido. Por favor, insira um link válido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                var querverifica = $@"-- Verifica se a coluna CDU_Link existe na tabela Geral_Entidade
IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_Link')
BEGIN
    -- Caso a coluna não exista, cria a coluna CDU_Link com o tipo nvarchar(max)
    ALTER TABLE Geral_Entidade
    ADD CDU_Link nvarchar(max);
END;";
                _BSO.DSO.ExecuteSQL(querverifica);

                var verificaExistente = $@"
IF EXISTS (SELECT 1 FROM Geral_Entidade WHERE ID = '{_idSelecionado}')
BEGIN
    -- Se o link já existir, faz o UPDATE
    UPDATE Geral_Entidade
    SET CDU_Link = '{txt_link.Text}'
    WHERE ID = '{_idSelecionado}';
END
ELSE
BEGIN
    -- Caso contrário, faz o INSERT
    INSERT INTO Geral_Entidade (CDU_Link)
    VALUES ('{txt_link.Text}');
END;";
                _BSO.DSO.ExecuteSQL(verificaExistente);

                MessageBox.Show("O link foi guardado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void checkBox24_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox24.Checked)
            {
                // Se o CheckBox estiver marcado, ativa o DateTimePicker para seleção de data
                dtpSaida.Enabled = true;
                dtpSaida.Visible = true;
                dtpSaida.CustomFormat = "dd/MM/yyyy"; // Ou o formato que preferir
                dtpSaida.Value = DateTime.Today;
            }
            else
            {
                // Se o CheckBox não estiver marcado, desabilita o DateTimePicker e limpa a data
                dtpSaida.Enabled = false;
                dtpSaida.Visible = false;
                dtpSaida.CustomFormat = " "; // Deixa a data em branco
                dtpSaida.Value = new DateTime(1753, 1, 1);
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {

            if (cb_obras.Text == null) // Verifica se algum item está selecionado
            {
                MessageBox.Show("Por favor, selecione uma obra antes de escolher a pasta.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // Interrompe a execução se nada estiver selecionadoNovoCodigoSelecionado
            }
            string codigoSelecionado = NovoCodigoSelecionado;
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Selecione a pasta para os documentos";
                folderDialog.ShowNewFolderButton = true;
                string checkColumnQuery = @"
                                        SELECT * 
                                        FROM INFORMATION_SCHEMA.COLUMNS 
                                        WHERE TABLE_NAME = 'COP_Obras' 
                                        AND COLUMN_NAME = 'CDU_CaminhoAuto'";
                var columnExists = _BSO.Consulta(checkColumnQuery);
                if (columnExists.NumLinhas() > 0)
                {
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        txtcaminhoAuto.Text = folderDialog.SelectedPath;
                        var update = $@"UPDATE COP_Obras
                                set CDU_CaminhoAuto = '{txtcaminhoAuto.Text}'
                                WHERE Codigo = '{codigoSelecionado}'";
                        _BSO.DSO.ExecuteSQL(update);

                    }
                }
                else
                {
                    // Cria a coluna se não existir
                    string alterTableQuery = @"
                    ALTER TABLE COP_Obras 
                    ADD CDU_CaminhoAuto NVARCHAR(500)"; // Ajuste o tipo de dado conforme necessário

                    _BSO.DSO.ExecuteSQL(alterTableQuery);
                }

            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            string caminhoPasta = txtcaminhoAuto.Text;

            // Verificar se o caminho da pasta existe
            if (Directory.Exists(caminhoPasta))
            {
                // Abrir a pasta no explorador de arquivos
                Process.Start("explorer.exe", caminhoPasta);
            }
            else
            {
                MessageBox.Show("O caminho da pasta não é válido.");
            }
        }

        private void cb_obras_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cb_obras.SelectedItem != null)
            {
                var selectedPair = (KeyValuePair<string, string>)cb_obras.SelectedItem;
                NovoCodigoSelecionado = selectedPair.Key;
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked)
            {
                // Se o CheckBox estiver marcado, ativa o DateTimePicker para seleção de data
                dtpnascimento.Enabled = true;
                dtpnascimento.Visible = true;
                dtpnascimento.CustomFormat = "dd/MM/yyyy"; // Ou o formato que preferir
                dtpnascimento.Value = DateTime.Today;
            }
            else
            {
                // Se o CheckBox não estiver marcado, desabilita o DateTimePicker e limpa a data
                dtpnascimento.Enabled = false;
                dtpnascimento.Visible = false;
                dtpnascimento.CustomFormat = " "; // Deixa a data em branco
                dtpnascimento.Value = new DateTime(1753, 1, 1);
            }
        }

        private void Bt_Nuvem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(LinkNuvem))
            {
                // Show the custom form where the user can enter the link
                using (AddLinkForm form = new AddLinkForm())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        LinkNuvem = form.Link; // Assuming Link is a public property of AddLinkForm
                        var updatelink = $@"UPDATE Geral_Entidade
                                        set CDU_LinkNuvem = '{LinkNuvem}'
                                        Where ID='{_idSelecionado}'";
                        _BSO.DSO.ExecuteSQL(updatelink);
                        TXTOP_linknuvem.Text = LinkNuvem;
                        // Try to open the link if it's not empty
                        if (!string.IsNullOrEmpty(LinkNuvem))
                        {
                            try
                            {
                                System.Diagnostics.Process.Start(LinkNuvem);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error opening the link: " + ex.Message);
                            }
                        }
                    }
                }
            }
            else
            {
                try
                {
                    System.Diagnostics.Process.Start(LinkNuvem);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening the link: " + ex.Message);
                }
            }
        }

        private void bt_gOpcoes_Click(object sender, EventArgs e)
        {
            DialogResult resultado = MessageBox.Show(
                "Tem a certeza de que deseja guardar estas opções?",
                "Confirmação",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (resultado == DialogResult.Yes)
            {
                var envidoop = CBOP_Enviado.Checked ? 1 : 0;
                var trataop = CBOP_SGS.Checked ? 1 : 0;
                string dataenvioStr = DTPOP_DataEnvio.Value.ToString("yyyy-MM-dd HH:mm:ss");
                LinkNuvem = TXTOP_linknuvem.Text;
                string empty = "";
                if (DTPOP_DataEnvio.Visible == false)
                {
                    var updateopcoes = $@"UPDATE Geral_Entidade
                                set CDU_DataEnvio = '{empty}',
                                    CDU_TrataSGS = '{trataop}',
                                    CDU_EmailEnviado = '{envidoop}',
                                    CDU_LinkNuvem = '{TXTOP_linknuvem.Text}'
                                Where ID='{_idSelecionado}'";
                    _BSO.DSO.ExecuteSQL(updateopcoes);


                    //update txt_email
                    var updateEmail = $@"UPDATE ec
                                    SET ec.Email = '{txt_email.Text}'
                                    FROM Geral_Entidade_Contactos ec
                                    INNER JOIN Geral_Entidade ge ON CAST(ge.ID AS uniqueidentifier) = ec.EntidadeID
                                    WHERE ge.ID = '{_idSelecionado}'
                                    ";
                    _BSO.DSO.ExecuteSQL(updateEmail);
                }
                else
                {
                    var updateopcoes = $@"UPDATE Geral_Entidade
                                set CDU_DataEnvio = '{dataenvioStr}',
                                    CDU_TrataSGS = '{trataop}',
                                    CDU_EmailEnviado = '{envidoop}',
                                    CDU_LinkNuvem = '{TXTOP_linknuvem.Text}'
                                Where ID='{_idSelecionado}'";
                    _BSO.DSO.ExecuteSQL(updateopcoes);
                    var updateEmail = $@"UPDATE ec
                                    SET ec.Email = '{txt_email.Text}'
                                    FROM Geral_Entidade_Contactos ec
                                    INNER JOIN Geral_Entidade ge ON CAST(ge.ID AS uniqueidentifier) = ec.EntidadeID
                                    WHERE ge.ID = '{_idSelecionado}'
                                    ";
                    _BSO.DSO.ExecuteSQL(updateEmail);
                }


                // Código para guardar as opções
                MessageBox.Show("As opções foram guardadas com sucesso.", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CBOP_Enviado_CheckedChanged(object sender, EventArgs e)
        {
            if (CBOP_Enviado.Checked == true)
            {
                DTPOP_DataEnvio.Visible = true;
                datavalor.Visible = false;
            }
            else
            {
                DTPOP_DataEnvio.Visible = false;
                datavalor.Visible = true;
            }

        }

        private void EnviarEmailOutlook(string destinatario, string assunto, string corpo)
        {
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.To = destinatario;
            mailItem.Subject = assunto;
            mailItem.Body = corpo;
            mailItem.Display(); // Mostra o Outlook com o email preenchido, mas não envia automaticamente
        }

        private void Bt_Caducado_Click(object sender, EventArgs e)
        {
            string query = $"SELECT id, Nome, CDU_EmailEnviado, CDU_DataEnvio FROM Geral_Entidade WHERE CDU_TrataSGS = 1 AND ID = '{_idSelecionado}' ";
            StdBELista dt = _BSO.Consulta(query);
            dt.Inicio();

            for (int i = 0; i < dt.NumLinhas(); i++)
            {
                string id = dt.DaValor<string>("id");
                string nomeEntidade = dt.DaValor<string>("Nome");

                // Lista de documentos da empresa
                List<string> documentosEmpresa = VerificaDocumentosDetalhados(id);

                // Dicionário de trabalhador -> documentos caducados
                Dictionary<string, List<string>> documentosTrabalhadores = VerificaDocumentosTrabalhadores(id);

                // Dicionário de equipamentos -> documentos caducados
                Dictionary<string, List<string>> documentosEquipamentos = VerificaDocumentosEquipamentos(id);

                // Dicionário de autorizações -> documentos caducados
                Dictionary<string, List<string>> documentosAutorizacoes = VerificaDocumentosAutorizacoes(id);

                if (documentosEmpresa.Count > 0 || documentosTrabalhadores.Count > 0 || documentosEquipamentos.Count > 0 || documentosAutorizacoes.Count > 0)
                {
                    StringBuilder corpo = new StringBuilder();
                    corpo.AppendLine("Prezado(a),");
                    corpo.AppendLine();
                    corpo.AppendLine($"A entidade \"{nomeEntidade}\" tem documentos caducados.");
                    corpo.AppendLine();

                    if (documentosEmpresa.Any())
                    {
                        corpo.AppendLine("📁 Documentos da Empresa:");
                        foreach (var doc in documentosEmpresa)
                        {
                            corpo.AppendLine($"- {doc}");
                        }
                        corpo.AppendLine();
                    }

                    if (documentosTrabalhadores.Any())
                    {
                        corpo.AppendLine("👷 Documentos por Trabalhador:");
                        foreach (var trabalhador in documentosTrabalhadores)
                        {
                            corpo.AppendLine($"\n{trabalhador.Key}:");
                            foreach (var doc in trabalhador.Value)
                            {
                                corpo.AppendLine($"- {doc}");
                            }
                        }
                        corpo.AppendLine();
                    }

                    if (documentosEquipamentos.Any())
                    {
                        corpo.AppendLine("🔧 Documentos de Equipamentos:");
                        foreach (var equipamento in documentosEquipamentos)
                        {
                            corpo.AppendLine($"\n{equipamento.Key}:");
                            foreach (var doc in equipamento.Value)
                            {
                                corpo.AppendLine($"- {doc}");
                            }
                        }
                        corpo.AppendLine();
                    }

                    if (documentosAutorizacoes.Any())
                    {
                        corpo.AppendLine("🔑 Documentos de Autorizações:");
                        foreach (var autorizacao in documentosAutorizacoes)
                        {
                            corpo.AppendLine($"\n{autorizacao.Key}:");
                            foreach (var doc in autorizacao.Value)
                            {
                                corpo.AppendLine($"- {doc}");
                            }
                        }
                        corpo.AppendLine();
                    }

                    corpo.AppendLine("\nPor favor, regularize esta situação com urgência.");
                    corpo.AppendLine("\nObrigado.");

                    EnviarEmailOutlook("departamento@email.pt", $"Alerta Documentos Caducados - {nomeEntidade}", corpo.ToString());
                }

                dt.Seguinte();
            }
        }

        private void check_AlertaCaducados_CheckedChanged(object sender, EventArgs e)
        {
            var result = check_AlertaCaducados.Checked;
            if (result)
            {
                Bt_Caducado.Enabled = false;
                var entidadeid = _idSelecionado;
                var update = $@"UPDATE Geral_Entidade
                                set CDU_IgnoraAlerta = '{result}'
                                WHERE ID = '{entidadeid}'";
                _BSO.DSO.ExecuteSQL(update);
            }
            else
            {
                Bt_Caducado.Enabled = true;
                var entidadeid = _idSelecionado;

                var update = $@"UPDATE Geral_Entidade
                                set CDU_IgnoraAlerta = '{result}'
                                WHERE ID = '{entidadeid}'";
                _BSO.DSO.ExecuteSQL(update);
            }


        }

        private List<string> VerificaDocumentosDetalhados(string id)
        {
            List<string> caducados = new List<string>();

            // Reutilizar a lógica da tua função anterior, mas guardar os nomes dos documentos caducados
            var campos = new Dictionary<string, string>()
            {
                {"CDU_ValidadeFinancas", "DND-Finanças"},
                {"CDU_ValidadeSegSocial", "DND-Segurança-Social"},
                {"CDU_ValidadeFolhaPag", "Mapa de Rem. – SS"},
                {"CDU_ValidadeComprovativoPagamento", "TSU"},
                {"CDU_ValidadeReciboSeguroAT", "Seguro AT"},
                {"CDU_ValidadeSeguroRC", "Seguro RC"},
                {"CDU_ValidadeSeguroAT", "condições Seguro AT"},
                {"CDU_ValidadeAlvara", "Alvará"},
                {"CDU_ValidadeCertidaoPermanente", "Certidão Permanente"}
            };

            string query = $"SELECT {string.Join(",", campos.Keys)} FROM Geral_Entidade WHERE ID = '{id}'";
            var res = _BSO.Consulta(query);



            /* var camposTrabalhadores = new Dictionary<string, string>()
             {
                 {"caminho1", "Cartão de cidadão ou residência"},
                 {"caminho2", "Ficha Medica"},
                 {"caminho3", "FormacaoProfissional"},
                 {"caminho4", "Trabalhoss especializados"},
                 {"caminho5", "Ficha Destribuiçao"}
             };

             //Por trabalhador
             string querytrab = $@"SELECT {string.Join(",", camposTrabalhadores.Keys)} FROM TDU_AD_Trabalhadores WHERE id_empresa = '{id}'";
             var resTrab = _BSO.Consulta(querytrab);
             resTrab.Inicio();
             var numtrab = resTrab.NumLinhas();
             if (numtrab == 0)
             {
                 return caducados;
             }
             for (int i = 0; i < numtrab; i++)
             {


                 foreach (var campo in camposTrabalhadores)
                 {

                     string valorOriginal = resTrab.DaValor<string>(campo.Key);

                     string valorDecodificado = WebUtility.HtmlDecode(valorOriginal);

                     var match = Regex.Match(valorDecodificado, @"\d{2}[\/\-]\d{2}[\/\-]\d{4}");


                     if (match.Success)
                     {
                         if (DateTime.TryParse(match.Value, out DateTime validade))
                         {
                             if (validade < DateTime.Now && validade != DateTime.MinValue)
                             {
                                 caducados.Add(campo.Value);
                             }
                         }
                     }
                 }

                 resTrab.Seguinte();

             }*/




            res.Inicio();

            foreach (var campo in campos)
            {
                DateTime validade;
                if (DateTime.TryParse(res.DaValor<string>(campo.Key), out validade))
                {
                    if (validade < DateTime.Now && validade != DateTime.MinValue)
                    {
                        caducados.Add(campo.Value);
                    }
                }
            }

            // Podes replicar esta mesma lógica para trabalhadores, equipamentos e autorizações se quiseres mais detalhe

            return caducados;
        }
        private Dictionary<string, List<string>> VerificaDocumentosTrabalhadores(string idEmpresa)
        {
            var resultado = new Dictionary<string, List<string>>();

            var camposTrabalhadores = new Dictionary<string, string>()
    {
        {"caminho1", "Cartão de cidadão ou residência"},
        {"caminho2", "Ficha Medica"},
        {"caminho3", "FormacaoProfissional"},
        {"caminho4", "Trabalhoss especializados"},
        {"caminho5", "Ficha Destribuiçao"}
    };

            // Supondo que tens um campo com o nome ou identificador do trabalhador
            string querytrab = $@"SELECT Nome, {string.Join(",", camposTrabalhadores.Keys)} FROM TDU_AD_Trabalhadores WHERE id_empresa = '{idEmpresa}'";
            var resTrab = _BSO.Consulta(querytrab);

            resTrab.Inicio();
            for (int i = 0; i < resTrab.NumLinhas(); i++)
            {
                string nomeTrab = resTrab.DaValor<string>("nome");

                var documentosCaducados = new List<string>();

                foreach (var campo in camposTrabalhadores)
                {
                    string valorOriginal = resTrab.DaValor<string>(campo.Key);
                    if (string.IsNullOrWhiteSpace(valorOriginal)) continue;

                    string valorDecodificado = WebUtility.HtmlDecode(valorOriginal);

                    var match = Regex.Match(valorDecodificado, @"\d{2}[\/\-]\d{2}[\/\-]\d{4}");
                    if (match.Success)
                    {
                        if (DateTime.TryParse(match.Value, out DateTime validade))
                        {
                            if (validade < DateTime.Now && validade != DateTime.MinValue)
                            {
                                documentosCaducados.Add(campo.Value);
                            }
                        }
                    }
                }

                if (documentosCaducados.Any())
                {
                    resultado[nomeTrab] = documentosCaducados;
                }

                resTrab.Seguinte(); // move para o próximo registo
            }


            return resultado;
        }

        private Dictionary<string, List<string>> VerificaDocumentosEquipamentos(string idEmpresa)
        {
            var resultado = new Dictionary<string, List<string>>();

            var camposEquipamentos = new Dictionary<string, string>()
    {
        {"caminho5", "Outro Documento Relevante"}
    };

            string queryEquip = $@"SELECT marca, {string.Join(",", camposEquipamentos.Keys)} FROM TDU_AD_Equipamentos WHERE id_empresa = '{idEmpresa}'";
            var resEquip = _BSO.Consulta(queryEquip);

            resEquip.Inicio();
            for (int i = 0; i < resEquip.NumLinhas(); i++)
            {
                string nomeEquip = resEquip.DaValor<string>("marca")?.Trim();
                if (string.IsNullOrEmpty(nomeEquip)) nomeEquip = "(Sem Nome)";

                var documentosCaducados = new List<string>();

                foreach (var campo in camposEquipamentos)
                {
                    string valorOriginal = resEquip.DaValor<string>(campo.Key);
                    if (string.IsNullOrWhiteSpace(valorOriginal)) continue;

                    string valorDecodificado = WebUtility.HtmlDecode(valorOriginal);

                    var match = Regex.Match(valorDecodificado, @"\d{2}[\/\-]\d{2}[\/\-]\d{4}");
                    if (match.Success)
                    {
                        if (DateTime.TryParse(match.Value, out DateTime validade))
                        {
                            if (validade < DateTime.Now && validade != DateTime.MinValue)
                            {
                                documentosCaducados.Add(campo.Value);
                            }
                        }
                    }
                }

                if (documentosCaducados.Any())
                {
                    resultado[nomeEquip] = documentosCaducados;
                }

                resEquip.Seguinte();
            }

            return resultado;
        }

        private Dictionary<string, List<string>> VerificaDocumentosAutorizacoes(string idEmpresa)
        {
            var resultado = new Dictionary<string, List<string>>();

            var camposAutorizacoes = new Dictionary<string, string>()
            {
                {"caminho1", "Contrato"},
                {"caminho2", "Horário de trabalho da empreitada"},
                {"caminho3", "Declaração de adesão ao PSS"},
                {"caminho4", "Declaração do resposável no estaleiro"}
                // Adicione mais campos conforme necessário para as autorizações.
            };

            // Supondo que tens um campo com o nome ou identificador da autorização
            string queryAutorizacoes = $@"SELECT Codigo_Obra, {string.Join(",", camposAutorizacoes.Keys)} FROM TDU_AD_Autorizacoes WHERE ID_Entidade = '{idEmpresa}'";
            var resAutorizacoes = _BSO.Consulta(queryAutorizacoes);

            resAutorizacoes.Inicio();
            for (int i = 0; i < resAutorizacoes.NumLinhas(); i++)
            {
                string nomeAutorizacao = resAutorizacoes.DaValor<string>("Codigo_Obra");

                var documentosCaducados = new List<string>();

                foreach (var campo in camposAutorizacoes)
                {
                    string valorOriginal = resAutorizacoes.DaValor<string>(campo.Key);
                    if (string.IsNullOrWhiteSpace(valorOriginal)) continue;

                    string valorDecodificado = WebUtility.HtmlDecode(valorOriginal);

                    var match = Regex.Match(valorDecodificado, @"\d{2}[\/\-]\d{2}[\/\-]\d{4}");
                    if (match.Success)
                    {
                        if (DateTime.TryParse(match.Value, out DateTime validade))
                        {
                            if (validade < DateTime.Now && validade != DateTime.MinValue)
                            {
                                documentosCaducados.Add(campo.Value);
                            }
                        }
                    }
                }

                if (documentosCaducados.Any())
                {
                    resultado[nomeAutorizacao] = documentosCaducados;
                }

                resAutorizacoes.Seguinte(); // move para o próximo registo
            }

            return resultado;
        }

        private void check_AlertaCaducados_CheckedChanged_1(object sender, EventArgs e)
        {
            var result = check_AlertaCaducados.Checked;
            if (result)
            {
                Bt_Caducado.Enabled = false;
                var entidadeid = _idSelecionado;
                var update = $@"UPDATE Geral_Entidade
                                set CDU_IgnoraAlerta = '{result}'
                                WHERE ID = '{entidadeid}'";
                _BSO.DSO.ExecuteSQL(update);
            }
            else
            {
                Bt_Caducado.Enabled = true;
                var entidadeid = _idSelecionado;

                var update = $@"UPDATE Geral_Entidade
                                set CDU_IgnoraAlerta = '{result}'
                                WHERE ID = '{entidadeid}'";
                _BSO.DSO.ExecuteSQL(update);
            }
        }
    }
}