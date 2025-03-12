using ErpBS100;
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

        private void AtualizarStatusDocumento(string tipoDocumento, string caminho, DateTime dataValidade)
        {
            try
            {
                // Atualizar a tabela Geral_Entidade com o caminho do documento e sua validade
                string colunaCaminho = "CDU_Caminho";
                string colunaAnexo = $"CDU_Anexo{tipoDocumento.Replace(" ", "")}";
                string colunaValidade = $"CDU_Validade{tipoDocumento.Replace(" ", "")}";

                string query = $@"UPDATE Geral_Entidade SET 
                                {colunaCaminho} = '{caminho}',
                                {colunaAnexo} = 1,
                                {colunaValidade} = '{dataValidade.ToString("yyyy-MM-dd")}'
                                WHERE Id = '{_idSelecionado}'";
                _BSO.DSO.ExecuteSQL(query);
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
                // Consulta SQL para verificar se o documento está anexado
                string coluna = $"CDU_Anexo{tipoDocumento.Replace(" ", "")}";
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
    }
}
