

using ErpBS100;
using StdBE100;
using StdPlatBS100;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class DadosJPA : Form
    {
        private readonly ErpBS _BSO;
        private readonly StdBSInterfPub _PSO;
        private readonly string _idJPA = "2A8C7ECD-309B-49F9-A337-203B45CED948";

        private TextBox TXT_Codigo;
        private TextBox TXT_Nome;
        private TextBox TXT_Contribuinte;
        private TextBox txtCaminhoPasta;
        private CheckBox checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6;
        private CheckBox checkBox8, checkBox9, checkBox10, checkBox13, checkBox28;
        private Button button1, button2, button3, button4, button5, button6;
        private Button button8, button9, button10, button12, button13;

        public DadosJPA(ErpBS BSO, StdBSInterfPub PSO)
        {
            InitializeComponent();
            _BSO = BSO;
            _PSO = PSO;
            ConfigurarFormulario();
            CarregarDados();
            InitializeButtonEvents();
        }

        private void ConfigurarFormulario()
        {
            this.Text = "Dados JPA - Joaquim Peixoto Azevedo & Filhos, Lda";
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterParent;

            // Painel principal
            Panel panelMain = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                AutoScroll = true
            };

            // Label título
            Label lblTitulo = new Label
            {
                Text = "JOAQUIM PEIXOTO AZEVEDO & FILHOS, LDA",
                Font = new Font("Calibri", 14F, FontStyle.Bold),
                Location = new Point(10, 10),
                AutoSize = true,
                ForeColor = Color.FromArgb(59, 89, 152)
            };

            // GroupBox Dados da Empresa
            GroupBox gbDadosEmpresa = new GroupBox
            {
                Text = "Dados da Empresa",
                Location = new Point(10, 50),
                Size = new Size(850, 120),
                Font = new Font("Calibri", 10F, FontStyle.Bold)
            };

            TXT_Codigo = new TextBox { Location = new Point(150, 25), Size = new Size(200, 23), ReadOnly = true };
            TXT_Nome = new TextBox { Location = new Point(150, 55), Size = new Size(650, 23), ReadOnly = true };
            TXT_Contribuinte = new TextBox { Location = new Point(150, 85), Size = new Size(200, 23), ReadOnly = true };

            Label lblCodigo = new Label { Text = "Código:", Location = new Point(20, 28), AutoSize = true };
            Label lblNome = new Label { Text = "Nome:", Location = new Point(20, 58), AutoSize = true };
            Label lblNIF = new Label { Text = "NIF:", Location = new Point(20, 88), AutoSize = true };

            gbDadosEmpresa.Controls.AddRange(new Control[] {
                lblCodigo, TXT_Codigo, lblNome, TXT_Nome, lblNIF, TXT_Contribuinte
            });

            // GroupBox Caminho dos Documentos
            GroupBox gbCaminho = new GroupBox
            {
                Text = "Caminho dos Documentos",
                Location = new Point(10, 180),
                Size = new Size(850, 80),
                Font = new Font("Calibri", 10F, FontStyle.Bold)
            };

            txtCaminhoPasta = new TextBox { Location = new Point(20, 30), Size = new Size(650, 23), ReadOnly = true };
            Button btnSelecionarPasta = new Button
            {
                Text = "Selecionar Pasta",
                Location = new Point(680, 28),
                Size = new Size(120, 28)
            };
            btnSelecionarPasta.Click += BtnSelecionarPasta_Click;

            gbCaminho.Controls.AddRange(new Control[] { txtCaminhoPasta, btnSelecionarPasta });

            // GroupBox Documentos
            GroupBox gbDocumentos = new GroupBox
            {
                Text = "Documentos da Empresa",
                Location = new Point(10, 270),
                Size = new Size(850, 250),
                Font = new Font("Calibri", 10F, FontStyle.Bold)
            };

            int yPos = 25;
            int xPosCheck = 20;
            int xPosButton = 330;

            // Criar checkboxes e botões para cada documento
            checkBox1 = CriarCheckboxDocumento("DND-Finanças", xPosCheck, yPos);
            button1 = CriarBotaoAnexar(xPosButton, yPos);
            yPos += 35;

            checkBox2 = CriarCheckboxDocumento("DND-Segurança-Social", xPosCheck, yPos);
            button2 = CriarBotaoAnexar(xPosButton, yPos);
            yPos += 35;

            checkBox3 = CriarCheckboxDocumento("Mapa de Rem. – SS", xPosCheck, yPos);
            button3 = CriarBotaoAnexar(xPosButton, yPos);
            yPos += 35;

            checkBox4 = CriarCheckboxDocumento("TSU", xPosCheck, yPos);
            button4 = CriarBotaoAnexar(xPosButton, yPos);
            yPos += 35;

            checkBox5 = CriarCheckboxDocumento("Seguro AT", xPosCheck, yPos);
            button5 = CriarBotaoAnexar(xPosButton, yPos);
            yPos += 35;

            checkBox6 = CriarCheckboxDocumento("Seguro RC", xPosCheck, yPos);
            button6 = CriarBotaoAnexar(xPosButton, yPos);

            yPos = 25;
            int xPosCheck2 = 450;
            int xPosButton2 = 760;

            checkBox8 = CriarCheckboxDocumento("Condições Seguro AT", xPosCheck2, yPos);
            button8 = CriarBotaoAnexar(xPosButton2, yPos);
            yPos += 35;

            checkBox9 = CriarCheckboxDocumento("Alvará", xPosCheck2, yPos);
            button9 = CriarBotaoAnexar(xPosButton2, yPos);
            yPos += 35;

            checkBox10 = CriarCheckboxDocumento("Certidão Permanente", xPosCheck2, yPos);
            button10 = CriarBotaoAnexar(xPosButton2, yPos);
            yPos += 35;

            checkBox13 = CriarCheckboxDocumento("Condições Seguro RC", xPosCheck2, yPos);
            button12 = CriarBotaoAnexar(xPosButton2, yPos);
            yPos += 35;

            checkBox28 = CriarCheckboxDocumento("Anexo D", xPosCheck2, yPos);
            button13 = CriarBotaoAnexar(xPosButton2, yPos);

            gbDocumentos.Controls.AddRange(new Control[] {
                checkBox1, button1, checkBox2, button2, checkBox3, button3,
                checkBox4, button4, checkBox5, button5, checkBox6, button6,
                checkBox8, button8, checkBox9, button9, checkBox10, button10,
                checkBox13, button12, checkBox28, button13
            });

            panelMain.Controls.AddRange(new Control[] {
                lblTitulo, gbDadosEmpresa, gbCaminho, gbDocumentos
            });

            this.Controls.Add(panelMain);
        }

        private CheckBox CriarCheckboxDocumento(string texto, int x, int y)
        {
            return new CheckBox
            {
                Text = texto,
                Location = new Point(x, y),
                AutoSize = false,
                Size = new Size(300, 20),
                Enabled = false
            };
        }

        private Button CriarBotaoAnexar(int x, int y)
        {
            return new Button
            {
                Text = "Anexar",
                Location = new Point(x, y - 3),
                Size = new Size(80, 25),
                Font = new Font("Calibri", 9F)
            };
        }

        private void CarregarDados()
        {
            Dictionary<string, string> entidade = new Dictionary<string, string>();
            GetEntidadeJPA(ref entidade);
            if (entidade.Count > 0)
            {
                SetInfoEntidade(entidade);
                CarregarStatusDocumentos();
            }
        }

        private void GetEntidadeJPA(ref Dictionary<string, string> entidade)
        {
            var query = $"SELECT * FROM Geral_Entidade WHERE ID = '{_idJPA}'";
            var dados = _BSO.Consulta(query);

            dados.Inicio();
            if (dados.NumLinhas() > 0)
            {
                string[] colunas = new string[] {
                    "Codigo", "Nome", "NIPC", "CDU_Caminho",
                    "CDU_AnexoFinancas", "CDU_ValidadeFinancas",
                    "CDU_AnexoSegSocial", "CDU_ValidadeSegSocial",
                    "CDU_AnexoFolhaPag", "CDU_ValidadeFolhaPag",
                    "CDU_AnexoComprovativoPagamento", "CDU_ValidadeComprovativoPagamento",
                    "CDU_AnexoReciboSeguroAT", "CDU_ValidadeReciboSeguroAT",
                    "CDU_AnexoSeguroRC", "CDU_ValidadeSeguroRC",
                    "CDU_AnexoSeguroAT", "CDU_ValidadeSeguroAT",
                    "CDU_AnexoAlvara", "CDU_ValidadeAlvara",
                    "CDU_AnexoCertidaoPermanente", "CDU_ValidadeCertidaoPermanente",
                    "CDU_AnexoSeguroResposabilidadeCivil", "CDU_ValidadeSeguroResposabilidadeCivil",
                    "CDU_AnexoAnexoD", "CDU_ValidadeAnexoD"
                };

                foreach (var coluna in colunas)
                {
                    var valor = dados.DaValor<string>(coluna);
                    entidade[coluna] = valor;
                }
            }
        }

        private void SetInfoEntidade(Dictionary<string, string> entidade)
        {
            TXT_Codigo.Text = entidade["Codigo"];
            TXT_Nome.Text = entidade["Nome"];
            TXT_Contribuinte.Text = entidade["NIPC"];
            txtCaminhoPasta.Text = entidade["CDU_Caminho"];
        }

        private void CarregarStatusDocumentos()
        {
            try
            {
                string query = $@"SELECT 
                    CDU_AnexoFinancas, CDU_ValidadeFinancas,
                    CDU_AnexoSegSocial, CDU_ValidadeSegSocial,
                    CDU_AnexoFolhaPag, CDU_ValidadeFolhaPag,
                    CDU_AnexoComprovativoPagamento, CDU_ValidadeComprovativoPagamento,
                    CDU_AnexoReciboSeguroAT, CDU_ValidadeReciboSeguroAT,
                    CDU_AnexoSeguroRC, CDU_ValidadeSeguroRC,
                    CDU_AnexoSeguroAT, CDU_ValidadeSeguroAT,
                    CDU_AnexoAlvara, CDU_ValidadeAlvara,
                    CDU_AnexoCertidaoPermanente, CDU_ValidadeCertidaoPermanente,
                    CDU_AnexoSeguroResposabilidadeCivil, CDU_ValidadeSeguroResposabilidadeCivil,
                    CDU_AnexoAnexoD, CDU_ValidadeAnexoD,
                    CDU_NumApoliceAt, CDU_NumApoliceRc
                    FROM Geral_Entidade WHERE id = '{_idJPA}'";

                var dados = _BSO.Consulta(query);
                if (dados.NumLinhas() > 0)
                {
                    dados.Inicio();
                    UpdateCheckboxFromDB(checkBox1, dados, "CDU_AnexoFinancas", "DND-Finanças", "CDU_ValidadeFinancas");
                    UpdateCheckboxFromDB(checkBox2, dados, "CDU_AnexoSegSocial", "DND-Segurança-Social", "CDU_ValidadeSegSocial");
                    UpdateCheckboxFromDB(checkBox3, dados, "CDU_AnexoFolhaPag", "Mapa de Rem. – SS", "CDU_ValidadeFolhaPag");
                    UpdateCheckboxFromDB(checkBox4, dados, "CDU_AnexoComprovativoPagamento", "TSU", "CDU_ValidadeComprovativoPagamento");
                    UpdateCheckboxFromDB(checkBox5, dados, "CDU_AnexoReciboSeguroAT", "Seguro AT", "CDU_ValidadeReciboSeguroAT");
                    UpdateCheckboxFromDB(checkBox6, dados, "CDU_AnexoSeguroRC", "Seguro RC", "CDU_ValidadeSeguroRC");
                    UpdateCheckboxFromDB(checkBox8, dados, "CDU_AnexoSeguroAT", "Condições Seguro AT", "CDU_ValidadeSeguroAT");
                    UpdateCheckboxFromDB(checkBox9, dados, "CDU_AnexoAlvara", "Alvará", "CDU_ValidadeAlvara");
                    UpdateCheckboxFromDB(checkBox10, dados, "CDU_AnexoCertidaoPermanente", "Certidão Permanente", "CDU_ValidadeCertidaoPermanente");
                    UpdateCheckboxFromDB(checkBox13, dados, "CDU_AnexoSeguroResposabilidadeCivil", "Condições Seguro RC", "CDU_ValidadeSeguroResposabilidadeCivil");
                    UpdateCheckboxFromDB(checkBox28, dados, "CDU_AnexoAnexoD", "Anexo D", "CDU_ValidadeAnexoD");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao carregar documentos: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateCheckboxFromDB(CheckBox checkBox, StdBELista dados, string colunaNome, string tipoDoc, string colunaValidade)
        {
            try
            {
                var valorObj = dados.Valor(colunaNome);
                int anexado = 0;

                if (valorObj is bool valorBool)
                {
                    anexado = valorBool ? 1 : 0;
                }
                else if (valorObj is int valorInt)
                {
                    anexado = valorInt;
                }
                else if (valorObj is string valorStr && !string.IsNullOrEmpty(valorStr))
                {
                    anexado = 1;
                }

                if (anexado == 1)
                {
                    checkBox.Checked = true;
                    checkBox.Enabled = true;

                    DateTime? validade = null;
                    try
                    {
                        // Tentar obter a validade de diferentes formas
                        var valorValidade = dados.Valor(colunaValidade);

                        if (valorValidade is DateTime dataValidade)
                        {
                            validade = dataValidade;
                        }
                        else if (valorValidade is string valorString && !string.IsNullOrEmpty(valorString))
                        {
                            if (DateTime.TryParse(valorString, out DateTime dataConvertida))
                            {
                                validade = dataConvertida;
                            }
                        }
                    }
                    catch { }

                    // Se ainda não encontrou a validade, tentar consultar diretamente
                    if (!validade.HasValue)
                    {
                        try
                        {
                            string queryValidade = $"SELECT {colunaValidade} FROM Geral_Entidade WHERE ID = '{_idJPA}'";
                            var dadosValidade = _BSO.Consulta(queryValidade);
                            if (dadosValidade != null && dadosValidade.NumLinhas() > 0)
                            {
                                dadosValidade.Inicio();
                                var valorValidadeDb = dadosValidade.Valor(colunaValidade);

                                if (valorValidadeDb is DateTime dataValidadeDb)
                                {
                                    validade = dataValidadeDb;
                                }
                                else if (valorValidadeDb is string strValidadeDb && !string.IsNullOrEmpty(strValidadeDb))
                                {
                                    if (DateTime.TryParse(strValidadeDb, out DateTime dataConvertidaDb))
                                    {
                                        validade = dataConvertidaDb;
                                    }
                                }
                            }
                        }
                        catch { }
                    }

                    if (validade.HasValue)
                    {
                        bool dataExpirada = validade.Value < DateTime.Today;
                        checkBox.Text = $"{tipoDoc} (Válido até: {validade.Value.ToShortDateString()})";
                        checkBox.ForeColor = dataExpirada ? Color.Red : SystemColors.ControlText;
                    }
                    else
                    {
                        checkBox.Text = $"{tipoDoc}";
                        checkBox.ForeColor = SystemColors.ControlText;
                    }

                    checkBox.AutoSize = true;
                }
                else
                {
                    checkBox.Text = tipoDoc;
                    checkBox.Checked = false;
                    checkBox.ForeColor = SystemColors.ControlText;
                }
            }
            catch { }
        }

        private void InitializeButtonEvents()
        {
            button1.Click += (sender, e) => AnexarDocumento("DND-Finanças");
            button2.Click += (sender, e) => AnexarDocumento("DND-Segurança-Social");
            button3.Click += (sender, e) => AnexarDocumento("FolhaPagamento");
            button4.Click += (sender, e) => AnexarDocumento("TSU");
            button5.Click += (sender, e) => AnexarDocumento("ReciboSeguroAT");
            button6.Click += (sender, e) => AnexarDocumento("SeguroRC");
            button8.Click += (sender, e) => AnexarDocumento("SeguroAT");
            button9.Click += (sender, e) => AnexarDocumento("Alvara");
            button10.Click += (sender, e) => AnexarDocumento("CertidaoPermanente");
            button12.Click += (sender, e) => AnexarDocumento("SeguroResposabilidadeCivil");
            button13.Click += (sender, e) => AnexarDocumento("AnexoD");
        }

        private void BtnSelecionarPasta_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Selecione a pasta para os documentos da JPA";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtCaminhoPasta.Text = folderDialog.SelectedPath;
                    var update = $@"UPDATE Geral_Entidade SET CDU_Caminho = '{txtCaminhoPasta.Text}' WHERE ID = '{_idJPA}'";
                    _BSO.DSO.ExecuteSQL(update);
                }
            }
        }

        private void AnexarDocumento(string tipoDocumento)
        {
            try
            {
                if (string.IsNullOrEmpty(txtCaminhoPasta.Text) || !Directory.Exists(txtCaminhoPasta.Text))
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

                        Label lblNumeroApolice = new Label
                        {
                            Text = "Número da Apólice RC:",
                            Left = 20,
                            Top = 20,
                            Width = 250
                        };

                        TextBox txtNumeroApolice = new TextBox
                        {
                            Left = 20,
                            Top = 50,
                            Width = 250
                        };

                        Button btnOk = new Button
                        {
                            Text = "OK",
                            DialogResult = DialogResult.OK,
                            Left = 110,
                            Top = 80
                        };

                        formApolice.Controls.Add(lblNumeroApolice);
                        formApolice.Controls.Add(txtNumeroApolice);
                        formApolice.Controls.Add(btnOk);
                        formApolice.AcceptButton = btnOk;

                        if (formApolice.ShowDialog() != DialogResult.OK)
                        {
                            return;
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

                        Label lblNumeroApolice = new Label
                        {
                            Text = "Número da Apólice AT:",
                            Left = 20,
                            Top = 20,
                            Width = 250
                        };

                        TextBox txtNumeroApolice = new TextBox
                        {
                            Left = 20,
                            Top = 50,
                            Width = 250
                        };

                        Button btnOk = new Button
                        {
                            Text = "OK",
                            DialogResult = DialogResult.OK,
                            Left = 110,
                            Top = 80
                        };

                        formApolice.Controls.Add(lblNumeroApolice);
                        formApolice.Controls.Add(txtNumeroApolice);
                        formApolice.Controls.Add(btnOk);
                        formApolice.AcceptButton = btnOk;

                        if (formApolice.ShowDialog() != DialogResult.OK)
                        {
                            return;
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

                        Label lblInfo = new Label
                        {
                            Text = "Informe a data de validade do documento:",
                            Left = 20,
                            Top = 20,
                            Width = 250
                        };

                        DateTimePicker dtpValidade = new DateTimePicker
                        {
                            Left = 20,
                            Top = 50,
                            Width = 250,
                            Format = DateTimePickerFormat.Short,
                            Value = DateTime.Now.AddMonths(1)
                        };

                        Button btnOk = new Button
                        {
                            Text = "OK",
                            DialogResult = DialogResult.OK,
                            Left = 110,
                            Top = 80
                        };

                        formValidade.Controls.Add(lblInfo);
                        formValidade.Controls.Add(dtpValidade);
                        formValidade.Controls.Add(btnOk);
                        formValidade.AcceptButton = btnOk;

                        if (formValidade.ShowDialog() != DialogResult.OK)
                            return;

                        dataValidade = dtpValidade.Value;
                    }
                }

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Title = $"Selecionar {tipoDocumento}";
                    openFileDialog.Filter = "Todos os arquivos (*.*)|*.*|Documentos PDF (*.pdf)|*.pdf";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string sourceFile = openFileDialog.FileName;
                        string nomeArquivo = "JPA";

                        string companyFolder = Path.Combine(txtCaminhoPasta.Text, nomeArquivo);
                        if (!Directory.Exists(companyFolder))
                            Directory.CreateDirectory(companyFolder);

                        string empresaFolder = Path.Combine(companyFolder, "EMPRESA");
                        if (!Directory.Exists(empresaFolder))
                            Directory.CreateDirectory(empresaFolder);

                        string fileName = $"{tipoDocumento.Replace(" ", "_")}_{nomeArquivo}_{DateTime.Now:yyyyMMdd}{Path.GetExtension(sourceFile)}";
                        string destFile = Path.Combine(empresaFolder, fileName);

                        if (File.Exists(destFile))
                        {
                            DialogResult result = MessageBox.Show(
                                $"O arquivo {fileName} já existe. Deseja substituí-lo?",
                                "Arquivo já existe",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (result == DialogResult.No)
                                return;
                        }

                        File.Copy(sourceFile, destFile, true);
                        AtualizarStatusDocumento(tipoDocumento, destFile, dataValidade);

                        // Se for Seguro AT (Condições Seguro AT), atualizar o número de Apólice AT
                        if (tipoDocumento == "SeguroAT" && !string.IsNullOrEmpty(numeroApoliceAt))
                        {
                            string checkColumnQuery = @"
                                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                                               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_NumApoliceAt')
                                BEGIN
                                    ALTER TABLE Geral_Entidade ADD CDU_NumApoliceAt NVARCHAR(50) NULL
                                END";
                            _BSO.DSO.ExecuteSQL(checkColumnQuery);

                            string updateNumApoliceQuery = $@"
                                UPDATE Geral_Entidade 
                                SET CDU_NumApoliceAt = '{numeroApoliceAt}'
                                WHERE ID = '{_idJPA}'";
                            _BSO.DSO.ExecuteSQL(updateNumApoliceQuery);
                        }

                        // Se for Seguro RC (Condições Seguro RC), atualizar o número de Apólice RC
                        if (tipoDocumento == "SeguroResposabilidadeCivil" && !string.IsNullOrEmpty(numeroApoliceRc))
                        {
                            string checkColumnQuery = @"
                                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
                                               WHERE TABLE_NAME = 'Geral_Entidade' AND COLUMN_NAME = 'CDU_NumApoliceRc')
                                BEGIN
                                    ALTER TABLE Geral_Entidade ADD CDU_NumApoliceRc NVARCHAR(50) NULL
                                END";
                            _BSO.DSO.ExecuteSQL(checkColumnQuery);

                            string updateNumApoliceQuery = $@"
                                UPDATE Geral_Entidade 
                                SET CDU_NumApoliceRc = '{numeroApoliceRc}'
                                WHERE ID = '{_idJPA}'";
                            _BSO.DSO.ExecuteSQL(updateNumApoliceQuery);
                        }

                        CarregarStatusDocumentos();

                        MessageBox.Show($"Documento '{tipoDocumento}' anexado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao anexar documento: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AtualizarStatusDocumento(string tipoDocumento, string caminho, DateTime dataValidade)
        {
            try
            {
                string colunaAnexo = "";
                string colunaValidade = "";

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
                    case "SeguroResposabilidadeCivil":
                        colunaAnexo = "CDU_AnexoSeguroResposabilidadeCivil";
                        colunaValidade = "CDU_ValidadeSeguroResposabilidadeCivil";
                        break;
                    case "AnexoD":
                        colunaAnexo = "CDU_AnexoAnexoD";
                        colunaValidade = "CDU_ValidadeAnexoD";
                        break;
                }

                string query = $@"UPDATE Geral_Entidade SET 
                            {colunaAnexo} = 1,
                            {colunaValidade} = '{dataValidade:yyyy-MM-dd}'
                            WHERE Id = '{_idJPA}'";
                _BSO.DSO.ExecuteSQL(query);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao atualizar status: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
