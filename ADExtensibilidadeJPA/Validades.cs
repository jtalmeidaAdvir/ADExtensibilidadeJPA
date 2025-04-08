
using ErpBS100;
using StdPlatBS100;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace ADExtensibilidadeJPA
{
    public partial class Validades : Form
    {
        private readonly ErpBS _BSO;
        private readonly StdBSInterfPub _PSO;
        private readonly string _idSelecionado;

        public Validades(ErpBS bSO, StdBSInterfPub pSO, string idSelecionado)
        {
            InitializeComponent();
            _BSO = bSO;
            _PSO = pSO;
            _idSelecionado = idSelecionado;
            LoadDocumentos();
        }

        private void LoadDocumentos()
        {
            // Create TabControl
            TabControl tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;
            this.Controls.Add(tabControl);

            // Add tabs
            TabPage tabEmpresa = new TabPage("Documentos Empresa");
            TabPage tabTrabalhadores = new TabPage("Documentos Trabalhadores");
            TabPage tabEquipamentos = new TabPage("Documentos Equipamentos");
            TabPage tabAutorizacoes = new TabPage("Documentos Autorizações");

            tabControl.TabPages.Add(tabEmpresa);
            tabControl.TabPages.Add(tabTrabalhadores);
            tabControl.TabPages.Add(tabEquipamentos);
            tabControl.TabPages.Add(tabAutorizacoes);

            // Load data for each tab
            LoadDocumentosEmpresa(tabEmpresa);
            LoadDocumentosTrabalhadores(tabTrabalhadores);
            LoadDocumentosEquipamentos(tabEquipamentos);
            LoadDocumentosAutorizacoes(tabAutorizacoes);
        }

        private void LoadDocumentosEmpresa(TabPage tab)
        {
            DataGridView grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AllowUserToAddRows = false;
            grid.MultiSelect = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            grid.Columns.Add("Documento", "Documento");
            grid.Columns.Add("Validade", "Validade");
            grid.Columns.Add("Estado", "Estado");

            var query = $@"SELECT 
                'Finanças' as Documento, CDU_ValidadeFinancas as Validade,
                CASE WHEN CDU_ValidadeFinancas < GETDATE() THEN 'Expirado' ELSE 'Válido' END as Estado
                FROM Geral_Entidade WHERE id = '{_idSelecionado}'
                AND CDU_ValidadeFinancas IS NOT NULL
                UNION ALL
                SELECT 'Segurança Social', CDU_ValidadeSegSocial,
                CASE WHEN CDU_ValidadeSegSocial < GETDATE() THEN 'Expirado' ELSE 'Válido' END
                FROM Geral_Entidade WHERE id = '{_idSelecionado}'
                AND CDU_ValidadeSegSocial IS NOT NULL";

            var dados = _BSO.Consulta(query);
            dados.Inicio();

            while (!dados.NoFim())
            {
                var doc = dados.Valor("Documento").ToString();
                var val = dados.Valor("Validade");
                var estado = dados.Valor("Estado").ToString();

                int index = grid.Rows.Add(doc, val, estado);
                if (estado == "Expirado")
                    grid.Rows[index].DefaultCellStyle.BackColor = Color.LightPink;

                dados.Seguinte();
            }

            grid.CellDoubleClick += (s, e) => {
                if (e.RowIndex >= 0)
                {
                    var doc = grid.Rows[e.RowIndex].Cells["Documento"].Value.ToString();
                    UpdateDocumento(doc, "empresa");
                }
            };

            tab.Controls.Add(grid);
        }

        private void LoadDocumentosTrabalhadores(TabPage tab)
        {
            DataGridView grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AllowUserToAddRows = false;

            grid.Columns.Add("Trabalhador", "Trabalhador");
            grid.Columns.Add("Documento", "Documento");
            grid.Columns.Add("Validade", "Validade");
            grid.Columns.Add("Estado", "Estado");

            var query = $@"SELECT 
                nome as Trabalhador,
                caminho1, caminho2, caminho3, caminho4, caminho5
                FROM TDU_AD_Trabalhadores 
                WHERE id_empresa = '{_idSelecionado}'";

            var dados = _BSO.Consulta(query);
            dados.Inicio();

            while (!dados.NoFim())
            {
                var trabalhador = dados.Valor("Trabalhador").ToString();
                for (int i = 1; i <= 5; i++)
                {
                    var caminhoValue = dados.Valor($"caminho{i}")?.ToString();
                    if (!string.IsNullOrEmpty(caminhoValue))
                    {
                        var dataMatch = System.Text.RegularExpressions.Regex.Match(caminhoValue, @"Válido até&#58; (\d{2}/\d{2}/\d{4})");
                        if (dataMatch.Success)
                        {
                            DateTime validade;
                            if (DateTime.TryParse(dataMatch.Groups[1].Value, out validade))
                            {
                                string estado = validade < DateTime.Today ? "Expirado" : "Válido";
                                int index = grid.Rows.Add(trabalhador, $"Documento {i}", validade.ToShortDateString(), estado);
                                if (estado == "Expirado")
                                    grid.Rows[index].DefaultCellStyle.BackColor = Color.LightPink;
                            }
                        }
                    }
                }
                dados.Seguinte();
            }

            grid.CellDoubleClick += (s, e) => {
                if (e.RowIndex >= 0)
                {
                    var trab = grid.Rows[e.RowIndex].Cells["Trabalhador"].Value.ToString();
                    var doc = grid.Rows[e.RowIndex].Cells["Documento"].Value.ToString();
                    UpdateDocumento(doc, "trabalhador", trab);
                }
            };

            tab.Controls.Add(grid);
        }

        private void LoadDocumentosEquipamentos(TabPage tab)
        {
            // Similar implementation to LoadDocumentosTrabalhadores but for equipamentos
            DataGridView grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AllowUserToAddRows = false;

            grid.Columns.Add("Equipamento", "Equipamento");
            grid.Columns.Add("Documento", "Documento");
            grid.Columns.Add("Validade", "Validade");
            grid.Columns.Add("Estado", "Estado");

            var query = $@"SELECT 
                marca as Equipamento,
                caminho5
                FROM TDU_AD_Equipamentos 
                WHERE id_empresa = '{_idSelecionado}'";

            var dados = _BSO.Consulta(query);
            dados.Inicio();

            while (!dados.NoFim())
            {
                var equipamento = dados.Valor("Equipamento").ToString();
                var caminhoValue = dados.Valor("caminho5")?.ToString();

                if (!string.IsNullOrEmpty(caminhoValue))
                {
                    var dataMatch = System.Text.RegularExpressions.Regex.Match(caminhoValue, @"Válido até&#58; (\d{2}/\d{2}/\d{4})");
                    if (dataMatch.Success)
                    {
                        DateTime validade;
                        if (DateTime.TryParse(dataMatch.Groups[1].Value, out validade))
                        {
                            string estado = validade < DateTime.Today ? "Expirado" : "Válido";
                            int index = grid.Rows.Add(equipamento, "Seguro", validade.ToShortDateString(), estado);
                            if (estado == "Expirado")
                                grid.Rows[index].DefaultCellStyle.BackColor = Color.LightPink;
                        }
                    }
                }
                dados.Seguinte();
            }

            grid.CellDoubleClick += (s, e) => {
                if (e.RowIndex >= 0)
                {
                    var equip = grid.Rows[e.RowIndex].Cells["Equipamento"].Value.ToString();
                    UpdateDocumento("Seguro", "equipamento", equip);
                }
            };

            tab.Controls.Add(grid);
        }

        private void LoadDocumentosAutorizacoes(TabPage tab)
        {
            DataGridView grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AllowUserToAddRows = false;

            grid.Columns.Add("Obra", "Obra");
            grid.Columns.Add("Documento", "Documento");
            grid.Columns.Add("Validade", "Validade");
            grid.Columns.Add("Estado", "Estado");

            var query = $@"SELECT 
                Nome_Obra as Obra,
                caminho1, caminho2, caminho3, caminho4
                FROM TDU_AD_Autorizacoes 
                WHERE ID_Entidade = '{_idSelecionado}'";

            var dados = _BSO.Consulta(query);
            dados.Inicio();

            while (!dados.NoFim())
            {
                var obra = dados.Valor("Obra").ToString();
                for (int i = 1; i <= 4; i++)
                {
                    var caminhoValue = dados.Valor($"caminho{i}")?.ToString();
                    if (!string.IsNullOrEmpty(caminhoValue))
                    {
                        var dataMatch = System.Text.RegularExpressions.Regex.Match(caminhoValue, @"Válido até&#58; (\d{2}/\d{2}/\d{4})");
                        if (dataMatch.Success)
                        {
                            DateTime validade;
                            if (DateTime.TryParse(dataMatch.Groups[1].Value, out validade))
                            {
                                string estado = validade < DateTime.Today ? "Expirado" : "Válido";
                                int index = grid.Rows.Add(obra, $"Documento {i}", validade.ToShortDateString(), estado);
                                if (estado == "Expirado")
                                    grid.Rows[index].DefaultCellStyle.BackColor = Color.LightPink;
                            }
                        }
                    }
                }
                dados.Seguinte();
            }

            grid.CellDoubleClick += (s, e) => {
                if (e.RowIndex >= 0)
                {
                    var obra = grid.Rows[e.RowIndex].Cells["Obra"].Value.ToString();
                    var doc = grid.Rows[e.RowIndex].Cells["Documento"].Value.ToString();
                    UpdateDocumento(doc, "autorizacao", obra);
                }
            };

            tab.Controls.Add(grid);
        }

        private void UpdateDocumento(string documento, string tipo, string identificador = "")
        {
            using (Form formValidade = new Form())
            {
                formValidade.Text = "Nova Data de Validade";
                formValidade.Width = 300;
                formValidade.Height = 150;
                formValidade.StartPosition = FormStartPosition.CenterParent;
                formValidade.FormBorderStyle = FormBorderStyle.FixedDialog;
                formValidade.MaximizeBox = false;
                formValidade.MinimizeBox = false;

                DateTimePicker dtpValidade = new DateTimePicker();
                dtpValidade.Format = DateTimePickerFormat.Short;
                dtpValidade.Location = new Point(20, 20);
                dtpValidade.Width = 250;

                Button btnOk = new Button();
                btnOk.Text = "Atualizar";
                btnOk.DialogResult = DialogResult.OK;
                btnOk.Location = new Point(100, 60);

                formValidade.Controls.AddRange(new Control[] { dtpValidade, btnOk });

                if (formValidade.ShowDialog() == DialogResult.OK)
                {
                    // Atualizar o documento dependendo do tipo
                    switch (tipo)
                    {
                        case "empresa":
                            // Atualizar documento da empresa
                            break;
                        case "trabalhador":
                            // Atualizar documento do trabalhador
                            break;
                        case "equipamento":
                            // Atualizar documento do equipamento
                            break;
                        case "autorizacao":
                            // Atualizar documento da autorização
                            break;
                    }

                    // Recarregar os dados
                    LoadDocumentos();
                }
            }
        }
    }
}
