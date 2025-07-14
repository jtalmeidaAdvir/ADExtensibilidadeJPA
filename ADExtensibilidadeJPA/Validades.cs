using ErpBS100;
using StdPlatBS100;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace ADExtensibilidadeJPA
{
    public partial class Validades : Form
    {
        private readonly ErpBS _BSO;
        private readonly StdBSInterfPub _PSO;
        private readonly List<string> _idsSelecionados;

        public Validades(ErpBS bSO, StdBSInterfPub pSO, List<string> idsSelecionados)
        {
            InitializeComponent();
            _BSO = bSO;
            _PSO = pSO;
            _idsSelecionados = idsSelecionados;
            LoadDocumentos();
        }

        private TabControl _tabControl;

        private void LoadDocumentos()
        {
            // Criar botões
            Button btnExpandir = new Button { Text = "Expandir Tudo", Dock = DockStyle.Bottom, Height = 30 };
            Button btnRecolher = new Button { Text = "Recolher Tudo", Dock = DockStyle.Bottom, Height = 30 };

            btnExpandir.Click += BtnExpandir_Click;
            btnRecolher.Click += BtnRecolher_Click;

            this.Controls.Add(btnRecolher);
            this.Controls.Add(btnExpandir);

            // Criar TabControl
            _tabControl = new TabControl();
            _tabControl.Dock = DockStyle.Fill;
            this.Controls.Add(_tabControl);

            // Adicionar tabs
            TabPage tabEmpresa = new TabPage("Documentos Empresa");
            TabPage tabTrabalhadores = new TabPage("Documentos Trabalhadores");
            TabPage tabEquipamentos = new TabPage("Documentos Equipamentos");
            TabPage tabAutorizacoes = new TabPage("Documentos Autorizações");

            _tabControl.TabPages.Add(tabEmpresa);
            _tabControl.TabPages.Add(tabTrabalhadores);
            _tabControl.TabPages.Add(tabEquipamentos);
            _tabControl.TabPages.Add(tabAutorizacoes);

            // Carregar conteúdo
            LoadDocumentosEmpresa(tabEmpresa);
            LoadDocumentosTrabalhadores(tabTrabalhadores);
            LoadDocumentosEquipamentos(tabEquipamentos);
            LoadDocumentosAutorizacoes(tabAutorizacoes);
        }

        private void BtnExpandir_Click(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab?.Controls.OfType<TreeView>().FirstOrDefault() is TreeView tv)
            {
                tv.ExpandAll();
            }
        }

        private void BtnRecolher_Click(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab?.Controls.OfType<TreeView>().FirstOrDefault() is TreeView tv)
            {
                tv.CollapseAll();
            }
        }

        private void LoadDocumentosEmpresa(TabPage tab)
        {
            TreeView treeView = new TreeView();
            treeView.Dock = DockStyle.Fill;
            treeView.Font = new Font("Segoe UI", 10);
            treeView.ShowLines = true;

            // Ajuste: Criando uma lista de IDs selecionados
            string ids = string.Join(",", _idsSelecionados.Select(id => $"'{id}'"));

            // Consulta ajustada para garantir que os documentos de múltiplas entidades sejam carregados
            var query = $@"
SELECT 
    Nome AS Entidade,
    'Finanças' AS Documento, CDU_ValidadeFinancas AS Validade,
    CASE WHEN CDU_ValidadeFinancas < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeFinancas IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Segurança Social' AS Documento, CDU_ValidadeSegSocial AS Validade,
    CASE WHEN CDU_ValidadeSegSocial < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeSegSocial IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Folha de Pagamento' AS Documento, CDU_ValidadeFolhaPag AS Validade,
    CASE WHEN CDU_ValidadeFolhaPag < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeFolhaPag IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Comprovativo de Pagamento' AS Documento, CDU_ValidadeComprovativoPagamento AS Validade,
    CASE WHEN CDU_ValidadeComprovativoPagamento < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeComprovativoPagamento IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Recibo Seguro de Acidentes de Trabalho' AS Documento, CDU_ValidadeReciboSeguroAT AS Validade,
    CASE WHEN CDU_ValidadeReciboSeguroAT < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeReciboSeguroAT IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Seguro de Responsabilidade Civil' AS Documento, CDU_ValidadeSeguroRC AS Validade,
    CASE WHEN CDU_ValidadeSeguroRC < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeSeguroRC IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Horário de Trabalho' AS Documento, CDU_ValidadeHorarioTrabalho AS Validade,
    CASE WHEN CDU_ValidadeHorarioTrabalho < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeHorarioTrabalho IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Seguro de Acidentes de Trabalho' AS Documento, CDU_ValidadeSeguroAT AS Validade,
    CASE WHEN CDU_ValidadeSeguroAT < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeSeguroAT IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Alvará' AS Documento, CDU_ValidadeAlvara AS Validade,
    CASE WHEN CDU_ValidadeAlvara < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeAlvara IS NOT NULL

UNION ALL

SELECT 
    Nome AS Entidade,
    'Certidão Permanente' AS Documento, CDU_ValidadeCertidaoPermanente AS Validade,
    CASE WHEN CDU_ValidadeCertidaoPermanente < GETDATE() THEN 'Expirado' ELSE 'Válido' END AS Estado
FROM Geral_Entidade 
WHERE id IN ({ids}) AND CDU_ValidadeCertidaoPermanente IS NOT NULL
";



            var dados = _BSO.Consulta(query);
            dados.Inicio();

            Dictionary<string, TreeNode> entidadesProcessadas = new Dictionary<string, TreeNode>();  // Dicionário para armazenar as entidades processadas

            while (!dados.NoFim())
            {
                string entidade = dados.Valor("Entidade").ToString();
                string documento = dados.Valor("Documento").ToString();
                DateTime validade = Convert.ToDateTime(dados.Valor("Validade"));
                string estado = dados.Valor("Estado").ToString();

                // Verifica se o nó da entidade já foi criado
                if (!entidadesProcessadas.ContainsKey(entidade))
                {
                    // Se a entidade não foi criada, cria o nó e adiciona ao dicionário
                    TreeNode entidadeNode = new TreeNode(entidade);
                    entidadeNode.NodeFont = new Font("Segoe UI", 10, FontStyle.Bold);
                    entidadesProcessadas[entidade] = entidadeNode;  // Adiciona ao dicionário
                    treeView.Nodes.Add(entidadeNode);  // Adiciona o nó ao TreeView
                }

                // Agora, adiciona o documento como um nó filho da entidade
                TreeNode docNode = new TreeNode($"{documento} - {validade:dd/MM/yyyy} ({estado})");
                docNode.ForeColor = estado == "Expirado" ? Color.Red : Color.Green;

                // Adiciona o documento ao nó da entidade correspondente
                entidadesProcessadas[entidade].Nodes.Add(docNode);

                dados.Seguinte();
            }

            treeView.NodeMouseDoubleClick += (s, e) =>
            {
                if (e.Node.Level == 1) // Documento
                {
                    string doc = e.Node.Text.Split('-')[0].Trim();
                    // UpdateDocumento(doc, "empresa");
                }
            };

            tab.Controls.Add(treeView);
        }

        private void LoadDocumentosTrabalhadores(TabPage tab)
        {
            TreeView treeView = new TreeView();
            treeView.Dock = DockStyle.Fill;
            treeView.Font = new Font("Segoe UI", 10);
            treeView.ShowLines = true;

            string ids = string.Join(",", _idsSelecionados.Select(id => $"'{id}'"));
            bool agruparPorEmpresa = _idsSelecionados.Count > 1;

            string[] nomesDocumentos = new string[]
            {
        "Cartão de Cidadão",
        "Ficha Médica de Aptidão",
        "Credenciação do Trabalhador",
        "Trabalhos Especializados",
        "Ficha de Distribuição de EPI´s"
            };

            var query = $@"
    SELECT 
        e.Nome AS Empresa,
        t.nome AS Trabalhador,
        t.caminho1, t.caminho2, t.caminho3, t.caminho4, t.caminho5
    FROM TDU_AD_Trabalhadores t
    INNER JOIN Geral_Entidade e ON t.id_empresa = e.id
    WHERE t.id_empresa IN ({ids})";

            var dados = _BSO.Consulta(query);
            dados.Inicio();

            Dictionary<string, TreeNode> empresasNodes = new Dictionary<string, TreeNode>();

            while (!dados.NoFim())
            {
                string empresa = dados.Valor("Empresa").ToString();
                string trabalhador = dados.Valor("Trabalhador").ToString();

                TreeNode trabalhadorNode = new TreeNode(trabalhador)
                {
                    NodeFont = new Font("Segoe UI", 10, FontStyle.Regular)
                };

                for (int i = 1; i <= 5; i++)
                {
                    string caminhoValue = dados.Valor($"caminho{i}")?.ToString();
                    if (!string.IsNullOrEmpty(caminhoValue))
                    {
                        var dataMatch = System.Text.RegularExpressions.Regex.Match(caminhoValue, @"Válido até&#58; (\d{2}/\d{2}/\d{4})");
                        if (dataMatch.Success)
                        {
                            if (DateTime.TryParse(dataMatch.Groups[1].Value, out DateTime validade))
                            {
                                string estado = validade < DateTime.Today ? "Expirado" : "Válido";
                                string nomeDocumento = nomesDocumentos[i - 1];
                                string textoNode = $"{nomeDocumento} - {validade:dd/MM/yyyy} ({estado})";

                                TreeNode docNode = new TreeNode(textoNode)
                                {
                                    ForeColor = estado == "Expirado" ? Color.Red : Color.Green
                                };
                                trabalhadorNode.Nodes.Add(docNode);
                            }
                        }
                    }
                }

                if (trabalhadorNode.Nodes.Count > 0)
                {
                    if (agruparPorEmpresa)
                    {
                        if (!empresasNodes.TryGetValue(empresa, out TreeNode empresaNode))
                        {
                            empresaNode = new TreeNode(empresa)
                            {
                                NodeFont = new Font("Segoe UI", 10, FontStyle.Bold)
                            };
                            empresasNodes[empresa] = empresaNode;
                            treeView.Nodes.Add(empresaNode);
                        }
                        empresaNode.Nodes.Add(trabalhadorNode);
                    }
                    else
                    {
                        treeView.Nodes.Add(trabalhadorNode);
                    }
                }

                dados.Seguinte();
            }

            treeView.NodeMouseDoubleClick += (s, e) =>
            {
                if ((_idsSelecionados.Count > 1 && e.Node.Level == 2) || (_idsSelecionados.Count == 1 && e.Node.Level == 1))
                {
                    string doc = e.Node.Text.Split('-')[0].Trim();
                    string trab = (_idsSelecionados.Count > 1) ? e.Node.Parent.Text : e.Node.Parent.Text;
                    // string empresa = (_idsSelecionados.Count > 1) ? e.Node.Parent.Parent.Text : _nomeDaEmpresa; // Se quiseres usar nome da empresa
                    // UpdateDocumento(doc, "trabalhador", trab);
                }
            };

            tab.Controls.Add(treeView);
        }

        private void LoadDocumentosEquipamentos(TabPage tab)
        {
            TreeView treeView = new TreeView();
            treeView.Dock = DockStyle.Fill;
            treeView.Font = new Font("Segoe UI", 10);
            treeView.ShowLines = true;

            string ids = string.Join(",", _idsSelecionados.Select(id => $"'{id}'"));
            bool agruparPorEmpresa = _idsSelecionados.Count > 1;

            var query = $@"
    SELECT 
        e.Nome AS Empresa,
        eq.marca AS Equipamento,
        eq.caminho5
    FROM TDU_AD_Equipamentos eq
    INNER JOIN Geral_Entidade e ON eq.id_empresa = e.id
    WHERE eq.id_empresa IN ({ids})";

            var dados = _BSO.Consulta(query);
            dados.Inicio();

            Dictionary<string, TreeNode> empresasNodes = new Dictionary<string, TreeNode>();

            while (!dados.NoFim())
            {
                string empresa = dados.Valor("Empresa").ToString();
                string equipamento = dados.Valor("Equipamento").ToString();

                TreeNode equipamentoNode = new TreeNode(equipamento)
                {
                    NodeFont = new Font("Segoe UI", 10, FontStyle.Regular)
                };

                string caminhoValue = dados.Valor("caminho5")?.ToString();

                if (!string.IsNullOrEmpty(caminhoValue))
                {
                    var dataMatch = System.Text.RegularExpressions.Regex.Match(caminhoValue, @"Válido até&#58; (\d{2}/\d{2}/\d{4})");
                    if (dataMatch.Success)
                    {
                        if (DateTime.TryParse(dataMatch.Groups[1].Value, out DateTime validade))
                        {
                            string estado = validade < DateTime.Today ? "Expirado" : "Válido";
                            string textoNode = $"Seguro - {validade:dd/MM/yyyy} ({estado})";

                            TreeNode docNode = new TreeNode(textoNode)
                            {
                                ForeColor = estado == "Expirado" ? Color.Red : Color.Green
                            };
                            equipamentoNode.Nodes.Add(docNode);
                        }
                    }
                }

                if (equipamentoNode.Nodes.Count > 0)
                {
                    if (agruparPorEmpresa)
                    {
                        if (!empresasNodes.TryGetValue(empresa, out TreeNode empresaNode))
                        {
                            empresaNode = new TreeNode(empresa)
                            {
                                NodeFont = new Font("Segoe UI", 10, FontStyle.Bold)
                            };
                            empresasNodes[empresa] = empresaNode;
                            treeView.Nodes.Add(empresaNode);
                        }
                        empresaNode.Nodes.Add(equipamentoNode);
                    }
                    else
                    {
                        treeView.Nodes.Add(equipamentoNode);
                    }
                }

                dados.Seguinte();
            }

            treeView.NodeMouseDoubleClick += (s, e) =>
            {
                if ((_idsSelecionados.Count > 1 && e.Node.Level == 2) || (_idsSelecionados.Count == 1 && e.Node.Level == 1))
                {
                    string doc = "Seguro";
                    string equipamento = (_idsSelecionados.Count > 1) ? e.Node.Parent.Text : e.Node.Parent.Text;
                    // UpdateDocumento(doc, "equipamento", equipamento);
                }
            };

            tab.Controls.Add(treeView);
        }

        private void LoadDocumentosAutorizacoes(TabPage tab)
        {
            TreeView treeView = new TreeView();
            treeView.Dock = DockStyle.Fill;
            treeView.Font = new Font("Segoe UI", 10);
            treeView.ShowLines = true;

            string ids = string.Join(",", _idsSelecionados.Select(id => $"'{id}'"));
            bool agruparPorEntidade = _idsSelecionados.Count > 1;

            string[] nomesDocumentos = new string[]
            {
        "Contrato ou Nota de Encomenda",
        "Horário de Trabalho da Empreitada",
        "Declaração de Adesão ao PSS",
        "Declaração do Responsável no Estaleiro"
            };

            var query = $@"
    SELECT 
        e.Nome AS Entidade,
        a.Nome_Obra AS Obra,
        a.caminho1, a.caminho2, a.caminho3, a.caminho4
    FROM TDU_AD_Autorizacoes a
    INNER JOIN Geral_Entidade e ON a.ID_Entidade = e.ID
    WHERE a.ID_Entidade IN ({ids})";

            var dados = _BSO.Consulta(query);
            dados.Inicio();

            Dictionary<string, TreeNode> entidadesNodes = new Dictionary<string, TreeNode>();

            while (!dados.NoFim())
            {
                string entidade = dados.Valor("Entidade").ToString();
                string obra = dados.Valor("Obra").ToString();

                TreeNode obraNode = new TreeNode(obra)
                {
                    NodeFont = new Font("Segoe UI", 10, FontStyle.Regular)
                };

                for (int i = 1; i <= 4; i++)
                {
                    string caminhoValue = dados.Valor($"caminho{i}")?.ToString();
                    if (!string.IsNullOrEmpty(caminhoValue))
                    {
                        var dataMatch = System.Text.RegularExpressions.Regex.Match(caminhoValue, @"Válido até&#58; (\d{2}/\d{2}/\d{4})");
                        if (dataMatch.Success)
                        {
                            if (DateTime.TryParse(dataMatch.Groups[1].Value, out DateTime validade))
                            {
                                string estado = validade < DateTime.Today ? "Expirado" : "Válido";
                                string nomeDocumento = nomesDocumentos[i - 1];
                                string textoNode = $"{nomeDocumento} - {validade:dd/MM/yyyy} ({estado})";

                                TreeNode docNode = new TreeNode(textoNode)
                                {
                                    ForeColor = estado == "Expirado" ? Color.Red : Color.Green
                                };
                                obraNode.Nodes.Add(docNode);
                            }
                        }
                    }
                }

                if (obraNode.Nodes.Count > 0)
                {
                    if (agruparPorEntidade)
                    {
                        if (!entidadesNodes.TryGetValue(entidade, out TreeNode entidadeNode))
                        {
                            entidadeNode = new TreeNode(entidade)
                            {
                                NodeFont = new Font("Segoe UI", 10, FontStyle.Bold)
                            };
                            entidadesNodes[entidade] = entidadeNode;
                            treeView.Nodes.Add(entidadeNode);
                        }
                        entidadeNode.Nodes.Add(obraNode);
                    }
                    else
                    {
                        treeView.Nodes.Add(obraNode);
                    }
                }

                dados.Seguinte();
            }

            treeView.NodeMouseDoubleClick += (s, e) =>
            {
                bool isDocumento = (_idsSelecionados.Count > 1 && e.Node.Level == 2) || (_idsSelecionados.Count == 1 && e.Node.Level == 1);
                if (isDocumento)
                {
                    string doc = e.Node.Text.Split('-')[0].Trim();
                    string obra = (_idsSelecionados.Count > 1) ? e.Node.Parent.Text : e.Node.Parent.Text;
                    // UpdateDocumento(doc, "autorizacao", obra);
                }
            };

            tab.Controls.Add(treeView);
        }

        private void UpdateDocumento(string doc, string tipo, string referencia)
        {
            // Exemplo de implementação para atualizações quando um documento é selecionado.
            MessageBox.Show($"Documento: {doc}\nTipo: {tipo}\nReferência: {referencia}", "Atualização de Documento");
        }
    }
}

