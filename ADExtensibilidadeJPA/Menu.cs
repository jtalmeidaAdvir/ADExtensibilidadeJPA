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
        }

        private void BTF4_Click(object sender, EventArgs e)
        {
            MetodoGetEntidades();
        }

        private void MetodoGetEntidades()
        {
            Dictionary<string, string> veiculo = new Dictionary<string, string>();
            GetEntidades(ref veiculo);

            if (veiculo.Count > 0)
            {
                SetInfoEntidades(veiculo);
            }
        }


        private void SetInfoEntidades(Dictionary<string, string> veiculo)
        {
            _ID = veiculo["id"]; 
            TXT_Codigo.Text = veiculo["Codigo"];
            TXT_Nome.Text = veiculo["Nome"];
            TXT_Contribuinte.Text = veiculo["NIPC"];
            TXT_Alvara.Text = veiculo["AlvaraNumero"];
            TXT_AlvaraValidade.Text = veiculo["AlvaraValidade"];

            TXT_NaoDivFinancas.Text = veiculo["CDU_NaoDivFinancas"];
            TXT_NaoDivSegSocial.Text = veiculo["CDU_NaoDivSegSocial"];
            TXT_FolhaPagSegSocial.Text = veiculo["CDU_FolhaPagSegSocial"];
            TXT_ReciboApoliceAT.Text = veiculo["CDU_ReciboApoliceAT"];
            TXT_ReciboRC.Text = veiculo["CDU_ReciboRC"];



            // Recupera os valores do banco de dados
            string reciboPagSegSocial = veiculo["CDU_ReciboPagSegSocial"];
            string apoliceAT = veiculo["CDU_ApoliceAT"];
            string apoliceRC = veiculo["CDU_ApoliceRC"];
            string horarioTrabalho = veiculo["CDU_HorarioTrabalho"];
            string decTrabIlegais = veiculo["CDU_DecTrabIlegais"];
            string decRespEstaleiro = veiculo["CDU_DecRespEstaleiro"];
            string decConhecimPSS = veiculo["CDU_DecConhecimPSS"];


            PreencherComboBox(cb_ReciboPagSegSocial, reciboPagSegSocial);
            PreencherComboBox(cb_ApoliceAT, apoliceAT);
            PreencherComboBox(cb_ApoliceRC, apoliceRC);
            PreencherComboBox(cb_HorarioTrabalho, horarioTrabalho);
            PreencherComboBox(cb_DecTrabIlegais, decTrabIlegais);
            PreencherComboBox(cb_DecRespEstaleiro, decRespEstaleiro);
            PreencherComboBox(cb_DecConhecimPSS, decConhecimPSS);


            var moradaCompleta = $"{veiculo["Morada"]}, {veiculo["Localidade"]}, {veiculo["CodPostal"]}, {veiculo["CodPostalLocal"]}";


            if (moradaCompleta == ", , , ")
            {
                moradaCompleta = "";
            }
            else
            {
                TXT_Sede.Text = moradaCompleta;
            }

            carregarObrasCB(veiculo);
        }

        private void carregarObrasCB(Dictionary<string, string> veiculo)
        {
            var BDObras = GetObrasSumbempreiteiro(veiculo["EntidadeId"]);


            var numLinhasObras = BDObras.NumLinhas();
            for (int i = 0; i < numLinhasObras; i++)
            {
                var txt = BDObras.DaValor<string>("Codigo") + " - " + BDObras.DaValor<string>("Descricao");
                cb_Obras.Items.Add(BDObras.DaValor<string>("Codigo"));
                BDObras.Seguinte();
            }
        }

        private StdBELista GetObrasSumbempreiteiro(string v)
        {
            var query = $@"SELECT * FROM COP_Obras 
                            where Tipo = 'S' AND EntidadeIDA = '{v}'";
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
        private void GetEntidades(ref Dictionary<string, string> veiculo)
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
                        veiculo[colunas[i].Trim()] = ResQuery[i].ToString();
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

        private void cb_Obras_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Obras.SelectedItem != null)
            {
                // Obtenha o código da obra selecionada
                string codigoObraSelecionada = cb_Obras.SelectedItem.ToString();

                // Aqui você pode usar o código conforme necessário
                MessageBox.Show($"Código da obra selecionada: {codigoObraSelecionada}");
                var queryGetObras = $@"SELECT * FROM COP_Obras 
                                        WHERE Codigo = '{codigoObraSelecionada}'";

                var DBObras = BSO.Consulta(queryGetObras);


                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.Cells["EntradaObra_"].Value = DBObras.DaValor<string>("CDU_EntradaObra");
                    row.Cells["SaidaObra_"].Value = DBObras.DaValor<string>("CDU_SaidaObra");
                    row.Cells["ContratoSubempreitada"].Value = DBObras.DaValor<string>("CDU_ContratoSubempreitada");
                    row.Cells["AutorizacaoEntrada"].Value = DBObras.DaValor<int>("CDU_AutorizacaoEntrada") == 1;

                    row.DefaultCellStyle.BackColor = Color.LightYellow;

                    break;

                }
            }
        }

        private void BT_Salvar_Click_Click(object sender, EventArgs e)
        {
            var querySalvar = $@"
                                UPDATE Geral_Entidade
                                SET 
                                    NIPC = '{TXT_Contribuinte.Text}', 
                                    AlvaraNumero = '{TXT_Alvara.Text}', 
                                    AlvaraValidade = '{TXT_AlvaraValidade.Text}', 
                                    CDU_NaoDivFinancas = '{TXT_NaoDivFinancas.Text}', 
                                    CDU_NaoDivSegSocial = '{TXT_NaoDivSegSocial.Text}', 
                                    CDU_FolhaPagSegSocial = '{TXT_FolhaPagSegSocial.Text}', 
                                    CDU_ReciboApoliceAT = '{TXT_ReciboApoliceAT.Text}', 
                                    CDU_ReciboRC = '{TXT_ReciboRC.Text}', 
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
        }
    }
}
