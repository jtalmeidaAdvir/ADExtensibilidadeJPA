using Primavera.Extensibility.BusinessEntities;
using Primavera.Extensibility.CustomForm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class Menu : CustomForm
    {
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


        }

        private void GetEntidades(ref Dictionary<string, string> veiculo)
        {
            string NomeLista = "Entidades";
            string Campos = "Codigo,Nome, NIPC, AlvaraNumero, AlvaraValidade, CDU_NaoDivFinancas, CDU_NaoDivSegSocial, CDU_FolhaPagSegSocial, CDU_ReciboApoliceAT, CDU_ReciboRC, CDU_Caminho, CDU_ReciboPagSegSocial, CDU_ApoliceAT, CDU_ApoliceRC, CDU_HorarioTrabalho, CDU_DecTrabIlegais, CDU_DecRespEstaleiro, CDU_DecConhecimPSS";
            string Tabela = "Geral_Entidade (NOLOCK)";
            string Where = "CDU_TrataSGS = 0";
            string CamposF4 = "Codigo,Nome, NIPC, AlvaraNumero, AlvaraValidade, CDU_NaoDivFinancas, CDU_NaoDivSegSocial, CDU_FolhaPagSegSocial, CDU_ReciboApoliceAT, CDU_ReciboRC, CDU_Caminho, CDU_ReciboPagSegSocial, CDU_ApoliceAT, CDU_ApoliceRC, CDU_HorarioTrabalho, CDU_DecTrabIlegais, CDU_DecRespEstaleiro, CDU_DecConhecimPSS";
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
    }
}
