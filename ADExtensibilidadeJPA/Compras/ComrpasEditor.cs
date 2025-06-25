

using DocumentFormat.OpenXml.Bibliography;
using Primavera.Extensibility.BusinessEntities;
using Primavera.Extensibility.BusinessEntities.ExtensibilityService.EventArgs;
using Primavera.Extensibility.Purchases.Editors;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ADExtensibilidadeJPA.Compras
{
    public class ComrpasEditor : EditorCompras
    {

        public override void ArtigoIdentificado(string Artigo, int NumLinha, ref bool Cancel, ExtensibilityEventArgs e)
        {
            if (this.DocumentoCompra.Tipodoc == "VFA" || this.DocumentoCompra.Tipodoc == "VFAO")
            {
                var query = $@"SELECT PVP1 FROM ArtigoMoeda WHERE Artigo = '{Artigo}'";
                var data = BSO.Consulta(query);


                this.DocumentoCompra.Linhas.GetEdita(NumLinha).CamposUtil["CDU_PVP1"].Valor = data.DaValor<string>("PVP1");
            }
        }
        public override void DepoisDeGravar(string Filial, string Tipo, string Serie, int NumDoc, ExtensibilityEventArgs e)
        {
            if(this.DocumentoCompra.Tipodoc == "VFA" || this.DocumentoCompra.Tipodoc == "VFAO")
            {
                var num = this.DocumentoCompra.Linhas.NumItens;


                for (int i = 1; i < num + 1; i++)
                {
                    var linha = this.DocumentoCompra.Linhas.GetEdita(i);


                    var pvp1 = Convert.ToSingle(linha.CamposUtil["CDU_PVP1"].Valor);
                    var update = $@"UPDATE ArtigoMoeda
                SET PVP1 = {pvp1.ToString(CultureInfo.InvariantCulture)}
                WHERE Artigo = '{linha.Artigo}'";
                    BSO.DSO.ExecuteSQL(update);

                }
            }

        }
        public override void DepoisDeTransformar(ExtensibilityEventArgs e)
        {
            if (this.DocumentoCompra.Tipodoc == "VFA" || this.DocumentoCompra.Tipodoc == "VFAO")
            {
                var num = this.DocumentoCompra.Linhas.NumItens;
                for (int i = 1; i < num + 1; i++)
                {
                    var linha = this.DocumentoCompra.Linhas.GetEdita(i);
                    var query = $@"SELECT PVP1 FROM ArtigoMoeda WHERE Artigo = '{linha.Artigo}'";
                    var data = BSO.Consulta(query);
                    if (data != null && data.NumLinhas() > 0)
                    {
                        linha.CamposUtil["CDU_PVP1"].Valor = data.DaValor<string>("PVP1");
                    }
                    else
                    {
                        linha.CamposUtil["CDU_PVP1"].Valor = "0";
                    }
                }
            }
        }
    }
}
