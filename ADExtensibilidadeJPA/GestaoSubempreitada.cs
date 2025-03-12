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
    }
}
