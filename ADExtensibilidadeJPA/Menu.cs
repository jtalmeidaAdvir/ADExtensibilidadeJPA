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
        private void cb_Obras_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Obras.SelectedItem is KeyValuePair<string, string> obraSelecionada)
            {
                string codigoObraSelecionada = obraSelecionada.Key;
                string queryGetObras = $@"SELECT * FROM TDU_AD_Obras 
                                  WHERE CDU_Obra = '{codigoObraSelecionada}'";

                var DBObras = BSO.Consulta(queryGetObras);

                dataGridView1.Rows.Clear();

                if (DBObras.NumLinhas() > 0)
                {
                    DBObras.Inicio();

                    while (!DBObras.NoFim())
                    {
                        dataGridView1.Rows.Add(
                            DBObras.DaValor<string>("CDU_EntradaObra"),
                            DBObras.DaValor<string>("CDU_SaidaObra"),
                            DBObras.DaValor<string>("CDU_ContratoSubempreitada"),
                            DBObras.DaValor<int>("CDU_AutorizacaoEntrada") == 1
                        );

                        int lastRowIndex = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[lastRowIndex].DefaultCellStyle.BackColor = Color.LightYellow;

                        DBObras.Seguinte();
                    }
                }
            }
        }

        private void carregarObrasCB(Dictionary<string, string> veiculo)
        {
            var BDObras = GetObrasSumbempreiteiro(veiculo["EntidadeId"]);

            cb_Obras.Items.Clear(); // Limpa antes de adicionar novos itens

            var numLinhasObras = BDObras.NumLinhas();
            while (!BDObras.NoFim())
            {
                string codigo = BDObras.DaValor<string>("Codigo");
                string descricao = BDObras.DaValor<string>("Descricao");

                // Adiciona um item ao ComboBox
                cb_Obras.Items.Add(new KeyValuePair<string, string>(codigo, $"{codigo} - {descricao}"));

                BDObras.Seguinte();
            }

            cb_Obras.DisplayMember = "Value"; // O que será exibido
            cb_Obras.ValueMember = "Key"; // O valor interno
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




        private void BT_Salvar_Click_Click(object sender, EventArgs e)
        {
            try
            {
                // Atualiza a tabela Geral_Entidade
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
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar os dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Verifica se há linhas no DataGridView
                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Não há dados para salvar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Verifica se uma obra foi selecionada no ComboBox
                if (cb_Obras.SelectedItem == null)
                {
                    MessageBox.Show("Selecione uma obra antes de salvar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Obtém o código da obra selecionada
                string codigoObraSelecionada = ((KeyValuePair<string, string>)cb_Obras.SelectedItem).Key;

                // Percorre cada linha do DataGridView
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    // Garante que a linha não esteja vazia
                    if (row.Cells[0].Value != null)
                    {
                        string entradaObra = row.Cells[0].Value.ToString();
                        string saidaObra = row.Cells[1].Value.ToString();
                        string contratoSubempreitada = row.Cells[2].Value.ToString();
                        bool autorizacaoEntrada = Convert.ToBoolean(row.Cells[3].Value);
                        Guid id = Guid.NewGuid();

                        // Monta a query de inserção
                        string queryUpsert = $@"
    IF EXISTS (
        SELECT 1 FROM TDU_AD_Obras 
        WHERE CDU_Obra = '{codigoObraSelecionada}' 
        AND CDU_EntradaObra = '{entradaObra}'
        AND CDU_SaidaObra = '{saidaObra}'
        AND CDU_ContratoSubempreitada = '{contratoSubempreitada}'
    )
    BEGIN
        UPDATE TDU_AD_Obras 
        SET CDU_AutorizacaoEntrada = {(autorizacaoEntrada ? 1 : 0)}
        WHERE CDU_Obra = '{codigoObraSelecionada}' 
        AND CDU_EntradaObra = '{entradaObra}'
        AND CDU_SaidaObra = '{saidaObra}'
        AND CDU_ContratoSubempreitada = '{contratoSubempreitada}';
    END
    ELSE
    BEGIN
        INSERT INTO TDU_AD_Obras 
        (CDU_Codigo, CDU_Obra, CDU_EntradaObra, CDU_SaidaObra, CDU_ContratoSubempreitada, CDU_AutorizacaoEntrada) 
        VALUES 
        ('{id}', '{codigoObraSelecionada}', '{entradaObra}', '{saidaObra}', '{contratoSubempreitada}', {(autorizacaoEntrada ? 1 : 0)});
    END
";

                        BSO.DSO.ExecuteSQL(queryUpsert);


                    }
                }

                MessageBox.Show("Registros adicionados com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar os dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



    }
}
