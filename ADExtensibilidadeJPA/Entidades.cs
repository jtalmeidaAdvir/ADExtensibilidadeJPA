using Microsoft.Office.Interop.Outlook;
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
    public partial class Entidades : CustomForm
    {
        private List<string> listaItens = new List<string>();

        public Entidades()
        {
            InitializeComponent();
            this.Load += new EventHandler(Entidades_Load); // Associa o evento Load ao método Entidades_Load
        }

        private void Entidades_Load(object sender, EventArgs e)
        {
            GetItemToCBBox(); // Chama a função ao carregar o formulário
            ConfigurarComboBox();
        }

        private void GetItemToCBBox()
        {
            var query = "SELECT Nome,* FROM Geral_Entidade WHERE Tipo = '7' OR  Tipo = '3'";
            var lista = BSO.Consulta(query);

            if (lista != null && lista.NumLinhas() > 0)
            {
                lista.Inicio();

                for (int i = 0; i < lista.NumLinhas(); i++)
                {
                    var item = lista.DaValor<string>("Nome");
                    if (!string.IsNullOrEmpty(item))
                    {
                        listaItens.Add(item); 
                        comboBox1.Items.Add(item);
                    }
                    lista.Seguinte();
                }
            }
        }

        private void ConfigurarComboBox()
        {
            comboBox1.DropDownStyle = ComboBoxStyle.DropDown; // Permite digitar
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend; // Sugestões automáticas
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems; // Fonte de sugestões será os itens da lista
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedItem = comboBox1.SelectedItem?.ToString(); // Pega o item selecionado

            if (!string.IsNullOrEmpty(selectedItem))
            {
                var query = $@"SELECT CDU_TrataSGS FROM Geral_Entidade WHERE Nome = '{selectedItem}'";
                var trataSGS = BSO.Consulta(query);

                checkBox1.Checked = trataSGS.DaValor<bool>("CDU_TrataSGS");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null || string.IsNullOrEmpty(comboBox1.SelectedItem.ToString()))
            {
                // Se não houver nenhum item selecionado, mostra a mensagem de alerta
                MessageBox.Show("Por favor, selecione um item no ComboBox.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                // Caso contrário, executa a atualização
                var update = $@"UPDATE Geral_Entidade
                    SET CDU_TrataSGS = '{checkBox1.Checked}'
                    WHERE Nome = '{comboBox1.SelectedItem?.ToString()}'";
                BSO.DSO.ExecuteSQL(update);
                MessageBox.Show("A atualização foi realizada com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
        }
    }

}
