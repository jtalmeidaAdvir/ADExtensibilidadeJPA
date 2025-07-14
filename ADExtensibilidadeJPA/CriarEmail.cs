using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class CriarEmail : Form
    {
        private ErpBS100.ErpBS BSO;
        private string IdSelecionado;

        public string Email { get; set; }
        public CriarEmail(ErpBS100.ErpBS bSO, string idSelecionado)
        {
            InitializeComponent();
            this.BSO = bSO;
            this.IdSelecionado = idSelecionado;
        }

        private void bt_Gravar_Click(object sender, System.EventArgs e)
        {
            string email = txt_email.Text;

            if (IsValidEmail(email))
            {
                Email = email;

                var upddate = $@"UPDATE Geral_Entidade_Contactos
                            SET Email = '{Email}'
                            WHERE EntidadeID = CAST('{IdSelecionado}' AS uniqueidentifier);";
                BSO.DSO.ExecuteSQL(upddate);









                MessageBox.Show("E-mail válido! Dados gravados com sucesso.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;  // Indica que o e-mail foi inserido corretamente
                this.Close();
            }
            else
            {
                MessageBox.Show("E-mail inválido! Por favor, insira um e-mail válido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool IsValidEmail(string email)
        {
            string pattern = @"^[^@\s]+@[^@\s]+\.[^@\s]+$"; 
            return Regex.IsMatch(email, pattern);
        }
    }
}
