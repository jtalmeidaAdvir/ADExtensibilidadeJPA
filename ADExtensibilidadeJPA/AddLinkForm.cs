using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADExtensibilidadeJPA
{
    public partial class AddLinkForm : Form
    {
        public string Link { get; private set; }
        public AddLinkForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string inputLink = txtLink.Text.Trim();

            if (IsValidUrl(inputLink))
            {
                Link = inputLink;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("O link inserido não é válido. Por favor, insira um URL válido.",
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private bool IsValidUrl(string url)
        {
            return Uri.TryCreate(url, UriKind.Absolute, out Uri uriResult) &&
                   (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
        }

    }
}
