using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OphiussauPnpPortForward.Models;
using Rule = FirewallManager.Rule;

namespace OphiussauPnpPortForward
{
    public partial class FrmAddRule : Form
    {
        public FrmAddRule()
        {
            InitializeComponent();
        }

        private void btOk_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtName.Text)) throw new Exception("Invalid Name");
                if (string.IsNullOrEmpty(txtIP.Text)) throw new Exception("Invalid IP");
                if (string.IsNullOrEmpty(txtStart.Text)) throw new Exception("Invalid Start Port");
                Global.Rules.Add(new PortFowardRules()
                                 {
                                     Name      = txtName.Text,
                                     IpAdress  = txtIP.Text,
                                     StartPort = txtStart.Text.ToInt(),
                                     EndPort   = txtEnd.Text.ToInt()
                                 });
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                throw;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void txtStart_Validated(object sender, EventArgs e)
        {
            txtEnd.Text = (txtStart.Text.ToInt() + 1).ToString();
        }
    }
}
