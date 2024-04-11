using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Open.Nat;
using OphiussauPnpPortForward.Models;
using OphiussauPnpPortForward.Properties;

namespace OphiussauPnpPortForward
{
    public partial class MainForm : Form
    {
        private BindingList<PortFowardRules> listBinding;
        private NatDiscoverer discoverer;
        public MainForm()
        {
            InitializeComponent();
        }

        private void addNewRuleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            (new FrmAddRule()).ShowDialog();
            LoadGrid();
        }

        private void saveRulesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.IO.File.WriteAllText("rules.json", JsonConvert.SerializeObject(Global.Rules, Formatting.Indented));
        }

        private async void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                LoadGrid();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void LoadGrid()
        {
            listBinding = new BindingList<PortFowardRules>(Global.Rules);

            dataGridView1.Columns.Clear();
            dataGridView1.DataSource = listBinding;
            foreach (DataGridViewTextBoxColumn col in dataGridView1.Columns)
            {
                col.ReadOnly = false;
            }

            dataGridView1.Columns.Add(new DataGridViewImageColumn() { Name = "clRefresh", HeaderText = "Refresh Status", Image = Properties.Resources.Refresh });
            dataGridView1.Columns.Add(new DataGridViewImageColumn() { Name = "clDel", HeaderText = "Delete Rule", Image = Properties.Resources.Delete });
            dataGridView1.Columns.Add(new DataGridViewImageColumn() { Name = "clRef1", HeaderText = "Refresh Router", Image = Properties.Resources.Refresh });
            dataGridView1.Columns.Add(new DataGridViewImageColumn() { Name = "clRef2", HeaderText = "Refresh Firewall", Image = Properties.Resources.Refresh });
        }

        private async void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "clDel")
            {
                int index = e.RowIndex;
                var obj = (PortFowardRules)listBinding[index];

                if(MessageBox.Show("Do you want delete this rule?","DELETE RULE!!!",MessageBoxButtons.YesNo)==DialogResult.Cancel) return;
                try
                {
                    Global.RemoveFirewallRules(obj);
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Error deleting Firewall rule. Please delete manually!!!");
                }
                try
                {
                    Global.RemoveRouterPort(obj);
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Error deleting Router rule. Please delete manually!!!");
                } 
                Global.Rules.RemoveAt(index);
                LoadGrid();
            }
            if (dataGridView1.Columns[e.ColumnIndex].Name == "clRef1")
            {
                int index = e.RowIndex;
                var obj = (PortFowardRules)listBinding[index];
                Global.OpenRouterPort(obj);
                RefreshLine(index);
            }
            if (dataGridView1.Columns[e.ColumnIndex].Name == "clRef2")
            {
                int index = e.RowIndex;
                var obj = (PortFowardRules)listBinding[index];
                Global.OpenFireWallPort(obj);
                RefreshLine(index);
            }

            if (dataGridView1.Columns[e.ColumnIndex].Name == "clRefresh")
            {
                int index = e.RowIndex;
                RefreshLine(index);

            }
        }

        private async void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
        }

        private async void RefreshLine(int index)
        {
            try
            {
                var obj = (PortFowardRules)listBinding[index];

                Color color = Color.PaleGreen;

                string msg = "";

                for (int port = obj.StartPort; port <= obj.EndPort; port++)
                {

                    bool firewallServerPort = Global.IsPortOpen(obj.Name, port);


                    if (!firewallServerPort)
                    {
                        color = Color.Orange;
                        msg += $"Missing firewall rules for port {port}";
                    }


                    foreach (DataGridViewCell cell in dataGridView1.Rows[index].Cells)
                        cell.Style.BackColor = color;
                }
                obj.Message = msg;
                foreach (DataGridViewCell cell in dataGridView1.Rows[index].Cells)
                    cell.Style.BackColor = color;

                //TODO:sometimes this give an error about canceled task in searcher - Search
                discoverer = new NatDiscoverer();
                var device = await discoverer.DiscoverDeviceAsyncWithoutToken();

                for (int port = obj.StartPort; port <= obj.EndPort; port++)
                {

                    Mapping serverMapping = await device.GetSpecificMappingAsync(Protocol.TcpUpd, port);
                    bool firewallServerPort = Global.IsPortOpen(obj.Name, port);


                    if (serverMapping == null)
                    {
                        color = Color.Orange;
                        msg += "Missing router rules";
                    }



                    foreach (DataGridViewCell cell in dataGridView1.Rows[index].Cells)
                        cell.Style.BackColor = color;
                }

                obj.Message = msg;

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
                foreach (DataGridViewCell cell in dataGridView1.Rows[index].Cells)
                    cell.Style.BackColor = Color.Red;
            }
        }
    }
}
