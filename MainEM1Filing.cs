using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EM1Filing
{
    public partial class MainEM1Filing : Form
    {
        public MainEM1Filing()
        {
            InitializeComponent();
        }

        private void EM1Filing_Load(object sender, EventArgs e)
        {

        }

        private void eM1FinalFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EM1FinalFile frm = new EM1FinalFile();
            frm.ShowDialog();
        }

        private void eM1TradingFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EM1Trading frm = new EM1Trading();
            frm.ShowDialog();
        }
    }
}
