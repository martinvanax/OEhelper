using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OEHelper
{
    public partial class ListEditor : Form
    {
        public ListEditor(Object o)
        {
            InitializeComponent();
            if ((o as List<EnvelopeLabel>) != null)
            {
                dataGridViewListEditor.DataSource = o;
            }
            else
            {
                MessageBox.Show("Neznamý typ objektu");
            }
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
