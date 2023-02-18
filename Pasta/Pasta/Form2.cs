using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pasta
{
    public partial class Form2 : Form
    {
        public Form2(DataGridView datagrid)
        {
            InitializeComponent();
            dataGridView1.DataSource = datagrid.DataSource;
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
    }
}
