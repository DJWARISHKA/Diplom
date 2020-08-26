using System;
using System.Windows.Forms;

namespace Diplom
{
    public partial class Profile : Form
    {
        public Profile()
        {
            InitializeComponent();
        }
        public int row;
        private void button1_Click(object sender, EventArgs e)
        {
            row = gridView1.GetSelectedRows()[0];
        }

        private void Profile_Load(object sender, EventArgs e)
        {
            excelDataSource1.Fill();
        }
    }
}
