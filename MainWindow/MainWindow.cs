using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class MainWindow : Form
    {
        private Table table;
        
        public MainWindow()
        {
            InitializeComponent();

            table = new Table(26, 10);
            table.BindTo(textBox1);
            table.DisplayAt(Controls);
        }

        private void MainWindow_FormClosed(object sender, FormClosedEventArgs e) {
            table?.Dispose();
            Dispose();
        }
        
    }
}