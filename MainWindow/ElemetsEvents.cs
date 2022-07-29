using System;
using System.IO;
using System.Windows.Forms;

namespace Excel
{
    
    public partial class MainWindow
    {
        
        private void button1_Click(object sender, EventArgs e) {
            table.EvaluateForCurrentCell();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            table.AddColumns(1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            table.AddRows(1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int dif = table.DifferenceBetweenRowsAndColumns();
            if (dif > 0) table.AddColumns(dif);
            else table.AddRows(Math.Abs(dif));
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e) {
            table.RemoveAt(Controls);
            table.Dispose();
            table = new Table(26, 10);
            table.BindTo(textBox1);
            table.DisplayAt(Controls);
        }
        
        private void saveToolStripMenuItem_Click(object sender, EventArgs e) {
            saveFileDialog1.Filter = "FakeExcel docs (*.fex)|*.fex|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel) return;

            string fileName = saveFileDialog1.FileName;
            File.WriteAllText(fileName, table.GetXML());
        }
        
        private void aboutProgramToolStripMenuItem_Click(object sender, EventArgs e) {
            string info = File.ReadAllText(@"D:\Univ_C#\Work\Excel\MainWindow\About.txt");
            MessageBox.Show(info);
        }
        
        private void openToolStripMenuItem_Click(object sender, EventArgs e) {
            openFileDialog1.Filter = "FakeExcel docs (*.fex)|*.fex|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel) return;

            try {
                string fileName = openFileDialog1.FileName;
                string XMLtable = File.ReadAllText(fileName);
                table.RemoveAt(Controls);
                table.Dispose();
                table = new Table(XMLtable);
                table.BindTo(textBox1);
                table.DisplayAt(Controls);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
    }
}