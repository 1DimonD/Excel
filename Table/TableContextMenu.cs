using System;
using System.Windows.Forms;

namespace Excel
{
    public partial class Table {
    
        private ContextMenuStrip CreateContextMenu(string name, string type)
        {
            var cms = new ContextMenuStrip();
            cms.Name = name;

            var addMenuItem = new ToolStripMenuItem("Add " + type);
            var deleteMenuItem = new ToolStripMenuItem("Delete " + type);

            cms.Items.AddRange(new ToolStripItem[] {addMenuItem, deleteMenuItem});

            if (type == "row") {
                addMenuItem.Click += addRowToolStripMenuItem_Click;
                deleteMenuItem.Click += deleteRowToolStripMenuItem_Click;
            } else if(type == "column") {
                addMenuItem.Click += addColumnToolStripMenuItem_Click;
                deleteMenuItem.Click += deleteColumnToolStripMenuItem_Click;
            }

            return cms;
        }
         
        private void addRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var tsmi = (ToolStripMenuItem) sender;
            var cms = (ContextMenuStrip) tsmi.GetCurrentParent();

            AddRows(1);
            MoveTableDown(Convert.ToInt32(cms.Name));
            EvaluateAll();
        }
        
        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var tsmi = (ToolStripMenuItem) sender;
            var cms = (ContextMenuStrip) tsmi.GetCurrentParent();
            
            DeleteRow(Convert.ToInt32(cms.Name));
            EvaluateAll();
        }
        
        private void addColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var tsmi = (ToolStripMenuItem) sender;
            var cms = (ContextMenuStrip) tsmi.GetCurrentParent();

            AddColumns(1);
            MoveTableRight(Convert.ToInt32(cms.Name));
            EvaluateAll();
        }

        private void deleteColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var tsmi = (ToolStripMenuItem) sender;
            var cms = (ContextMenuStrip) tsmi.GetCurrentParent();

            DeleteColumn(Convert.ToInt32(cms.Name));
            EvaluateAll();
        }
        
    }
}