using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Excel {
    public partial class Table : IDisposable { 

        private DataGridView _content;

        private TextBox _displayForCellContent;

        private Interpreter _interpreter;

        private Dictionary<string, MyCell> Cells;

        public Table(int width, int height) {
            InitializeTable(width, height);
        }

        private void InitializeTable(int width, int height) {
            _content = new DataGridView();
            ((System.ComponentModel.ISupportInitialize) _content).BeginInit();
            _content.AllowUserToAddRows = false;
            _content.AllowUserToDeleteRows = false;
            _content.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            _content.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _content.GridColor = System.Drawing.SystemColors.ActiveBorder;
            _content.Location = new System.Drawing.Point(13, 77);
            _content.Name = "Content";
            _content.RowHeadersWidth = 70;
            _content.RowTemplate.Height = 24;
            _content.Size = new System.Drawing.Size(729, 522);
            _content.TabIndex = 3;
            
            _content.CellEnter += Content_CellEnter;
            _content.EditingControlShowing += Content_EditingControlShowing;
            _content.CellLeave += Content_CellLeave;

            ((System.ComponentModel.ISupportInitialize) (_content)).EndInit();
            
            Cells = new Dictionary<string, MyCell>();
            
            AddColumns(width);
            AddRows(height);

            for (int i = 0; i < height; i++) {
                for (int j = 0; j < width; j++) {
                    string name = GetName(j, i);
                    if(!Cells.ContainsKey(name)) Cells.Add(name, new MyCell(name, j, i));
                }
            }

            _displayForCellContent = new TextBox();
            _interpreter = new Interpreter(@"Interpreter\Interpreter_run.exe");
        }

        public string GetName(DataGridViewCell cell) {
            return cell.OwningColumn.Name + (cell.RowIndex + 1);
        }

        public string GetName(int col, int row) {
            return GenerateColumnIndex((uint) col) + (row + 1);
        }

        public static string GenerateColumnIndex(uint num) {
            if (num / 26 == 0) return "" + (char) (65 + num % 26);
            return GenerateColumnIndex(num / 26 - 1) + (char) (65 + num % 26);
        }

        public void DisplayAt(Control.ControlCollection controls) {
            controls.Add(_content);
        }
        
        public void RemoveAt(Control.ControlCollection controls) {
            controls.Remove(_content);
        }

        public void BindTo(TextBox textBox1) {
            _displayForCellContent = textBox1;

            textBox1.KeyUp += delegate(object sender, KeyEventArgs args) {
                if (_content.CurrentCell != null) {
                    if (args.KeyCode == Keys.Enter) {
                        _content.CurrentCell = null;
                        EvaluateForCurrentCell();
                        textBox1.Text = "";
                    } else {
                        _content.CurrentCell.Value = textBox1.Text;
                        textBox1.Focus();
                        textBox1.SelectionStart = textBox1.Text.Length;
                    }
                }
            };

            textBox1.Leave += delegate(object sender, EventArgs e) { EvaluateForCurrentCell(); };
        }

        private void Content_CellLeave(object sender, DataGridViewCellEventArgs e) {
            if(!_displayForCellContent.Focused) EvaluateForCurrentCell();
        }

        private void Content_CellEnter(object sender, DataGridViewCellEventArgs e) {
            _content.CurrentCell.Value = _displayForCellContent.Text = 
                Cells[GetName(_content.CurrentCell)].Expression;
        }
        
        private void Content_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl tb)
            {
                tb.KeyUp -= Content_KeyUp;
                tb.KeyUp += Content_KeyUp;
            }
        }

        private void Content_KeyUp(object sender, KeyEventArgs e) {
            _displayForCellContent.Text = _content.CurrentCell.GetEditedFormattedValue(
                _content.CurrentCell.RowIndex, DataGridViewDataErrorContexts.Display).ToString();
        }

        public void Dispose() {
            _content?.Dispose();
            _interpreter?.Dispose();
        }

    }
}