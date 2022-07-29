using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Excel {
	public partial class Table {
		
		public void AddColumns(int count) {
			var cols = _content.Columns;
			for (int i = 0; i < count; i++) {
				string index = GenerateColumnIndex((uint)cols.Count);
				for (int j = 0; j < _content.Rows.Count; j++) {
					string name = GetName(cols.Count, j);
					Cells.Add(name, new MyCell(name, cols.Count, j));
				}
				cols.AddRange(CreateColumn(index));
			}
		}

		private DataGridViewTextBoxColumn CreateColumn(string index) {
			var col = new DataGridViewTextBoxColumn();
			col.Name = col.HeaderText = index;
			col.SortMode = DataGridViewColumnSortMode.NotSortable;
			col.HeaderCell.ContextMenuStrip = CreateContextMenu(_content.Columns.Count.ToString(), "column");
			return col;
		}
		
		private void DeleteColumn(int num) {
			try {
				for (int i = 0; i < _content.Rows.Count; i++) {
					string name = GetName(_content.Rows[i].Cells[num]);
					if(Cells[name].DependentCells.Count > 0) throw new Exception();
				}
				
				MoveTableLeft(Convert.ToInt32(num));
				
				for (int i = 0; i < _content.Rows.Count; i++) {
					string name = GetName(_content.Rows[i].Cells[_content.Columns.Count - 1]);
					Cells.Remove(name);
				}
				_content.Columns.RemoveAt(_content.Columns.Count - 1);
			}
			catch (Exception ex) {
				MessageBox.Show("Some cells are on depend");
			}
		}
		
		public void AddRows(int count) {
			for (int i = 0; i < count; i++)
			{
				int index = _content.RowCount++;
				
				for (int j = 0; j < _content.Columns.Count; j++) {
					string name = GetName(j, index);
					Cells.Add(name, new MyCell(name, j, index));
				}
				
				_content.Rows[index].HeaderCell.Value = (index + 1).ToString();
				_content.Rows[index].HeaderCell.ContextMenuStrip = CreateContextMenu(index.ToString(), "row");
			}
		}
		
		private void DeleteRow(int num) {
			try {
				foreach (DataGridViewCell cell in _content.Rows[num].Cells) {
					if (Cells[GetName(cell)].DependentCells.Count > 0) throw new Exception();
				}
				
				MoveTableUp(num);
				
				foreach (DataGridViewCell cell in _content.Rows[_content.Rows.Count - 1].Cells) {
					Cells.Remove(GetName(cell));
				}
				_content.Rows.RemoveAt(_content.Rows.Count - 1);
			}
			catch (Exception ex) {
				MessageBox.Show("Some cells are on depend");
			}
		}
		
		public int DifferenceBetweenRowsAndColumns() => _content.RowCount - _content.Columns.Count;

		private void SwitchCells(int col, int row, MyCell toSwitch) {
			string name = GetName(col, row);
			Cells[name] = toSwitch.Clone();
			Cells[name].RefreshCell(name, row, col);

			foreach (MyCell cell in Cells[name].CellsIDependOn) {
				cell.DependentCells.Remove(toSwitch);
				cell.DependentCells.Add(Cells[name]);
			}

			var toDelete = new List<MyCell>();
			
			foreach (MyCell cell in Cells[name].DependentCells) {
				cell.CellsIDependOn.Remove(toSwitch);
				if (cell.Column == 0 || cell.Row == 0) {
					toDelete.Add(cell);
					continue;
				}

				string oldName = toSwitch.Name;
				int startIndex = 0;
				string newExpression = "";
				foreach (var indexOfCoincidence in FindIndexesOfCoincidences(cell.Expression, "${")) {
					int length = cell.Expression.Substring(indexOfCoincidence).IndexOf("}") - 2;
					string varName = cell.Expression.Substring(indexOfCoincidence + 2, length);
					
					if (varName == oldName) {
						newExpression += cell.Expression.Substring(startIndex, indexOfCoincidence + 2 - startIndex) + name;
						startIndex = indexOfCoincidence + 2 + length;
					}
				}

				if (startIndex != 0) {
					newExpression += cell.Expression.Substring(startIndex);
					SetCellsExpression(cell, newExpression);
					RecurentEvaluation(cell);
				}
			}

			foreach (var cell in toDelete)
				Cells[name].DependentCells.Remove(cell);
		}
		
		private void MoveTableUp(int startIndex) {
			for (int i = 0; i < _content.Columns.Count; i++) {
				for (int j = startIndex + 1; j < _content.Rows.Count; j++) {
					_content.Rows[j - 1].Cells[i].Value = _content.Rows[j].Cells[i].Value;
					SwitchCells(i, j - 1, Cells[GetName(i, j)] );
				}
			}
		}

		private void MoveTableLeft(int startIndex) {
			for (int i = 0; i < _content.Rows.Count; i++) {
				for (int j = startIndex + 1; j < _content.Columns.Count; j++) {
					_content.Rows[i].Cells[j - 1].Value = _content.Rows[i].Cells[j].Value;
					SwitchCells(j - 1, i, Cells[GetName(j, i)] );
				}
			}
		}
        
		private void MoveTableDown(int startIndex) {
			for (int i = 0; i < _content.Columns.Count; i++) {
				for (int j = _content.Rows.Count - 1; j > startIndex; j--) {
					if (j == startIndex + 1) {
						_content.Rows[j].Cells[i].Value = "";
						SwitchCells(i, j, new MyCell(GetName(i, j), i, j) );
					} else {
						_content.Rows[j].Cells[i].Value = _content.Rows[j - 1].Cells[i].Value;
						SwitchCells(i, j, Cells[GetName(i, j - 1)] );
					}
				}
			}
		}

		private void MoveTableRight(int startIndex) {
			for (int i = 0; i < _content.Rows.Count; i++) {
				for (int j = _content.Columns.Count - 1; j > startIndex; j--) {
					if (j == startIndex + 1) {
						_content.Rows[i].Cells[j].Value = "";
						SwitchCells(j, i, new MyCell(GetName(j, i), j, i) );
					} else {
						_content.Rows[i].Cells[j].Value = _content.Rows[i].Cells[j - 1].Value;
						SwitchCells(j, i, Cells[GetName(j - 1, i)] );
					}
				}
			}
		}
		
	}

}