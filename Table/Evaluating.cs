using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Excel {
	public partial class Table {

		private void EvaluateAll() {
			foreach (var cell in Cells) {
				RecurentEvaluation(cell.Value);
			}

			_content.RefreshEdit();
		}
		
		public void EvaluateForCurrentCell() {
			if (_content.CurrentCell == null) return;

			string name = GetName(_content.CurrentCell);
			try {
				SetCellsExpression(Cells[name], _displayForCellContent.Text);
				RecurentEvaluation(Cells[name]);
				_content.RefreshEdit();
			}
			catch(Exception ex) {
				MessageBox.Show("Program has some troubles. Restart it and don`t do such actions anymore.");
			}
		}

		private void SetCellsExpression(MyCell cell, string value) {
			MyCell tmp = cell.Clone();
			try {
				RefreshDependencies(cell, value);

				cell.Expression = value;
			}
			catch (Exception e) {
				cell = tmp;
				foreach (var cellIDependOn in cell.CellsIDependOn) {
					if (!cellIDependOn.DependentCells.Contains(cell)) cellIDependOn.DependentCells.Add(cell);
				}
				
				if(e.GetType().ToString() == "System.Collections.Generic.KeyNotFoundException") 
					MessageBox.Show("Cell name is incorrect");
				else MessageBox.Show(e.Message);
			}
		}
		
		private void RefreshDependencies(MyCell cell, string value) {
			foreach (var cellDependOn in cell.CellsIDependOn) {
				cellDependOn.DependentCells.Remove(cell);
			}
			cell.CellsIDependOn.Clear();

			foreach (var indexOfCoincidence in FindIndexesOfCoincidences(value, "${")) {
				int length = value.Substring(indexOfCoincidence).IndexOf("}") - 2;
				if (length < 0) throw new Exception("Can`t define a Cell");

				string name = value.Substring(indexOfCoincidence + 2, length);

				if (CheckDepend(Cells[name], cell)) throw new Exception("Infinity referring");
                
				if (!cell.CellsIDependOn.Contains(Cells[name])) cell.CellsIDependOn.Add(Cells[name]);
			}
            
			foreach (var cellIDependOn in cell.CellsIDependOn) {
				if (!cellIDependOn.DependentCells.Contains(cell)) cellIDependOn.DependentCells.Add(cell);
			}
		}
		
		private IEnumerable<int> FindIndexesOfCoincidences(string line, string part) {
			int index = 0, shiftedIndex;
			while(
				Convert.ToBoolean( ( shiftedIndex = line.Substring(index).IndexOf(part) ) + 1)
			) {
				yield return (index += shiftedIndex);
				index += part.Length;
			}
		}
        
		private bool CheckDepend(MyCell cell, MyCell cellForCheck) {
			if (cell == cellForCheck) return true;
			if (cell.CellsIDependOn.Contains(cellForCheck)) return true;
            
			foreach (var cell1 in cell.CellsIDependOn) {
				if (CheckDepend(cell1, cellForCheck)) return true;
			}
            
			return false;
		}
		
		private void RecurentEvaluation(MyCell cell) {
			foreach (var cellIDepend in cell.CellsIDependOn) {
				if (cellIDepend.Result == "" && cellIDepend.Expression == "") throw new Exception("Using an empty cell");
				if(cellIDepend.Result == "") RecurentEvaluation(cellIDepend);
			}

			if (cell.Expression != "") {
				cell.Result = _interpreter.Evaluate(cell.Name, cell.Expression);
				_content.Rows[cell.Row].Cells[cell.Column].Value = cell.Result;
			}
			
			foreach (var dependentCell in cell.DependentCells) {
				RecurentEvaluation(dependentCell);
			}
		}
		
	}

}