using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Excel {
	
	public class MyCell {

		public string Name;
		public int Row, Column;
		public string Expression { set; get; }
		public string Result { set; get; }

		public List<MyCell> DependentCells;
		public List<MyCell> CellsIDependOn;

		public MyCell(string name, int column, int row) {
			Name = name;
			Row = row;
			Column = column;
			DependentCells = new List<MyCell>();
			CellsIDependOn = new List<MyCell>();
			Expression = "";
			Result = "";
		}

		public void RefreshCell(string name, int row, int column) {
			Name = name;
			Row = row;
			Column = column;
		}
		
		public MyCell Clone() {
			return MemberwiseClone() as MyCell;
		}
	}
	
}