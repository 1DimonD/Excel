using System;
using System.IO;
using System.Windows.Forms;
using Excel;
using Xunit;

namespace TestExcel {
	
	public class Tests {
		
		[Fact]
		public void TestColsRows() {
			Table table = new Table(10, 10);
			Assert.True(table.DifferenceBetweenRowsAndColumns() == 0);
			
			table.AddColumns(4);
			Assert.True(table.DifferenceBetweenRowsAndColumns() == -4);

			table.AddRows(4);
			Assert.True(table.DifferenceBetweenRowsAndColumns() == 0);
			
			table.Dispose();
		}

		[Theory]
		[InlineData("2 ^ 0", "1")]
		[InlineData("4 ^ (1/2)", "2")]
		[InlineData("inc -1", "0")]
		[InlineData("1 < 2", "1")]
		[InlineData("2 <> 2", "0")]
		public void TestEval(string expr, string result) {
			Interpreter inter = new Interpreter(@"Interpreter\Interpreter_run.exe");

			string res = inter.Evaluate("A1", expr);
			Assert.Equal(result, res);
			
			inter.Dispose();
		}

		[Fact]
		public void TestXML() {
			string xml = File.ReadAllText(@"..\..\..\..\USERS_FILES\Fake_Excel.fex");
			Table table = new Table(xml);

			Assert.Equal(xml, table.GetXML());

			table.Dispose();
		}
	}
}