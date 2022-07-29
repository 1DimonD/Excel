using System;
using System.Xml;


namespace Excel {
	public partial class Table {

		public Table(string XMLRepresenting) {
			XmlDocument xDoc = new XmlDocument();
			xDoc.LoadXml(XMLRepresenting);
			
			XmlNode currentNode = xDoc.DocumentElement;
			currentNode = currentNode?.FirstChild;
			if(currentNode?.Name != "width") throw new Exception("Wrong file coding");
			int width = Convert.ToInt32(currentNode.InnerText);
			
			currentNode = currentNode.NextSibling;
			if(currentNode?.Name != "height") throw new Exception("Wrong file coding");
			int height = Convert.ToInt32(currentNode?.InnerText);
				
			InitializeTable(width, height);

			currentNode = currentNode?.NextSibling;
			if(currentNode?.Name != "cells") throw new Exception("Wrong file coding");
			currentNode = currentNode.FirstChild;
			
			 do {
				string name = currentNode.Name;
				SetCellsExpression(Cells[name], currentNode.InnerText);
			 } while ((currentNode = currentNode.NextSibling) != null);
			
			EvaluateAll();
		}

		public string GetXML() {
			XmlDocument xDoc = new XmlDocument();

			XmlElement table = xDoc.CreateElement("table");
			xDoc.AppendChild(table);

			XmlElement width = xDoc.CreateElement("width");
			width.InnerText = _content.ColumnCount.ToString();
			table.AppendChild(width);

			XmlElement height = xDoc.CreateElement("height");
			height.InnerText = _content.RowCount.ToString();
			table.AppendChild(height);

			XmlElement cells = xDoc.CreateElement("cells");
			foreach (var originCell in Cells) {
				XmlElement cell = xDoc.CreateElement(originCell.Key);
				cell.InnerText = originCell.Value.Expression;
				cells.AppendChild(cell);
			}
			table.AppendChild(cells);

			return xDoc.InnerXml;
		} 
		
	}
}