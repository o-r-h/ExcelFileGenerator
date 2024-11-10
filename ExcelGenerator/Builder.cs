using ExcelGenerator.Classes;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace ExcelGenerator
{
   
    public class BuilderExcel<T> 
    {   
        public string FileName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
		public ExcelCellStyle TitleCellStyles {  get; set; }
		public ExcelCellStyle TableHeaderCellStyles { get; set; }
		public List<T> listOfRecords { get; set; }
		public string TableHeaderNames { get; set; }= string.Empty;
		public string Title { get; set; } = string.Empty;
        public int ColStart { get; set; } = 1;
        public int RowStart { get; set; } = 1;
		public string MessageError {  get; set; } =string.Empty;
		public int ColumnWidth { get; set; } = 25;

		private bool IsValidToBuild(){
			if( string.IsNullOrEmpty(this.FileName))
			{
				this.MessageError = "Error, fileName is mandatory";
				return false;
			}
			if (string.IsNullOrEmpty(this.SheetName))
			{
				this.MessageError = "Error, SheetName is mandatory";
				return false;
			}
			if (this.TitleCellStyles == null)
			{
				this.MessageError = "Error, TitleCellStyles must have value";
				return false;
			}

			if (this.TableHeaderCellStyles == null)
			{
				this.MessageError = "Error, TableHeaderCellStyles must have value";
				return false;
			}

			return true;

		}

		public ExcelPackage Builder(int initialCellRowPos, int initialColRowPos)
		{
			if (!this.IsValidToBuild()) {
				throw new Exception($"Exception,{ this.MessageError} ");
			}
			ExcelPackage expack = new ExcelPackage();
			ExcelFile xls = new ExcelFile();
			ExcelWorkSheet sheet = new ExcelWorkSheet();
			List<ExcelCellStyle> excelCellStyles = new List<ExcelCellStyle>();

			sheet.Cells = new List<Cell>();
			xls.ExcelWorkSheets = new List<ExcelWorkSheet>();
			xls.ExcelFileName = this.FileName;

			sheet.Name = this.SheetName;
			excelCellStyles.Add(TitleCellStyles);
			excelCellStyles.Add(TableHeaderCellStyles);

			xls.ExcelWorkSheets.Add(sheet);
			xls.ExcelWorkSheets[0].ExcelCellStyles = excelCellStyles;

			int ipos = initialCellRowPos;
			int jpos = initialColRowPos;
			int colNbr = ExcelHelper.GetQuantityFieldInClass<T>();
			List<Cell> lista = new List<Cell>();
			List<Cell> cellList = new List<Cell>();
			string[] listHeaderNames = null;

			//if title is not null;
			if (!string.IsNullOrEmpty(this.Title))
			{
				Cell cell = new Cell { ColPos = 1, RowPos = 1, Value = this.Title };
				cell.Style = xls.ExcelWorkSheets[0].ExcelCellStyles[0];
				cellList.Add(cell);
				sheet.Cells.Add(cell);
				this.RowStart = ipos;
				ipos++;

			}
			//if TableHeaderNames is not null;
			if (!string.IsNullOrEmpty(this.TableHeaderNames))
			{
				 listHeaderNames = TableHeaderNames.Split(',');
				int hx = 1;
				foreach (string s in listHeaderNames) {

					Cell cell = new Cell { ColPos = hx, RowPos = ipos, Value = s };
					cell.Style = xls.ExcelWorkSheets[0].ExcelCellStyles[1];
					cellList.Add(cell);
					sheet.Cells.Add(cell);
				
					hx++;
				}
				ipos++;

			}

			colNbr = ExcelHelper.GetQuantityFieldInClass<T>();
			jpos = initialColRowPos;
			foreach (T item in listOfRecords)
			{

				foreach (PropertyInfo property in item.GetType().GetProperties())
				{
					Cell c = new Cell();
					c.Value = property.GetValue(item).ToString();
					c.ColPos = jpos;
					c.RowPos = ipos;
					c.Type = property.GetType().ToString();
					jpos++;
					cellList.Add(c);
				}
				jpos = initialColRowPos;
				ipos++;
			}


			foreach (var item in cellList)
			{
				sheet.Cells.Add(item);
			}


			foreach (var item in xls.ExcelWorkSheets)
			{
				expack.Workbook.Worksheets.Add(item.Name);
			}

			int x = 1;
			foreach (var item in xls.ExcelWorkSheets)
			{
				foreach (var subitem in item.Cells)
				{
					expack.Workbook.Worksheets[x].Cells[subitem.RowPos, subitem.ColPos].Value = subitem.Value;
				}
				x++;
			}


			// setup values and styles for table content
			int sheetIndex = 0;
			foreach (var worksheet in xls.ExcelWorkSheets)
			{
				var excelWorksheet = expack.Workbook.Worksheets[sheetIndex + 1];

				foreach (var subitem in worksheet.Cells)
				{
					var cellToStyle = excelWorksheet.Cells[subitem.RowPos, subitem.ColPos];
					cellToStyle.Value = subitem.Value;

					// Asignar estilo desde tu lista de estilos personalizada
					if (subitem.Style != null)
					{
						ApplyCustomStyle(cellToStyle, subitem.Style);
					}
				}
				sheetIndex++;
			}

			//Setup column's witdh
            for (int i = 1; i <= listHeaderNames.Length; i++)
            {
				expack.Workbook.Worksheets[1].Column(i).Width = this.ColumnWidth;
			}
            return expack;

		}

		
		private void ApplyCustomStyle(ExcelRange cell, ExcelCellStyle customStyle)
		{
			try
			{
				if (customStyle != null)
				{
					cell.Style.Font.Bold = customStyle.FontBold;
					cell.Style.Font.Size = customStyle.FontSize;
					cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
					cell.Style.Fill.BackgroundColor.SetColor(customStyle.BackgroundColor);
					cell.Style.Font.Color.SetColor(customStyle.TextColor);
				}

			}
			catch (Exception ex)
			{
				throw ex;
			}

		}

	}


}
