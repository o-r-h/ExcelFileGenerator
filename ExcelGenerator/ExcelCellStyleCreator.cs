using ExcelGenerator.Classes;
using System.Drawing;

namespace ExcelGenerator
{
	public class ExcelCellStyleCreator
	{
		

		public Color BackgroundColor { get; set; } = Color.White;
		public Color TextColor { get; set; } = Color.Black;
		public int FontSize { get; set; } = 9;
		public bool Fontbold { get; set; } = false;
		public string FontName { get; set; } = "Arial";

		public ExcelCellStyle CreateStyle()
		{
			ExcelCellStyle style = new ExcelCellStyle
			{
				ExcelStyleName = "CustomStyle",
				FontSize = this.FontSize,
				FontName = this.FontName,
				FontBold = this.Fontbold,
				BackgroundColor = this.BackgroundColor,
				TextColor = this.TextColor
			};

			return style;
		}
	}
}
