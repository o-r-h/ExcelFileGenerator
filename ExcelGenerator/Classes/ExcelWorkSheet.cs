using System.Collections.Generic;

namespace ExcelGenerator.Classes
{
    public class ExcelWorkSheet
    {
        public string Name { get; set; }
        public List<ExcelCellStyle> ExcelCellStyles { get; set; }
        public List<Cell> Cells { get; set; }
        public List<ExcelChartTypeLine> ChartLines { get; set; }

        
    }
}
