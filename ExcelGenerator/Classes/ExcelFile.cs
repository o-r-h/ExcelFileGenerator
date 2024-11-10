using System.Collections.Generic;

namespace ExcelGenerator.Classes
{
    public class ExcelFile
    {
        public const string ExcelContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public string ExcelFileName { get; set; }
        public string PrefixSheetName { get; set; }
        public List<ExcelWorkSheet> ExcelWorkSheets { get; set; }

    }
}
