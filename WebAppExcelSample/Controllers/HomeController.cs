﻿using ExcelGenerator;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Web.Mvc;
using WebAppExcelSample.Classes;


namespace WebAppExcelSample.Controllers
{
    public class HomeController : Controller
    {
     

        public HomeController()
        {

        }

               
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


		public ActionResult ExcelExport()
		{
			try
			{
				
				var fileBytes = TestLibrary();
				string fullFileName = "Testing-excel-generator.xlsx";
				Response.Headers.Add("Excel-File-Name", fullFileName);
				return File(fileBytes,
								"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
								fullFileName);


			}
			catch (Exception ex)
			{
				throw ex;
			}

			//IF YOU WANT TO SAVE IN A PHYSICAL PATH
			//string filePath = @"C:\Test\test-excel.xlsx";
			//string password = "SuperUltraClave";
			//excelPackage.SaveAs(new FileInfo(filePath), password);
		}


        private byte[] TestLibrary(){
            try
            {
				List<ExcelGenerator.Classes.ExcelCellStyle> excelCellStylesList = new List<ExcelGenerator.Classes.ExcelCellStyle>();
				ExcelGenerator.Classes.ExcelCellStyle titleCellStyle = new ExcelGenerator.Classes.ExcelCellStyle();
				ExcelGenerator.Classes.ExcelCellStyle tableHeaderStyle = new ExcelGenerator.Classes.ExcelCellStyle();

				titleCellStyle = new ExcelCellStyleCreator
				{
					BackgroundColor = Color.Blue,
					TextColor = Color.White,
					Fontbold = true,
					FontName = "Arial",
					FontSize = 14
				}.CreateStyle();

				tableHeaderStyle = new ExcelCellStyleCreator
				{
					BackgroundColor = Color.CadetBlue,
					TextColor = Color.White,
					Fontbold = false,
					FontName = "Arial",
					FontSize = 12
				}.CreateStyle();


				var builderExcel = new BuilderExcel<Example>
				{
					Title = "", //null or empty if you don't want a title
					SheetName = "Test-Excel",
					TableHeaderNames = "Item,Name,PageNumber",
					TableHeaderCellStyles = tableHeaderStyle,
					TitleCellStyles = titleCellStyle,
                    ColumnWidth = 25,
					listOfRecords = new List<Example>(GetAllexamples())
				
				};


				var result = builderExcel.BuilderExcelFile(1, 1);
				return result;
				//turn builderExcel.Builder(1, 1);
			}
            catch (Exception ex)
            {

                throw ex;
            }
		

        }

        private List<Example> GetAllexamples()
        {
            List<Example> cellList = new List<Example>();
            cellList.Add(new Example { Id = 1, NameExample = "Alfa", PageNumber = 95 });
            cellList.Add(new Example { Id = 2, NameExample = "Beta", PageNumber = 96 });
            cellList.Add(new Example { Id = 3, NameExample = "Delta", PageNumber = 97 });
            cellList.Add(new Example { Id = 4, NameExample = "Gamma", PageNumber = 98 });

            return cellList;
        }

	


	}
}