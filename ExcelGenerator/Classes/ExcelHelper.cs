using System;
using System.Collections.Generic;
using System.Reflection;


namespace ExcelGenerator.Classes
{
    public class ExcelHelper
    {

        public static int GetQuantityFieldInClass<T>()
        {
            int result = 0;
            Type t = typeof(T);
            result = t.GetProperties().Length;
            return result;
        }


        public static List<Cell> CreateCellTable<T>(int initialCellRowPos, int initialColRowPos, List<T> listOfRecords)
        {
            int ipos = initialCellRowPos;
            int jpos = initialColRowPos;
            int colNbr = ExcelHelper.GetQuantityFieldInClass<T>();
            List<Cell> lista = new List<Cell>();

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
                    lista.Add(c);
                }
                jpos = initialColRowPos;
                ipos++;
            }

            return lista;
        }
    }
}
