using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
namespace ReadExcel
{
    class ExcellReader
    {
        
        public static void ReadFile()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Burak Dal\source\repos\ReadExcel\ReadExcel\bin\Debug\Meta.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int colCount = xlRange.Columns.Count;
            int rowCount = xlRange.Rows.Count;
            for (int i = 2; i <= rowCount; i++)
            {
                string 
                    id=null,sym1f=null,sym1=null,
                    sym2=null, sym2f = null, sym3=null, sym3f = null, 
                    sym4=null, sym4f = null, sym5 =null, sym5f = null, 
                    sym6 =null, sym6f = null;

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    id = xlRange.Cells[i, 1].Value2.ToString();
                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                    sym1 = xlRange.Cells[i, 2].Value2.ToString();
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null)
                    sym1f = xlRange.Cells[i, 4].Value2.ToString();
                if (xlRange.Cells[i, 6] != null && xlRange.Cells[i, 6].Value2 != null)
                    sym2 = xlRange.Cells[i, 6].Value2.ToString();
                if (xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null)
                    sym2f = xlRange.Cells[i, 8].Value2.ToString();
                if (xlRange.Cells[i, 10] != null && xlRange.Cells[i,10].Value2 != null)
                    sym3 = xlRange.Cells[i, 10].Value2.ToString();
                if (xlRange.Cells[i, 12] != null && xlRange.Cells[i,12].Value2 != null)
                    sym3f = xlRange.Cells[i, 12].Value2.ToString();
                if (xlRange.Cells[i, 15] != null && xlRange.Cells[i, 15].Value2 != null)
                    sym4 = xlRange.Cells[i, 15].Value2.ToString();
                if (xlRange.Cells[i, 17] != null && xlRange.Cells[i, 17].Value2 != null)
                    sym4f = xlRange.Cells[i, 17].Value2.ToString();
                if (xlRange.Cells[i, 19] != null && xlRange.Cells[i, 19].Value2 != null)
                    sym5 = xlRange.Cells[i, 19].Value2.ToString();
                if (xlRange.Cells[i, 21] != null && xlRange.Cells[i, 21].Value2 != null)
                    sym5f = xlRange.Cells[i, 21].Value2.ToString();
                if (xlRange.Cells[i, 23] != null && xlRange.Cells[i,23].Value2 != null)
                    sym6 = xlRange.Cells[i, 23].Value2.ToString();
                if (xlRange.Cells[i, 25] != null && xlRange.Cells[i,25].Value2 != null)
                    sym6f = xlRange.Cells[i, 23].Value2.ToString();

                var entry = new EntryTriangle(sym1, sym2, sym3,sym1f,sym2f,sym3f);
                var close= new CloseTriangle(sym4, sym5, sym6, sym4f, sym5f, sym6f);
                var row = new Row(entry, close, id);
                RowList.Add(row);

            }
        }
    }
}
