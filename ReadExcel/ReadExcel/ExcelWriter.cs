using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ReadExcel
{
    class ExcelWriter
    {

        //private static Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //private static Excel.Workbook xlWorkBook;
        //private static Excel.Worksheet xlWorkSheet;
        //private static int lastIndex = 0;
        //private static object misValue;
        private static StringBuilder sb = new StringBuilder();
        public static void initalizeDocument()
        {
            sb.Append("Date;Symbol1;Ask1;Bid1;Quantity1;" +
            "Symbol2;Ask1;Bid2;Quantitiy2;Symbol3;Ask3;Bid3;Quanity3;Return\n");

            //object misValue = System.Reflection.Missing.Value;
            //xlWorkBook = xlApp.Workbooks.Add(misValue);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //xlWorkSheet.Cells[1, 1] = "Symbol1";
            //xlWorkSheet.Cells[1, 2] = "Ask1";
            //xlWorkSheet.Cells[1, 3] = "Quantity1";
            //xlWorkSheet.Cells[1, 4] = "Symbol2";
            //xlWorkSheet.Cells[1, 5] = "Bid2";
            //xlWorkSheet.Cells[1, 6] = "Quantity2";
            //xlWorkSheet.Cells[1, 7] = "Symbol3";
            //xlWorkSheet.Cells[1, 8] = "Bid3";
            //xlWorkSheet.Cells[1, 9] = "Quantity3";
            //xlWorkSheet.Cells[1, 10] = "Return";
            //lastIndex = 1;
        }

        public static void AddRow(int index,string date)
        {
            Row row = RowList.Get(index);
            string rowS = date+";"+row.EntryTriangle.Symbol1.Name + ";" +
                row.EntryTriangle.Symbol1.MarketData.Ask.ToString() + ";" +
                row.EntryTriangle.Symbol1.MarketData.Bid.ToString() + ";" +
                row.EntryTriangle.Symbol1.MarketData.Quantity.ToString() + ";" +
                row.EntryTriangle.Symbol2.Name.ToString() + ";" +
                row.EntryTriangle.Symbol2.MarketData.Ask.ToString() + ";" +
                row.EntryTriangle.Symbol2.MarketData.Bid.ToString() + ";" +
                row.EntryTriangle.Symbol2.MarketData.Quantity.ToString() + ";" +
                row.EntryTriangle.Symbol3.Name.ToString() + ";" +
                row.EntryTriangle.Symbol3.MarketData.Ask.ToString() + ";" +
                row.EntryTriangle.Symbol3.MarketData.Bid.ToString() + ";" +
                row.EntryTriangle.Symbol3.MarketData.Quantity.ToString() + ";"+row.EntryTriangle.CalculateArbitrage().ToString() + "\n";
            sb.Append(rowS);




            //lastIndex += 1;
            //xlWorkSheet.Cells[lastIndex, 1] = row.EntryTriangle.Symbol1.Name;
            //xlWorkSheet.Cells[lastIndex, 2] = row.EntryTriangle.Symbol1.MarketData.Ask;
            //xlWorkSheet.Cells[lastIndex, 3] = row.EntryTriangle.Symbol1FactorAsset;
            //xlWorkSheet.Cells[lastIndex, 4] = row.EntryTriangle.Symbol2.Name;
            //xlWorkSheet.Cells[lastIndex, 5] = row.EntryTriangle.Symbol2.MarketData.Bid;
            //xlWorkSheet.Cells[lastIndex, 6] = row.EntryTriangle.Symbol2FactorAsset;
            //xlWorkSheet.Cells[lastIndex, 7] = row.EntryTriangle.Symbol3.Name;
            //xlWorkSheet.Cells[lastIndex, 8] = row.EntryTriangle.Symbol3.MarketData.Bid;
            //xlWorkSheet.Cells[lastIndex, 9] = row.EntryTriangle.Symbol3FactorAsset;
            //xlWorkSheet.Cells[lastIndex, 10] = row.EntryTriangle.CalculateArbitrage();
        }

        public static string saveExcell()
        { 
            //string time = DateTime.Now.ToString("yyyy_mm_dd HH:MM:ss");
            //xlWorkBook.SaveAs("D:\\" + time + ".xlsx", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);
            return sb.ToString();
        }

    }
}
