using Newtonsoft.Json;

namespace ReadExcel
{
     public class Helper
     {
        public static int FillList()
        {
            ExcelWriter.initalizeDocument();
            return RowList.initalizeList();
        }
        public static string GetRow(int i)
        {
            string output = "Entry";
            Row row=RowList.Get(i);
            output += ","+row.EntryTriangle.Symbol1.Name;

            output += "," + row.EntryTriangle.Symbol2.Name;

            output += "," + row.EntryTriangle.Symbol3.Name;

            return output;
        }
        public static string MapRow(int rowIndex,string date,
            double sm1Ask,double sm1Bid,double sm2Ask,double sm2Bid,double sm3Ask,double sm3Bid)
        {
            
            Row row = RowList.Get(rowIndex);
            row.EntryTriangle.Symbol1.MarketData.Ask = sm1Ask;
            row.EntryTriangle.Symbol1.MarketData.Bid = sm1Bid;
            row.EntryTriangle.Symbol2.MarketData.Ask = sm2Ask;
            row.EntryTriangle.Symbol2.MarketData.Bid = sm2Bid;
            row.EntryTriangle.Symbol3.MarketData.Ask = sm3Ask;
            row.EntryTriangle.Symbol3.MarketData.Bid = sm3Bid;
            row.EntryTriangle.SetFactorValue();
            ExcelWriter.AddRow(rowIndex,date);
            if (row.EntryTriangle.CalculateArbitrage() > 0 && !row.EntryTriangle.isOpen)
            {
                return row.EntryTriangle.Symbol1.Name + "," + row.EntryTriangle.Symbol1FactorAsset + "," +
                    row.EntryTriangle.Symbol2.Name + "," + row.EntryTriangle.Symbol2FactorAsset + ","
                    + row.EntryTriangle.Symbol3.Name + "," + row.EntryTriangle.Symbol3FactorAsset;
            }
            else
            {
                return "-1";
            }
            
           

        }
        public static void OpenPosition(int index,string date,
            double sm1Ask, double sm1Bid, double sm2Ask, double sm2Bid, double sm3Ask, double sm3Bid)
        {
            Row row = RowList.Get(index);
            if (!row.EntryTriangle.isOpen)
            {
                MapRow(index,date, sm1Ask,sm1Bid,sm2Ask, sm2Bid,sm3Ask, sm3Bid);
                row.EntryTriangle.isOpen = true;
            }
        }

        public static string CloseRobot()
        {
            string output = ExcelWriter.saveExcell();
            return output;
        }
     }
}
