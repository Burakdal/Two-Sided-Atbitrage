using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class  RowList
    {
        private static List<Row> rows = new List<Row>();

        public static void Add(Row row)
        {
            rows.Add(row);
        }
        public static Row Get(int i)
        {
            return rows[i];
        }
        public static int initalizeList()
        {
            ExcellReader.ReadFile();
            return rows.Count;
        }
    }
}
