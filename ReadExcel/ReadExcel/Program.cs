using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            
            ExcellReader.ReadFile();
            Console.Write("readed");
            Console.ReadLine();

        }
    }
}
