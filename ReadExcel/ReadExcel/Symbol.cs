using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class Symbol
    {
        public string Name { get; set; }
        public MarketData MarketData { get; set; }
        public Symbol(string name)
        {
            this.Name = name;
            this.MarketData = new MarketData();

        }
        
    }
}
