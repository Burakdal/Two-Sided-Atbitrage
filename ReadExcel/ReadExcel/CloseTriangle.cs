using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class CloseTriangle : SymbolTriangle
    {
        public CloseTriangle(string sym1, string sym2, string sym3,string sym1f,string sym2f,string sym3f) 
            : base(sym1, sym2, sym3,sym1f,sym2f,sym3f)
        {
        }

        public override double CalculateArbitrage()
        {
            return base.CalculateArbitrage();
        }
    }
}
