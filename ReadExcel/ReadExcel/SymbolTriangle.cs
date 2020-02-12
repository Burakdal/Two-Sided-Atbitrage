using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class SymbolTriangle
    {
        public Symbol Symbol1 { get; set; }
        public Symbol Symbol2 { get; set; }
        public Symbol Symbol3 { get; set; }
        public string Symbol1FactorAsset,
            
            Symbol2FactorAsset, 
            
            Symbol3FactorAsset;
        public bool isOpen { get; set; }
        public SymbolTriangle(string sym1,string sym2,string sym3,string sym1f,string sym2f,string sym3f)
        {
            
            this.Symbol1FactorAsset = ControlFactorValue(sym1f);
            this.Symbol2FactorAsset = ControlFactorValue(sym2f);
            this.Symbol3FactorAsset = ControlFactorValue(sym3f);
            this.Symbol1 = new Symbol(sym1);
            this.Symbol2 = new Symbol(sym2);
            this.Symbol3 = new Symbol(sym3);
            this.Symbol1.MarketData.Quantity = 1;
            this.Symbol2.MarketData.Quantity = 1;
            this.Symbol3.MarketData.Quantity = 1;
        }

        public virtual double CalculateArbitrage()
        {
            throw new NotImplementedException();
            return 1;
        }
        
        public virtual string ControlFactorValue(string factor)
        {
            if(factor!="1")
            {
                return factor.Substring(6);
            }
            return factor;
        }
        public virtual void SetFactorValue()
        {   

            if (this.Symbol2FactorAsset != "1")
            {
                this.Symbol2.MarketData.Quantity = this.Symbol3.MarketData.Bid;
                return;
            }
            else if(this.Symbol3FactorAsset != "1")
            {
                this.Symbol3.MarketData.Quantity = this.Symbol2.MarketData.Bid;
                return;
            }
        }


    }
}
