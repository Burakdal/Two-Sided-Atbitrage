#import "ReadExcel.dll"
#import
#include <Trade\Trade.mqh>
#include "..\Include\FileOp.mqh";

CTrade m_trade;
int lengthOfSymbols;
File file;
int OnInit()
  {
    lengthOfSymbols=Helper::FillList();
    return(INIT_SUCCEEDED);
  }

void OnTick()
  {  
     for(int i=0;i<lengthOfSymbols;i++)
     {
        double priceOutput[6];
        string output=Helper::GetRow(i);
        ushort u_sep=StringGetCharacter(",",0);
        string result[];
        int l=StringSplit(output,u_sep,result);
        string allPricess="";
        if(result[0]=="Entry")
        {
            for(int i=1;i<ArraySize(result);i++)
            {
                  double Ask = SymbolInfoDouble(result[i], SYMBOL_ASK);
                  double Bid = SymbolInfoDouble(result[i], SYMBOL_BID);
                  priceOutput[i*2-1]=Ask;
                  priceOutput[i*2-2]=Bid;
            }
            
        }
        Print(result[1]," ",result[2]," ",result[3]);
        string date=TimeToString(TimeCurrent(),TIME_DATE|TIME_MINUTES);
        string calculationReturn=Helper::MapRow(
        i,date,priceOutput[0] ,priceOutput[1] ,priceOutput[2],priceOutput[3],priceOutput[4],priceOutput[5]
        );
        if(calculationReturn!="-1")
        {
            //ushort u_sep=StringGetCharacter(",",0);
            //string result[];
            //int l=StringSplit(calculationReturn,u_sep,result);
            //m_trade.Buy(NormalizeDouble(StringToDouble(result[1]),2), result[0]);
            //double price1=m_trade.ResultPrice();
            //m_trade.Sell(NormalizeDouble(StringToDouble(result[3]),2), result[2]);
            //double price2=m_trade.ResultPrice();
            //m_trade.Sell(NormalizeDouble(StringToDouble(result[5]),2), result[4]);
            //double price3=m_trade.ResultPrice();
            //Helper::OpenPosition(i,price1 ,price2 ,price3);
        }
      }
      
     //Print(priceOutput);
  }
  
void OnDeinit(const int reason)
 {
   string totalOutput=Helper::CloseRobot();
   string name = TimeToString(TimeCurrent(),TIME_DATE|TIME_MINUTES);
   StringReplace(name, ":", ".");
   file.WriteAllText("Log"+name+".csv",totalOutput);
   Print("Finished");

 }
  

