using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class Row
    {
        public EntryTriangle EntryTriangle { get; set; }
        public CloseTriangle CloseTriangle { get; set; }
        public string  Id { get; set; }

        public Row(EntryTriangle entry,CloseTriangle close,string id)
        {
            this.EntryTriangle = entry;
            this.CloseTriangle = close;
            this.Id = id;
        }

        

    }
}
