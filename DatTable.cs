using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataTable_Intima_
{
    class DatTable
    {
        private string[,] list;

        public DatTable(string[,] list)
        {
            this.list = list;
        }

        public DatTable(string dateTime, string tagType, string type, string value)
        {
            this.dateTime = dateTime;
            this.tagType = tagType;
            this.type = type;
            this.value = value;
        }

        public string dateTime { get; set; }
        public string tagType { get; set; }
        public string type { get; set; }
        public string value { get; set; }
    }
}
