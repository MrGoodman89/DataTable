using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace DataTable_Intima_
{
    class DatTable 
    {
        public DatTable(string DateTime, string TagType, string Type, string Value)
        {
            this.DateTime = DateTime;
            this.TagType = TagType;
            this.Type = Type;
            this.Value = Value;
        }

        public string DateTime { get; set; }
        public string TagType { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }
}
