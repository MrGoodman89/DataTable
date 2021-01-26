using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataTable_Intima_
{
    class DataComparer: IComparer<List<string>>
    {
        public int Compare(List<string> x, List<string> y)
        {
            return x[2].CompareTo(y[2]);
        }
    }
}
