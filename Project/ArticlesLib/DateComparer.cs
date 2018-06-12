using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArticlesLib {
    public class DateComparer : IComparer {
        public int Compare(object a, object b) {
            Date x = (Date)a;
            Date y = (Date)b;
            if ((x.Year == y.Year) && (x.Month == y.Month) && (x.Day == y.Day))
                return 0;
            if ((x.Year < y.Year) || (x.Year == y.Year) && (x.Month < y.Month) || (x.Year == y.Year) && (x.Month == y.Month) && (x.Day < y.Day))
                return -1;
            return 1;
        }
    }
}
