using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ArticlesLib {
    public class Date : IComparable {
        string value;
        private int day, month, year;
        public Date() { }
        public Date(string _value) {
            this.value = _value;
            string pattern = @"^(\d+)/(\d+)/(\d+)";
            Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
            Match m = r.Match(value);
            for (int i = 1; i <= 3; i++) {
                Group g = m.Groups[i];
                CaptureCollection cc = g.Captures;
                Capture c = cc[0];
                if (i == 1)
                    day = Int32.Parse(c.Value);
                else if (i == 2)
                    month = Int32.Parse(c.Value);
                else if (i == 3)
                    year = Int32.Parse(c.Value);
            }
        }
        public static implicit operator string(Date d) {
            return d.ToString();
        }
        public static implicit operator Date(string d) {
            return new Date(d);
        }

        public override string ToString() {
            return value;
        }

        public int Day {
            get {
                return day;
            }
        }

        public int Month {
            get {
                return month;
            }
        }

        public int Year {
            get {
                return year;
            }
        }

        int IComparable.CompareTo(object obj) {
            Date d = (Date)obj;
            IComparer c = new DateComparer();
            return c.Compare(this, d);
        }
    }
}
