using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArticlesLib {
    public interface IArticle {
        string Title { get; set; }
        string Author { get; set; }
        string Language { get; set; }
        string Publisher { get; set; }
        string Content { get; set; }
        string Source { get; set; }
        Date Date { get; }
        string DateString { get; set; }
        int Page { get; set; }
        int Words { get; }
    }
}
