using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArticlesLib {
    public interface IArticleCollection {      
        void Import(string filename);
        void Print();
        HashSet<Article> Articles { get; }
        string Name { get; set; }
    }
}
