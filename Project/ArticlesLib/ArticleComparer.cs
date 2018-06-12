using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArticlesLib {
    public class ArticleComparer : EqualityComparer<Article> {
        public override bool Equals(Article x, Article y) {
            if (x == null || y == null)
                return x == y;
            return x.Title.Equals(y.Title) && x.Source.Equals(y.Source);
        }

        public override int GetHashCode(Article obj) {
            return obj == null ? 0 : (obj.Title.GetHashCode() ^ obj.Source.GetHashCode());
        }
    }
}
