using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArticlesLib {
    public class Article : IArticle {
        private string title;
        private string author;
        private string language;
        private string publisher;
        private string source;
        private string content;
        private Date date;
        private string dateString;
        private int page;
        private int words;

        public Article() { }

        public string Title {
            get {
                return title;
            }
            set {
                this.title = value;
            }
        }

        public string Author {
            get {
                return author;
            }
            set {
                this.author = value;
            }
        }

        public string Language {
            get {
                return language;
            }
            set {
                this.language = value;
            }
        }

        public string Publisher {
            get {
                return publisher;
            }
            set {
                this.publisher = value;
            }
        }

        public string DateString {
            get {
                return dateString;
            }
            set {
                this.dateString = value;
                this.date = new Date(value);
            }
        }

        public Date Date {
            get {
                return date;
            }
        }

        public string Source {
            get {
                return source;
            }
            set {
                this.source = value;
            }
        }

        public int Page {
            get {
                return page;
            }
            set {
                this.page = value;
            }
        }

        public string Content {
            get {
                return content;
            }
            set {
                this.content = value;
                this.words = this.content.Split(new char[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;
            }
        }

        public int Words {
            get {
                return words;
            }
        }

    }
}
