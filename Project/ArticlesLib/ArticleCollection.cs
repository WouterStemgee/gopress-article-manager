using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;


namespace ArticlesLib {
    public class ArticleCollection : IArticleCollection {
        public HashSet<Article> collection = new HashSet<Article>(new ArticleComparer());
        public List<Article> ModifiedItems { get; set; }
        public string name = "";

        public ArticleCollection() { }

        public ArticleCollection(string filename) {
            Import(filename);
        }

        public HashSet<Article> Articles {
            get {
                return collection;
            }
        }

        public string Name {
            get {
                return name;
            }
            set {
                name = value;
            }
        }

        public void Import(string filename) {
            string title = "";
            try {
                Article currentArticle;
                StringBuilder content;
                XDocument doc = XDocument.Load(filename);
                string rootLocalName = doc.Root.Name.LocalName;
                if (rootLocalName != "ArticleCollection") {
                    IEnumerable<XElement> articles = doc.Root.Elements("article.published");
                    if (rootLocalName != "articles.published") {
                        articles = articles.Concat(new[] { doc.Root });
                    }
                    int count = 0;
                    foreach (XElement article in articles) {
                        Console.Out.WriteLine(count++);
                        content = new StringBuilder();
                        foreach (XElement line in article.Element("body").Element("body.content").Element("body.p").Elements("p")) {
                            content.AppendLine(line.Value);
                        }
                        int page = 0;
                        int.TryParse(article.Element("body").Element("body.head").Attribute("pageNumber").Value, out int isArticle);
                        if (isArticle != 0) {
                            page = isArticle;
                        }
                        currentArticle = new Article {
                            Language = article.Attribute(XNamespace.Xml + "lang").Value,
                            Title = article.Element("body").Element("body.head").Element("headline").Element("hl1").Element("p").Value,
                            Publisher = article.Element("head").Element("id").Attribute("scope").Value,
                            Source = article.Element("head").Element("id").Attribute("title").Value,
                            Page = page,
                            DateString = DateTime.ParseExact(article.Element("head").Element("id").Attribute("pubdate").Value, "yyyy-MM-ddTHH:mm:ss.FFFzzz", System.Globalization.CultureInfo.InvariantCulture).ToShortDateString(),
                            Content = content.ToString(),
                            Author = article.Element("body").Element("body.end").Element("tagline").Element("authortagline").Element("p").Value
                        };
                        title = article.Element("body").Element("body.head").Element("headline").Element("hl1").Element("p").Value;
                        collection.Add(currentArticle);
                    }
                } else {
                    ArticleCollection ac = default(ArticleCollection);
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(filename);
                    string xmlString = xmlDocument.OuterXml;
                    using (StringReader read = new StringReader(xmlString)) {
                        Type outType = typeof(ArticleCollection);
                        XmlSerializer serializer = new XmlSerializer(outType);
                        using (XmlReader reader = new XmlTextReader(read)) {
                            ac = (ArticleCollection)serializer.Deserialize(reader);
                            reader.Close();
                        }
                        read.Close();
                    }
                    collection = ac.Articles;
                }
            } catch (Exception e) {

                Console.Out.WriteLine(e.Message + " laatste title = " + title);
            }
        }

        public void Print() {
            Dictionary<string, HashSet<Article>> dict = new Dictionary<string, HashSet<Article>>();
            foreach (Article a in collection) {
                if (!dict.ContainsKey(a.Source)) {
                    dict[a.Source] = new HashSet<Article>(new ArticleComparer());
                }
                dict[a.Source].Add(a);
            }
            foreach (KeyValuePair<string, HashSet<Article>> entry in dict) {
                Console.Out.WriteLine("---ARTICLES FROM: " + entry.Key + " (" + entry.Value.Count + ")");
                foreach (Article a in entry.Value) {
                    Console.Out.WriteLine("\t======================================");
                    Console.Out.WriteLine("\tTitle=" + a.Title);
                    Console.Out.WriteLine("\tLanguage=" + a.Language);
                    Console.Out.WriteLine("\tAuthor=" + a.Author);
                    Console.Out.WriteLine("\tPublisher=" + a.Publisher);
                    Console.Out.WriteLine("\tSource=" + a.Source);
                    Console.Out.WriteLine("\tPage=" + a.Page);
                    Console.Out.WriteLine("\tWords=" + a.Words);
                    Console.Out.WriteLine("\tDate=" + a.Date);
                    Console.Out.WriteLine("\t======================================");
                    Console.Out.WriteLine();
                }
            }
        }
    }
}
