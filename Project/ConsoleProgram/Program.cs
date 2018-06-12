using ArticlesLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleProgram {
    public class Program {
        public static void Main(string[] args) {
            IArticleCollection collection = new ArticleCollection();
            try {
                collection.Import(@"D:\Dropbox\MAPR Michiel\Project\hugo_claus.xml");
                collection.Print();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Out.Write("[SUCCESS]");
                Console.ResetColor();
            } catch (Exception e) {
                PrintError(e.Message);
            }
            Console.ReadLine();
        }

        public static void PrintError(string errorMessage) {
            Console.Out.WriteLine();
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.Out.Write("[ERR] ");
            Console.ResetColor();
            Console.Out.WriteLine(errorMessage);
        }
    }
}
