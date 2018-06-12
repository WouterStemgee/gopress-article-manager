using ArticlesLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ArticlesUI {
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class CollectionControl : UserControl {
        private ArticleCollection collection = new ArticleCollection();
        private GridViewColumnHeader listViewSortCol = null;
        private SortAdorner listViewSortAdorner = null;

        public CollectionControl(string name) {
            InitializeComponent();
            collection.Name = name;
            lvCollection.ItemsSource = collection.Articles;       
        }

        public CollectionControl(string name, ArticleCollection other) {
            InitializeComponent();
            collection.Name = other.Name;
            foreach(Article article in other.Articles) {
                collection.Articles.Add(article);
            }
            lvCollection.ItemsSource = collection.Articles;
        }

        public ArticleCollection Collection {
            get {
                return collection;
            }
        }

        private void lvCollection_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            Article selected = (Article)lvCollection.SelectedItem;
            if (selected != null) {
                MainWindow window = (MainWindow)Application.Current.MainWindow;
                window.lblTitle.Content = selected.Title;
                window.txtArticleContent.Document.Blocks.Clear();
                window.txtArticleContent.Document.Blocks.Add(new Paragraph(new Run(selected.Content)));
            }   
        }


        private void lvCollectionColumnHeader_Click(object sender, RoutedEventArgs e) {
            GridViewColumnHeader column = (sender as GridViewColumnHeader);
            string sortBy = column.Tag.ToString();
            if (listViewSortCol != null) {
                AdornerLayer.GetAdornerLayer(listViewSortCol).Remove(listViewSortAdorner);
                lvCollection.Items.SortDescriptions.Clear();
            }

            ListSortDirection newDir = ListSortDirection.Ascending;
            if (listViewSortCol == column && listViewSortAdorner.Direction == newDir)
                newDir = ListSortDirection.Descending;

            listViewSortCol = column;
            listViewSortAdorner = new SortAdorner(listViewSortCol, newDir);
            AdornerLayer.GetAdornerLayer(listViewSortCol).Add(listViewSortAdorner);
            lvCollection.Items.SortDescriptions.Add(new SortDescription(sortBy, newDir));
        }

    }
}
