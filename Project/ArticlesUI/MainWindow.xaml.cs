using ArticlesLib;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml;
using System.Xml.Serialization;

namespace ArticlesUI {

    public partial class MainWindow : System.Windows.Window {

        private ArticleCollection sourceCollection = new ArticleCollection();
        private GridViewColumnHeader listViewSortCol = null;
        private SortAdorner listViewSortAdorner = null;

        public MainWindow() {
            InitializeComponent();
            InitializeCollections();
        }

        #region Procedures
        public void InitializeCollections() {
            System.IO.Directory.CreateDirectory(@"C:\Articles\");
            System.IO.Directory.CreateDirectory(@"C:\Articles\exports\");
            System.IO.Directory.CreateDirectory(@"C:\Articles\collections\");
            string[] filePaths = Directory.GetFiles(@"C:\Articles\collections\", "*.xml");
            foreach (string fileName in filePaths) {
                ArticleCollection collection = DeSerializeCollection<ArticleCollection>(fileName);
                if (collection != null) {
                    TabItem tab = new TabItem { Content = new CollectionControl(collection.Name, collection), Header = collection.Name };
                    tabControl.Items.Add(tab);
                    tabControl.SelectedItem = tab;
                    CollectionControl control = (CollectionControl)tabControl.SelectedContent;
                    if (control != null) {
                        control.lvCollection.Items.Refresh();
                    }
                }
            }
            ArticleCollection src = DeSerializeCollection<ArticleCollection>(@"C:\Articles\source.xml");
            if (src != null) {
                foreach (Article article in src.Articles) {
                    sourceCollection.Articles.Add(article);
                }
                lvSource.ItemsSource = sourceCollection.Articles;
                lvSource.Items.Refresh();
                if (lvSource.Items.Count != 0)
                    lvSource.SelectedItem = lvSource.Items[0];
            }

            if (sourceCollection.Articles.Count == 0) {
                sourceCollection = new ArticleCollection();
                lvSource.ItemsSource = sourceCollection.Articles;
                lvSource.Items.Refresh();
                if (lvSource.Items.Count != 0)
                    lvSource.SelectedItem = lvSource.Items[0];
            }

            lvSource.ItemsSource = sourceCollection.Articles;
            lvSource.Items.Refresh();
        }

        public void ImportXML() {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog {
                DefaultExt = ".xml",
                Filter = "XML files (*.xml)|*.xml",
                Multiselect = true,
                Title = "Importeer XML-bestanden"
            };
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true) {
                lblStatus.Text = "Importing XML File...";
                try {
                    int empty = 0;
                    if (sourceCollection.Articles.Count == 0) {
                        sourceCollection = new ArticleCollection(dlg.FileNames[0]);
                        empty = 1;
                    }

                    for (int i = empty; i < dlg.FileNames.Length; i++) {
                        ArticleCollection col = new ArticleCollection(dlg.FileNames[i]);
                        foreach (Article a in col.Articles) {
                            sourceCollection.Articles.Add(a);
                        }
                    }
                    lvSource.ItemsSource = sourceCollection.Articles;
                    lvSource.Items.Refresh();
                    if (lvSource.Items.Count != 0)
                        lvSource.SelectedItem = lvSource.Items[0];
                    lblStatus.Text = "Done (Total of " + sourceCollection.Articles.Count + " articles loaded)";

                } catch (Exception e) {
                    lblStatus.Text = e.Message;
                }
            }
        }

        public void MoveToCollection() {
            int index = lvSource.SelectedIndex;
            List<Article> selected = new List<Article>();
            foreach (Article article in lvSource.SelectedItems) {
                selected.Add(article);
            }
            CollectionControl control = (CollectionControl)tabControl.SelectedContent;
            if (control != null && selected.Count != 0) {
                foreach (Article article in selected) {
                    control.Collection.Articles.Add(article);
                    sourceCollection.Articles.Remove(article);
                }
                control.lvCollection.Items.Refresh();
                lvSource.Items.Refresh();
                lvSource.SelectedIndex = index;
                lblStatus.Text = "Article(s) added to collection: " + control.Collection.Name;
            }
        }

        public void MoveToSource() {
            CollectionControl control = (CollectionControl)tabControl.SelectedContent;
            if (control != null) {
                int index = control.lvCollection.SelectedIndex;
                List<Article> selected = new List<Article>();
                foreach (Article article in control.lvCollection.SelectedItems) {
                    selected.Add(article);
                }
                if (selected.Count != 0) {
                    foreach (Article article in selected) {
                        control.Collection.Articles.Remove(article);
                        sourceCollection.Articles.Add(article);
                    }
                    control.lvCollection.Items.Refresh();
                    lvSource.Items.Refresh();
                    control.lvCollection.SelectedIndex = index;
                    lblStatus.Text = "Article(s) removed from collection: " + control.Collection.Name;
                }
            }
        }

        public void DeleteSourceArticle() {
            int index = lvSource.SelectedIndex;
            List<Article> selected = new List<Article>();
            foreach (Article article in lvSource.SelectedItems) {
                selected.Add(article);
            }
            if (selected.Count != 0) {
                foreach (Article article in selected) {
                    sourceCollection.Articles.Remove(article);
                }
                lvSource.Items.Refresh();
                lvSource.SelectedIndex = index;
                lblStatus.Text = "Article(s) removed from source collection";
            }
        }

        public void DeleteCollectionArticle() {
            CollectionControl control = (CollectionControl)tabControl.SelectedContent;
            if (control != null) {
                int index = control.lvCollection.SelectedIndex;
                List<Article> selected = new List<Article>();
                foreach (Article article in control.lvCollection.SelectedItems) {
                    selected.Add(article);
                }
                if (selected.Count != 0) {
                    foreach (Article article in selected) {
                        control.Collection.Articles.Remove(article);
                    }
                    control.lvCollection.Items.Refresh();
                    control.lvCollection.SelectedIndex = index;
                    lblStatus.Text = "Article(s) removed from collection: " + control.Collection.Name;
                }
            }
        }

        public void AddCollection(string collectionName = "") {
            if (collectionName.Equals("")) {
                InputDialog inputDialog = new InputDialog("Geef een naam voor de nieuwe collectie:", collectionName);
                if (inputDialog.ShowDialog() == true && inputDialog.DialogResult == true) {
                    collectionName = inputDialog.Answer;
                    TabItem tab = new TabItem { Content = new CollectionControl(collectionName), Header = collectionName };
                    tabControl.Items.Add(tab);
                    tabControl.SelectedItem = tab;
                }
            }
        }

        public void RemoveCollection() {
            CollectionControl control = (CollectionControl)tabControl.SelectedContent;
            try {
                TabItem ti = tabControl.SelectedItem as TabItem;

                System.IO.File.Delete(@"C:\Articles\collections\" + ti.Header + ".xml");
            } catch (Exception e) {
                lblStatus.Text = e.Message;
            }
            tabControl.Items.Remove(tabControl.SelectedItem);
        }

        public void SaveCollections() {
            foreach (TabItem item in tabControl.Items) {
                CollectionControl control = (CollectionControl)item.Content;
                if (control != null) {
                    string fileName = @"C:\Articles\collections\" + control.Collection.Name + ".xml";
                    SerializeCollection<ArticleCollection>(control.Collection, fileName);
                }
            }
            SerializeCollection<ArticleCollection>(sourceCollection, @"C:\Articles\source.xml");
        }

        public void ExportCollections() {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            try {
                if (xlApp != null) {
                    xlApp.Visible = true;
                    Workbook wb = xlApp.Workbooks.Add(XlSheetType.xlWorksheet);
                    Worksheet ws1 = wb.ActiveSheet as Worksheet;
                    ws1.Name = "Source";
                    ws1.Cells[1, 1] = "Something";

                    foreach (TabItem item in tabControl.Items) {
                        CollectionControl control = (CollectionControl)item.Content;
                        if (control != null) {
                            Worksheet ws = wb.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing) as Worksheet;
                            ws.Name = control.Collection.Name;

                            // Alle aanwezige jaartallen ophalen uit collectie
                            List<int> jaartallen = new List<int>();
                            foreach (Article a in control.Collection.Articles) {
                                int jaartal = Int32.Parse(a.Date.ToString().Substring(a.Date.ToString().Length - 4, 4));
                                if (!jaartallen.Contains(jaartal)) {
                                    jaartallen.Add(jaartal);
                                }
                            }
                            jaartallen.Sort();
                            Dictionary<int, int> jaartal_indices = new Dictionary<int, int>();
                            for (int i = 0; i < jaartallen.Count; i++) {
                                jaartal_indices[jaartallen[i]] = i;
                            }
                            for (int i = 0; i < jaartallen.Count; i++) {
                                ws.Cells[1, i + 2] = jaartallen[i];
                            }
                            ws.Cells[1, jaartallen.Count + 2] = "Totaal";

                            // Alle data in collectie inlezen naar SortedDictionary container
                            SortedDictionary<string, SortedDictionary<int, int>> collection_data = new SortedDictionary<string, SortedDictionary<int, int>>();
                            foreach (Article a in control.Collection.Articles) {
                                string publisher = a.Source;
                                int jaartal = Int32.Parse(a.Date.ToString().Substring(a.Date.ToString().Length - 4, 4));

                                if (!collection_data.ContainsKey(publisher)) {
                                    collection_data[publisher] = new SortedDictionary<int, int>();
                                }
                                if (!collection_data[publisher].ContainsKey(jaartal)) {
                                    collection_data[publisher][jaartal] = 0;
                                }
                                collection_data[publisher][jaartal] = collection_data[publisher][jaartal] + 1;
                            }

                            // Export data van in SortedDictionary naar Excel cells
                            int index = 0;
                            foreach (KeyValuePair<string, SortedDictionary<int, int>> krant in collection_data) {
                                ws.Cells[index + 2, 1] = krant.Key;
                                int aantal_artikels_per_krant = 0;
                                for (int fill = 0; fill < jaartallen.Count; fill++) {
                                    ws.Cells[index + 2, fill + 2] = 0;
                                }
                                foreach (KeyValuePair<int, int> jaartal in collection_data[krant.Key]) {
                                    ws.Cells[index + 2, jaartal_indices[jaartal.Key] + 2] = jaartal.Value;
                                    aantal_artikels_per_krant += jaartal.Value;
                                }
                                ws.Cells[index + 2, jaartal_indices.Count + 2] = aantal_artikels_per_krant;
                                index++;
                            }
                        }
                    }
                    wb.SaveAs(@"C:\Articles\exports\collections_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm-tt") + ".xlsx");
                    //wb.Close(Type.Missing, Type.Missing, Type.Missing);
                    //xlApp.UserControl = true;
                    //xlApp.Quit();
                    lblStatus.Text = "Collections successfully exported to Excel file.";
                }
            } catch (Exception e) {
                Console.Out.WriteLine(e.Message);
            } finally {
                Marshal.ReleaseComObject(xlApp);
            }
        }

        public void SearchText() {
            int matches = 0;
            string keyword = txtSearch.Text;
            if (keyword.Length >= 3) {
                TextRange text = new TextRange(txtArticleContent.Document.ContentStart, txtArticleContent.Document.ContentEnd);
                TextPointer current = text.Start.GetInsertionPosition(LogicalDirection.Forward);
                while (current != null) {
                    string textInRun = current.GetTextInRun(LogicalDirection.Forward);
                    if (!string.IsNullOrWhiteSpace(textInRun)) {
                        int index = textInRun.ToLower().IndexOf(keyword.ToLower());
                        if (index != -1) {
                            matches++;
                            TextPointer selectionStart = current.GetPositionAtOffset(index, LogicalDirection.Forward);
                            TextPointer selectionEnd = selectionStart.GetPositionAtOffset(keyword.Length, LogicalDirection.Forward);
                            TextRange selection = new TextRange(selectionStart, selectionEnd);
                            selection.ApplyPropertyValue(TextElement.BackgroundProperty, Brushes.Yellow);
                            selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);

                        }
                    }
                    current = current.GetNextContextPosition(LogicalDirection.Forward);
                }
                lblStatus.Text = "Search complete! Found " + matches + " matches for keyword: \"" + keyword + "\".";
            } else {
                lblStatus.Text = "Error: keyword must be longer than 2 characters.";
            }
        }

        public void SerializeCollection<ArticleCollection>(ArticleCollection ac, string fileName) {
            if (ac == null) { return; }
            try {
                XmlDocument xmlDocument = new XmlDocument();
                XmlSerializer serializer = new XmlSerializer(ac.GetType());
                using (MemoryStream stream = new MemoryStream()) {
                    serializer.Serialize(stream, ac);
                    stream.Position = 0;
                    xmlDocument.Load(stream);
                    xmlDocument.Save(fileName);
                    stream.Close();
                    lblStatus.Text = "Collections saved successfully.";
                }
            } catch (Exception e) {
                lblStatus.Text = e.Message;
            }
        }

        public ArticleCollection DeSerializeCollection<ArticleCollection>(string fileName) {
            if (string.IsNullOrEmpty(fileName)) { return default(ArticleCollection); }
            ArticleCollection objectOut = default(ArticleCollection);
            try {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(fileName);
                string xmlString = xmlDocument.OuterXml;
                using (StringReader read = new StringReader(xmlString)) {
                    Type outType = typeof(ArticleCollection);
                    XmlSerializer serializer = new XmlSerializer(outType);
                    using (XmlReader reader = new XmlTextReader(read)) {
                        objectOut = (ArticleCollection)serializer.Deserialize(reader);
                        reader.Close();
                    }
                    read.Close();
                }
            } catch (Exception e) {
                lblStatus.Text = e.Message;
            }
            return objectOut;
        }
        #endregion

        #region Events
        private void btnImport_Click(object sender, RoutedEventArgs e) {
            ImportXML();
        }

        private void btnAddCollection_Click(object sender, RoutedEventArgs e) {
            AddCollection();
        }

        private void btnRemoveCollection_Click(object sender, RoutedEventArgs e) {
            // TODO: Confirmation save collection?
            RemoveCollection();
        }

        private void btnExportCollection_Click(object sender, RoutedEventArgs e) {
            ExportCollections();
        }

        private void btnSaveCollection_Click(object sender, RoutedEventArgs e) {
            SaveCollections();
        }

        private void lvSource_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            Article selected = (Article)lvSource.SelectedItem;
            if (selected != null) {
                lblTitle.Content = selected.Title;
                txtArticleContent.Document.Blocks.Clear();
                txtArticleContent.Document.Blocks.Add(new Paragraph(new Run(selected.Content)));
            }
        }

        private void lvSource_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Return) {
                MoveToCollection();
            } else if (e.Key == Key.Delete) {
                DeleteSourceArticle();
            }
        }

        private void tabControl_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Return) {
                MoveToSource();
            } else if (e.Key == Key.Delete) {
                DeleteCollectionArticle();
            }
        }

        private void Main_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.S && Keyboard.IsKeyDown(Key.LeftCtrl)) {
                SaveCollections();
            }
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Enter || e.Key == Key.Return) {
                SearchText();
            }
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e) {
            SearchText();
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e) {
            TextPointer selectionStart = txtArticleContent.Document.ContentStart;
            TextPointer selectionEnd = txtArticleContent.Document.ContentEnd;
            TextRange selection = new TextRange(selectionStart, selectionEnd);
            selection.ApplyPropertyValue(TextElement.BackgroundProperty, Brushes.White);
            selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);
        }

        private void lvSourceColumnHeader_Click(object sender, RoutedEventArgs e) {
            GridViewColumnHeader column = (sender as GridViewColumnHeader);
            string sortBy = column.Tag.ToString();
            if (sortBy.Equals("Date")) {

            }
            if (listViewSortCol != null) {
                AdornerLayer.GetAdornerLayer(listViewSortCol).Remove(listViewSortAdorner);
                lvSource.Items.SortDescriptions.Clear();
            }

            ListSortDirection newDir = ListSortDirection.Ascending;
            if (listViewSortCol == column && listViewSortAdorner.Direction == newDir)
                newDir = ListSortDirection.Descending;

            listViewSortCol = column;
            listViewSortAdorner = new SortAdorner(listViewSortCol, newDir);
            AdornerLayer.GetAdornerLayer(listViewSortCol).Add(listViewSortAdorner);
            lvSource.Items.SortDescriptions.Add(new SortDescription(sortBy, newDir));
        }
        #endregion

    }
}
