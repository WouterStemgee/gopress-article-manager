using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace ArticlesUI {
    /// <summary>
    /// Interaction logic for InputDialog.xaml
    /// </summary>
    public partial class InputDialog : Window {
        public InputDialog(string question, string defaultAnswer = "Input") {
            InitializeComponent();
            lblQuestion.Content = question;
            txtAnswer.Text = defaultAnswer;
            this.Title = defaultAnswer;
            txtAnswer.Focus();
        }

        private void btnDialogOk_Click(object sender, RoutedEventArgs e) {
            if (txtAnswer.Text.Equals("")) {
                txtAnswer.Focus();
            } else {
                this.DialogResult = true;
            }
        }

        private void Window_ContentRendered(object sender, EventArgs e) {
            txtAnswer.SelectAll();
            txtAnswer.Focus();
        }

        public string Answer {
            get { return txtAnswer.Text; }
        }
    }
}
