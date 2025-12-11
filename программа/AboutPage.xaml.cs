using System.Windows;
using System.Windows.Controls;

namespace ComputerPassport
{
    public partial class AboutPage : Page
    {
        public AboutPage()
        {
            InitializeComponent();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            // Находим родительское окно и закрываем его
            Window.GetWindow(this)?.Close();
        }
    }
}