using System.Windows;

namespace ComputerPassport
{
    public partial class InventoryNumberDialog : Window
    {
        public string InventoryNumber { get; private set; }

        public InventoryNumberDialog()
        {
            InitializeComponent();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            InventoryNumber = txtInventoryNumber.Text?.Trim();
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            txtInventoryNumber.Text = "";
        }
    }
}