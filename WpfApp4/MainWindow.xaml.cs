using Microsoft.Win32;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp4
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "СSV file (*.csv)|*.csv";
            if (openFileDialog.ShowDialog() == true)
                txtEditor.Text = openFileDialog.FileName;
            
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (txtEditor.Text != null)
            {
                ConverterLogic converter = new ConverterLogic(txtEditor.Text);
                converter.Convert();
            }
            else if (txtEditor.Text == null)
                MessageBox.Show("Вы не выбрали файл!");
        }
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("cmd", $"/c start https://contacts.google.com/"));
        }

       
    }
}
