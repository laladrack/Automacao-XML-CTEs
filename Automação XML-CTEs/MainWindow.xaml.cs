using Microsoft.VisualBasic;
using Microsoft.Win32;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Automação_XML_CTEs
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

        private void dateTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            dateTextBox.Text = DateTime.Now.ToString("dd/MM/yyyy hh:mm");
        }

        private void folderSelectButton_Click(object sender, RoutedEventArgs e)
        {
            string XMLsFolder;
            var folderDialog = new OpenFolderDialog
            {
                Title = "Selecione a pasta com os XMLs.",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (folderDialog.ShowDialog() == true)
            {
                XMLsFolder = folderDialog.FolderName;
                folderSelectLabel.Content = $"Selecionada: {XMLsFolder}";
            }
        }

        private void savingFolderButton_Click(object sender, RoutedEventArgs e)
        {
            string saveFolder;
            var folderDialog = new OpenFolderDialog
            {
                Title = "Selecione a pasta para salvar o relatório.",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (folderDialog.ShowDialog() == true)
            {
                saveFolder = folderDialog.FolderName;
                folderSelectLabel.Content = $"Selecionada: {saveFolder}";
            }
        }
    }
}