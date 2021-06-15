using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;

namespace RenderTableCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string selectedFile;
        private string renderTableFile;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.Filter = "RenPy files (*.rpy)|*.rpy|All files (*.*)|*.*";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == true)
            {
                selectedFile = openFileDialog.FileName;
                ChosenFile.Text = $"Selected File: {selectedFile}";
                CreateRenderTableButton.Visibility = Visibility.Visible;
            }
            renderTableFile = Path.ChangeExtension(selectedFile.Trim(), ".docx");
        }

        private void CreateRenderTableButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
