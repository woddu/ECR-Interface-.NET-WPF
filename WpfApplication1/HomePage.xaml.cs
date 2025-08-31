using MaterialDesignThemes.Wpf;
using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for HomePage.xaml
    /// </summary>
    public partial class HomePage : UserControl
    {

        public event EventHandler<string[]> FileChosen;

        public HomePage()
        {
          InitializeComponent();
    }

        private void ChooseFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select an Excel file",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
            };

            if (dialog.ShowDialog() == true) // true means user clicked OK
            {

                FileChosen?.Invoke(this, new string[]
                {
                    dialog.FileName,
                    Path.GetFileNameWithoutExtension(dialog.FileName)
                });
            }
        }

        public void ShowError(string title, string error)
        {
            MessageBox.Show(
                "Error: " + error,          // Title bar text
                title,    // Message text
                MessageBoxButton.OK,        // Buttons to display
                MessageBoxImage.Error       // Icon type
            );
        }

        public void SetFileName(string fileName)
        {
            FileName.Content = fileName;
            btnChooseFile.Content = "Choose a Different File";
        }

    public void SetLoading(bool isLoading) {
      ButtonProgressAssist.SetIsIndeterminate(btnChooseFile, isLoading);
    }
    
  }
}
