using MaterialDesignThemes.Wpf;
using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for HomePage.xaml
  /// </summary>
  public partial class HomePage : UserControl {

    public event EventHandler<string[]> FileChosen;
    public event EventHandler<bool> QuarterChosen;
    public event EventHandler<int> TrackChosen;

    public HomePage() {
      InitializeComponent();
      spQuarter.Visibility = Visibility.Collapsed;
      cBoxTracks.Visibility = Visibility.Visible;
    }

    private void Track_SelectionChanged(object sender, SelectionChangedEventArgs e) {
      if (cBoxTracks.SelectedItem is ComboBoxItem selectedItem) {
        int index = cBoxTracks.SelectedIndex;
        TrackChosen?.Invoke(this, index);
      }
    }

    public void ChangeQuarter(bool q) {
      if (q) {
        rbtn1st.IsChecked = true;
        rbtn2nd.IsChecked = false;
      } else {
        rbtn2nd.IsChecked = true;
        rbtn1st.IsChecked = false;
      }
    }

    private void Change_Quarter(object sender, RoutedEventArgs e) {
      if (sender is RadioButton rbutton) {
        bool tagValue = bool.TryParse(rbutton.Tag?.ToString(), out var result) && result;
        QuarterChosen?.Invoke(this, tagValue);
      }
    }
    private void ChooseFile_Click(object sender, RoutedEventArgs e) {
      var dialog = new OpenFileDialog {
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

    public void ShowError(string title, string error) {
      MessageBox.Show(
          "Error: " + error,          // Title bar text
          title,    // Message text
          MessageBoxButton.OK,        // Buttons to display
          MessageBoxImage.Error       // Icon type
      );
    }

    public void SetFileName(string fileName) {
      FileName.Text = fileName;
      btnChooseFile.Content = "Choose a Different File";
      btnChooseFile.IsEnabled = true;
      spQuarter.Visibility = Visibility.Visible;
      cBoxTracks.Visibility = Visibility.Visible;
    }

    public void SetLoading(bool isLoading) {
      ButtonProgressAssist.SetIsIndeterminate(btnChooseFile, isLoading);
    }

  }
}
