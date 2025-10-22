using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for SamplePage.xaml
  /// </summary>
  public partial class SamplePage : UserControl {
    public event EventHandler<string> AddMale;
    public event EventHandler<string> AddFemale;
    public event Action<uint, string, bool> NameClicked;

    private System.ComponentModel.ICollectionView _collectionViewMale;
    private System.ComponentModel.ICollectionView _collectionViewFemale;
    private bool genderIsMale = true;

    public ObservableCollection<string> MaleNames { get; set; } = new ObservableCollection<string>();
    public ObservableCollection<string> FemaleNames { get; set; } = new ObservableCollection<string>();

    public SamplePage() {
      InitializeComponent();
      //DataContext = this;

      _collectionViewMale = System.Windows.Data.CollectionViewSource.GetDefaultView(MaleNames);
      _collectionViewFemale = System.Windows.Data.CollectionViewSource.GetDefaultView(FemaleNames);
      listMale.ItemsSource = _collectionViewMale;
      listFemale.ItemsSource = _collectionViewFemale;
    }

    private void LettersOnlyTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      e.Handled = !e.Text.All(c => char.IsLetter(c) || c == ',' || c == ' ' || c == '.');
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) {
      var text = (sender as TextBox)?.Text ?? string.Empty;

      _collectionViewMale.Filter = item => {
        if (string.IsNullOrWhiteSpace(text)) return true;
        return (item as string)?.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;
      };
      _collectionViewFemale.Filter = item => {
        if (string.IsNullOrWhiteSpace(text)) return true;
        return (item as string)?.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;
      };

      _collectionViewMale.Refresh();
      _collectionViewFemale.Refresh();
    }

    private void Gender_Checked(object sender, RoutedEventArgs e) {
      if (sender is not RadioButton rb) return;
      if (rb.Name == rbMale.Name) {
        genderIsMale = true;
      } else if (rb.Name == rbFemale.Name) {
        genderIsMale = false;
      }
    }

    private void AddName_Click(object sender, RoutedEventArgs e) {
      if(string.IsNullOrWhiteSpace(txtName.Text)) return;

      if (genderIsMale) {
        AddMale?.Invoke(this, txtName.Text.ToUpper());
      } else {
        AddFemale?.Invoke(this, txtName.Text.ToUpper());
      }
    }

    private void listMale_SelectionChanged(object sender, SelectionChangedEventArgs e) {
      if (listMale.SelectedItem != null) {
        int index = MaleNames.IndexOf(listMale.SelectedItem.ToString());
        NameClicked?.Invoke((uint)index, listMale.SelectedItem.ToString(), true);
        listMale.SelectedIndex = -1;
      }
    }

    private void listFemale_SelectionChanged(object sender, SelectionChangedEventArgs e) {
      if (listFemale.SelectedItem != null) {
        int index = FemaleNames.IndexOf(listFemale.SelectedItem.ToString());
        NameClicked?.Invoke((uint)index, listFemale.SelectedItem.ToString(), false);
        listFemale.SelectedIndex = -1;
      }
    }

    public void EnableAddButton(bool enable = true) => btnAdd.IsEnabled = enable;
  }
}
