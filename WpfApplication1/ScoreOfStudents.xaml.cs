using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for ScoreOfStudents.xaml
  /// </summary>
  public partial class ScoreOfStudents : UserControl {

    public ObservableCollection<StudentWithScore> MaleStudentsWithScores { get; set; } = new ObservableCollection<StudentWithScore>();
    public ObservableCollection<StudentWithScore> FemaleStudentsWithScores { get; set; } = new ObservableCollection<StudentWithScore>();

    public ICollectionView MaleStudentsView { get; }
    public ICollectionView FemaleStudentsView { get; }

    public EventHandler SaveScores;

    public ScoreOfStudents() {
      InitializeComponent();
      DataContext = this;

      MaleStudentsView = CollectionViewSource.GetDefaultView(MaleStudentsWithScores);
      FemaleStudentsView = CollectionViewSource.GetDefaultView(FemaleStudentsWithScores);
    }

    private void SaveScores_Click(object sender, RoutedEventArgs e) =>
      SaveScores?.Invoke(this, EventArgs.Empty);

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) {
      var text = (sender as TextBox)?.Text ?? string.Empty;

      if (string.IsNullOrWhiteSpace(text)) {
        MaleStudentsView.Filter = null;
        FemaleStudentsView.Filter = null;
      } else {
        MaleStudentsView.Filter = o => ((StudentWithScore)o).Name
            .Contains(text, StringComparison.OrdinalIgnoreCase);
        FemaleStudentsView.Filter = o => ((StudentWithScore)o).Name
            .Contains(text, StringComparison.OrdinalIgnoreCase);
      }

      MaleStudentsView.Refresh();
      FemaleStudentsView.Refresh();
    }

    private void MaleStudentScoreTextChanged(object sender, TextChangedEventArgs e) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection

      string newValue = tb.Text;
      string oldValue = MaleStudentsWithScores[index].Score;

      btnSaveScores.IsEnabled = (newValue != oldValue);
    }
    private void FemaleStudentScoreTextChanged(object sender, TextChangedEventArgs e) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection

      string newValue = tb.Text;
      string oldValue = FemaleStudentsWithScores[index].Score;

      btnSaveScores.IsEnabled = (newValue != oldValue);
    }

    private void LettersOnlyTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      e.Handled = !e.Text.All(c => char.IsLetter(c) || c == ',' || c == ' ' || c == '.');
    }
  }

  public class StudentWithScore {
    public string Name { get; set; }
    public string Score { get; set; }
  }
}
