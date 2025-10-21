using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for ScoreOfStudents.xaml
  /// </summary>
  public partial class ScoreOfStudents : UserControl {

    public List<string> InitialMaleStudentsScores { get; set; } = new List<string>();

    public List<string> InitialFemaleStudentsScores { get; set; } = new List<string>();

    public ObservableCollection<StudentWithScore> MaleStudentsWithScores { get; set; } = new ObservableCollection<StudentWithScore>();
    public ObservableCollection<StudentWithScore> FemaleStudentsWithScores { get; set; } = new ObservableCollection<StudentWithScore>();

    public ICollectionView MaleStudentsView { get; }
    public ICollectionView FemaleStudentsView { get; }

    public EventHandler SaveScores;

    private ScoreType _type;
    public ScoreType Type {
      get { return _type; }
      set {
        _type = value;
        tbType.Text = _type switch {
          ScoreType.WrittenWorks => "Written Works",
          ScoreType.PerformanceTasks => "Performance Tasks",
          ScoreType.Exam => "Exam",
          _ => "Unknown"
        };
      }
    }

    private int _highestScore;
    public int HighestScore {
      get { return _highestScore; }
      set { tbHighestScore.Text = "Highest Score: " + (_highestScore = value).ToString(); }
    }

    private uint _columnIndex;
    public uint ColumnIndex {
      get { return _columnIndex; }
      set {
        _columnIndex = value;
        tbType.Text = tbType.Text + " #" + (_columnIndex + 1).ToString(); }
    }

    private bool _suppressLostFocus;
    public ScoreOfStudents() {
      InitializeComponent();
      DataContext = this;

      MaleStudentsView = CollectionViewSource.GetDefaultView(MaleStudentsWithScores);
      FemaleStudentsView = CollectionViewSource.GetDefaultView(FemaleStudentsWithScores);
      btnClear.Visibility = Visibility.Collapsed;
    }

    private void MaleStudentScore_LostFocus(object sender, RoutedEventArgs e) {
      if (_suppressLostFocus) {
        _suppressLostFocus = false;
        return;
      }
      SaveFromTextBox(sender, true);
    }

    private void MaleStudentScore_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Enter) {
        _suppressLostFocus = true;
        SaveFromTextBox(sender, true);
        e.Handled = true;
      }
    }

    private void FemaleStudentScore_LostFocus(object sender, RoutedEventArgs e) {
      if (_suppressLostFocus) {
        _suppressLostFocus = false;
        return;
      }
      SaveFromTextBox(sender, false);
    }

    private void FemaleStudentScore_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Enter) {
        _suppressLostFocus = true;
        SaveFromTextBox(sender, false);
        e.Handled = true;
      }
    }

    private void SaveScores_Click(object sender, RoutedEventArgs e) =>
      SaveScores?.Invoke(this, EventArgs.Empty);

    private void ClearSearch_Click(object sender, RoutedEventArgs e) {
      searchBox.Text = string.Empty;
      searchBox.Focus();
      btnClear.Visibility = Visibility.Collapsed;
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) {
      var text = (sender as TextBox)?.Text ?? string.Empty;

      if (string.IsNullOrWhiteSpace(text)) {
        MaleStudentsView.Filter = null;
        FemaleStudentsView.Filter = null;
        btnClear.Visibility = Visibility.Collapsed;
      } else {
        MaleStudentsView.Filter = o => ((StudentWithScore)o).Name
            .Contains(text, StringComparison.OrdinalIgnoreCase);
        FemaleStudentsView.Filter = o => ((StudentWithScore)o).Name
            .Contains(text, StringComparison.OrdinalIgnoreCase);
        btnClear.Visibility = Visibility.Visible;
      }

      MaleStudentsView.Refresh();
      FemaleStudentsView.Refresh();
    }

    private void MaleStudentScoreTextChanged(object sender, TextChangedEventArgs e) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection

      string newValue = tb.Text;
      string oldValue = InitialMaleStudentsScores[index];
      btnSaveScores.IsEnabled = (newValue != oldValue);
    }
    private void FemaleStudentScoreTextChanged(object sender, TextChangedEventArgs e) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection

      string newValue = tb.Text;
      string oldValue = InitialFemaleStudentsScores[index];
      btnSaveScores.IsEnabled = (newValue != oldValue);
    }

    private void SaveFromTextBox(object sender, bool isMale = true) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection

      string newValue = tb.Text;
      string oldValue;
      if (isMale) {
        oldValue = InitialMaleStudentsScores[index];
      } else {
        oldValue = InitialFemaleStudentsScores[index];
      }
      if (!string.IsNullOrWhiteSpace(newValue) && (newValue != oldValue)) {
        SaveScores?.Invoke(this, EventArgs.Empty);
        searchBox.Focus();
      }
    }
    public void ResetSaveButton() {
      btnSaveScores.IsEnabled = false;
    }
    private void LettersOnlyTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      e.Handled = !e.Text.All(c => char.IsLetter(c) || c == ',' || c == ' ' || c == '.');
    }

    private void NumberOnlyTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      var tb = (TextBox)sender;

      // Predict what the text will look like after this input
      string newText = tb.Text.Insert(tb.SelectionStart, e.Text);

      if (int.TryParse(newText, out int value)) {
        if (value > HighestScore) {
          e.Handled = true; // block the input
        }
      } else {
        e.Handled = true; // block non-numeric input
      }
    }
  }

  public class StudentWithScore {
    public string Name { get; set; }
    public string Score { get; set; }
  }

}
