using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Input;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for HighestScores.xaml
  /// </summary>
  public partial class HighestScores : UserControl {

    public ObservableCollection<string> WrittenWorks { get; set; } = new ObservableCollection<string>();

    public ObservableCollection<string> PerformanceTasks { get; set; } = new ObservableCollection<string>();

    private string Exam { get; set; }

    public HighestScores() {
      InitializeComponent();
      DataContext = this;
    }

    public void SetWrittenWorkSaveEnabled(bool isEnabled) {
      btnSaveWrittenWorks.IsEnabled = isEnabled;
    }
    public void SetPerformanceTaskSaveEnabled(bool isEnabled) {
      btnSavePerformanceTasks.IsEnabled = isEnabled;
    }

    private void WrittenScoresTextChanged(object sender, TextChangedEventArgs e) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection

      string newValue = tb.Text;
      string oldValue = WrittenWorks[index];

      btnSaveWrittenWorks.IsEnabled = (newValue != oldValue);
    }

    private void PerformanceScoresTextChanged(object sender, TextChangedEventArgs e) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection
      
      string newValue = tb.Text;
      string oldValue = PerformanceTasks[index];

      btnSavePerformanceTasks.IsEnabled = (newValue != oldValue);
    }

    private void ExamTextChanged(object sender, TextChangedEventArgs e) {
      btnSaveExam.IsEnabled = !(txtExam.Text != Exam);
    }

    private void NumberOnlyTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      e.Handled = !int.TryParse(e.Text, out _);
    }

    public void SetExam(string exam) {
      Exam = exam;
      txtExam.Text = exam;
    }
  }
}
