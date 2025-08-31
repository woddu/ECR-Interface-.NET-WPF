
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for StudentDetails.xaml
  /// </summary>
  public partial class StudentDetails : UserControl {

    public List<string> OriginalWrittenWork { get; set; } = new List<string>();

    public List<string> OriginalPerformanceTask { get; set; } = new List<string>();

    public ObservableCollection<FieldDefinition> WrittenWorks { get; set; } = new ObservableCollection<FieldDefinition>();

    public ObservableCollection<FieldDefinition> PerformanceTasks { get; set; } = new ObservableCollection<FieldDefinition>();

    public string Exam;

    public StudentDetails() {
      InitializeComponent();
      DataContext = this;
    }

    private void WrittenScoresTextChanged(object sender, TextChangedEventArgs e) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection

      string newValue = tb.Text;
      string oldValue = OriginalWrittenWork[index];
      btnSaveWrittenWorks.IsEnabled = (newValue != oldValue);
    }

    private void PerformanceScoresTextChanged(object sender, TextChangedEventArgs e) {
      var tb = (TextBox)sender;
      int index = (int)tb.Tag; // index in the collection

      string newValue = tb.Text;
      string oldValue = OriginalPerformanceTask[index];
      btnSavePerformanceTasks.IsEnabled = (newValue != oldValue);
    }
    private void NumberOnlyTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      e.Handled = !int.TryParse(e.Text, out _);
    }

    public void SetName(string name) => lblName.Content = name;

    public void SetExam(string exam) => lblExam.Content = exam;

    public void SetGrade(string grade) {

    }

  }

}
