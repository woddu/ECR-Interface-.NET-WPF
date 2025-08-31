
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for StudentDetails.xaml
  /// </summary>
  public partial class StudentDetails : UserControl {

    public ObservableCollection<FieldDefinition> WrittenWorks { get; set; } = new ObservableCollection<FieldDefinition>();

    public ObservableCollection<FieldDefinition> PerformanceTasks { get; set; } = new ObservableCollection<FieldDefinition>();

    public string Exam;

    public StudentDetails() {
      InitializeComponent();
      DataContext = this;
    }

    public void SetName(string name) => lblName.Content = name;

    public void SetExam(string exam) => lblExam.Content = exam;

    public void SetGrade(string grade) {

    }

  }

}
