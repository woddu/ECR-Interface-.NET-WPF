
using System;
using System.Diagnostics;
using System.Windows;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window {
    private WorkbookService _workbookService = new WorkbookService();

    private HomePage homePage = new HomePage();
    private SamplePage studentsPage = new SamplePage();
    private HighestScores highestScores = new HighestScores();
    private StudentDetails studentDetails = new StudentDetails();

    public MainWindow() {
      InitializeComponent();
      MainContent.Content = homePage;
      homePage.FileChosen += HomePage_FileChosen;
      btnHighestScores.IsEnabled = btnStudents.IsEnabled = false;
      studentsPage.AddMale += StudentsPage_AddMale;
      studentsPage.AddFemale += StudentsPage_AddFemale;
      studentsPage.NameClicked += StudentsPage_NameClicked;
      
    }

    private void StudentsPage_NameClicked(object sender, Tuple<uint, string> tuple) {
      studentDetails.WrittenWorks.Clear();
      studentDetails.PerformanceTasks.Clear();
      var scores = _workbookService.ReadStudentScores(tuple.Item1);
      studentDetails.SetName(tuple.Item2);
      for (int i = 0; i < _workbookService.WrittenWorks.Count; i++) {
        studentDetails.WrittenWorks.Add(new FieldDefinition {
          Label = _workbookService.WrittenWorks[i],
          Value = scores.Item1[i]
        });
      }

      for (int i = 0; i < _workbookService.PerformanceTasks.Count; i++) {
        studentDetails.PerformanceTasks.Add(new FieldDefinition {
          Label = _workbookService.PerformanceTasks[i],
          Value = scores.Item2[i]
        });
      }

      MainContent.Content = studentDetails;
    }

    private void Home_Click(object sender, RoutedEventArgs e) {
      //btnHome.Background = (System.Windows.Media.SolidColorBrush)new System.Windows.Media.BrushConverter()
      //                     .ConvertFromString("#005BB5");
      //btnSample.Background = (System.Windows.Media.SolidColorBrush)new System.Windows.Media.BrushConverter()
      //                     .ConvertFromString("#FF004C9A");
      //btnHome.Padding = new Thickness(15, 10, 20, 10);
      //btnHome.BorderThickness = new Thickness(5, 0, 0, 0);
      MainContent.Content = homePage;
    }

    
    private void HighestScores_Click(object sender, RoutedEventArgs e) {
      MainContent.Content = highestScores;
    }

    private void Students_Click(object sender, RoutedEventArgs e) {
      //btnHome.Background = (System.Windows.Media.SolidColorBrush)new System.Windows.Media.BrushConverter()
      //                     .ConvertFromString("#FF004C9A");
      //btnSample.Background = (System.Windows.Media.SolidColorBrush)new System.Windows.Media.BrushConverter()
      //                     .ConvertFromString("#005BB5");
      //btnSample.Padding = new Thickness(15, 10, 20, 10);
      //btnSample.BorderThickness = new Thickness(5, 0, 0, 0);
      MainContent.Content = studentsPage;
    }

    private void HomePage_FileChosen(object sender, string[] file) {
      _workbookService.LoadWorkbook(file[0]);
      if (!_workbookService.IsFileECR()) {
        homePage.ShowError("Missing Sheets", "Missing Sheets: " + _workbookService.GetMissingSheets());
        return;
      }
      homePage.SetFileName(file[1]);
      var names = _workbookService.ReadAllNames();
      names.Item1.ForEach(name => studentsPage.MaleNames.Add(name));
      names.Item2.ForEach(name => studentsPage.FemaleNames.Add(name));
      Application.Current.Dispatcher.BeginInvoke(new Action(() => {
        _workbookService.ReadHighestPossibleScores();
        _workbookService.WrittenWorks.ForEach(score => highestScores.WrittenWorks.Add(score));
        _workbookService.PerformanceTasks.ForEach(score => highestScores.PerformanceTasks.Add(score));

      }));
      MainContent.Content = studentsPage;
      this.Title = file[1];
      btnStudents.IsEnabled = btnHighestScores.IsEnabled = true;
    }

    private void StudentsPage_AddMale(object sender, string newName) {
      studentsPage.EnableAddButton(false);
      studentsPage.MaleNames.Clear();
      _workbookService.AppendAndSortNames(newName).ForEach(name => studentsPage.MaleNames.Add(name));
      studentsPage.EnableAddButton();
    }

    private void StudentsPage_AddFemale(object sender, string newName) {
      studentsPage.EnableAddButton(false);
      studentsPage.FemaleNames.Clear();
      _workbookService.AppendAndSortNames(newName, false).ForEach(name => studentsPage.FemaleNames.Add(name));
      studentsPage.EnableAddButton();
    }

  }
}
