
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window {

    private WorkbookService _workbookService = new WorkbookService();

    private HomePage homePage = new HomePage();
    private SamplePage studentsPage = new SamplePage();
    private HighestScores highestScoresPage = new HighestScores();
    private StudentDetails studentDetails = new StudentDetails();

    public MainWindow() {
      InitializeComponent();
      MainContent.Content = homePage;
      homePage.FileChosen += HomePage_FileChosen;
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

    private void Tab_Checked(object sender, RoutedEventArgs e) {
      if (MainContent == null) return;
      if (sender is not RadioButton rb) return;

      if (rb.Name == rbtnFile.Name) {
        MainContent.Content = homePage;
      } else if (rb.Name == rbtnStudents.Name) {
        MainContent.Content = studentsPage;
      } else if (rb.Name == rbtnScores.Name) {
        MainContent.Content = highestScoresPage;
      }

    }

    private async void HomePage_FileChosen(object sender, string[] file) {
      homePage.SetLoading(true);

      // Step 2: Run heavy work in background and return a WorkbookResult
      (bool isValid, string fileName, List<string> maleNames, List<string> femaleNames) = await Task.Run(() =>
      {
        _workbookService.LoadWorkbook(file[0]);

        if (!_workbookService.IsFileECR()) {
          return ( 
            false,
            "",
            [],
            []
          );
        }

        var names = _workbookService.ReadAllNames();

        return (
          true,
          file[1],
          names.Item1?.ToList(),
          names.Item2?.ToList()
        );
      });

      // Step 3: Back on UI thread — safe to update UI
      if (!isValid) {
        homePage.ShowError("Missing Sheets", "Missing Sheets: " + _workbookService.MissingSheets);
        homePage.SetLoading(false);
        return;
      }

      homePage.SetFileName(fileName);

      studentsPage.MaleNames.Clear();
      maleNames.ForEach(name => studentsPage.MaleNames.Add(name));

      studentsPage.FemaleNames.Clear();
      femaleNames.ForEach(name => studentsPage.FemaleNames.Add(name));

      highestScoresPage.WrittenWorks.Clear();
      _workbookService.WrittenWorks.ForEach(score => highestScoresPage.WrittenWorks.Add(score));

      highestScoresPage.PerformanceTasks.Clear();
      _workbookService.PerformanceTasks.ForEach(score => highestScoresPage.PerformanceTasks.Add(score));

      rbtnScores.IsEnabled = rbtnStudents.IsEnabled = true;
      rbtnFile.IsChecked = false;
      rbtnStudents.IsChecked = true;

      homePage.SetLoading(false);
      MainContent.Content = studentsPage;
      this.Title = fileName;
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
