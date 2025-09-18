
using DocumentFormat.OpenXml.Spreadsheet;
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
      homePage.QuarterChosen += Homepage_QuarterChosen;
      homePage.TrackChosen += HomePage_Track_Chosen;

      studentsPage.AddMale += StudentsPage_AddMale;
      studentsPage.AddFemale += StudentsPage_AddFemale;
      studentsPage.NameClicked += StudentsPage_NameClicked;

      highestScoresPage.SaveExamClicked += HighestScores_SaveExamClicked;
      highestScoresPage.SaveWrittenWorksClicked += HighestScores_SaveWrittenWorksClicked;
      highestScoresPage.SavePerformanceTasksClicked += HighestScores_SavePerformanceTasksClicked;

      studentDetails.SaveExamClicked += StudentDetails_SaveExamClicked;
      studentDetails.SaveWrittenWorksClicked += StudentDetails_SaveWrittenWorksClicked;
      studentDetails.SavePerformanceTasksClicked += StudentDetails_SavePerformanceTasksClicked;

    }

    private async void HomePage_Track_Chosen(object sender, int e) {
      if (WorkbookService.tracks[e] != _workbookService.Track) {
        homePage.SetLoading(true);
        _workbookService.Track = WorkbookService.tracks[e];
        homePage.SetLoading(false);
      }
    }

    private async void Homepage_QuarterChosen(object sender, bool e) {      
      if (e != _workbookService.Quarter1) {
        homePage.SetLoading(true);
        await Task.Run(() => {
          _workbookService.ChangeQuarter();
        });

        homePage.ChangeQuarter(_workbookService.Quarter1);
        homePage.SetLoading(false);
      }
    }

    private async void StudentDetails_SaveExamClicked(object sender, EventArgs e) {
      (string exam, string grade, int transmutedGrade) = await Task.Run(() => {
        _workbookService.EditStudentExam(
          studentDetails.Exam,
          studentDetails.StudentRow
          );
        var (writtenWorks, performanceTasks, exam, grade) =
          _workbookService.ReadStudentScores(studentDetails.StudentRow);
        var transmuted = _workbookService.GetComputedGrade(writtenWorks, performanceTasks, exam);
        return (exam, grade, transmuted);
      });
      studentDetails.Exam = exam;

      studentDetails.SetGrade(transmutedGrade.ToString());
      studentDetails.SetSaveExamBtnEnabled(false);
    }

    private async void StudentDetails_SaveWrittenWorksClicked(object sender, EventArgs e) {
      (List<string> writtenWorksScores, string grade, int transmutedGrade) = await Task.Run(() => {
        _workbookService.EditStudentScore(
          studentDetails.WrittenWorks.Select(p => p.Value).ToList(),
          studentDetails.StudentRow
          );

        var (writtenWorks, performanceTasks, exam, grade) =
          _workbookService.ReadStudentScores(studentDetails.StudentRow);
        var transmuted = _workbookService.GetComputedGrade(writtenWorks, performanceTasks, exam);
        return (writtenWorks, grade, transmuted);
      });
      studentDetails.WrittenWorks.Clear();
      studentDetails.OriginalWrittenWork.Clear();
      for (int i = 0; i < _workbookService.WrittenWorks.Count; i++) {
        studentDetails.WrittenWorks.Add(new FieldDefinition {
          Label = _workbookService.WrittenWorks[i],
          Value = writtenWorksScores[i]
        });
        studentDetails.OriginalWrittenWork.Add(writtenWorksScores[i]);
      }

      studentDetails.SetGrade(transmutedGrade.ToString());
      studentDetails.SetSaveWrittenWorksBtnEnabled(false);
    }

    private async void StudentDetails_SavePerformanceTasksClicked(object sender, EventArgs e) {
      (List<string> performanceTasksScores, string grade, int transmutedGrade) = await Task.Run(() => {
        _workbookService.EditStudentScore(
          studentDetails.PerformanceTasks.Select(p => p.Value).ToList(),
          studentDetails.StudentRow,
          false
          );

        var (writtenWorks, performanceTasks, exam, grade) =
          _workbookService.ReadStudentScores(studentDetails.StudentRow);
        var transmuted = _workbookService.GetComputedGrade(writtenWorks, performanceTasks, exam);
        return (performanceTasks, grade, transmuted);
      });
      studentDetails.PerformanceTasks.Clear();
      studentDetails.OriginalPerformanceTask.Clear();
      for (int i = 0; i < _workbookService.PerformanceTasks.Count; i++) {
        studentDetails.PerformanceTasks.Add(new FieldDefinition {
          Label = _workbookService.PerformanceTasks[i],
          Value = performanceTasksScores[i]
        });
        studentDetails.OriginalPerformanceTask.Add(performanceTasksScores[i]);
      }

      studentDetails.SetGrade(transmutedGrade.ToString());
      studentDetails.SetSavePerformanceTasksBtnEnabled(false);
    }

    private async void HighestScores_SaveExamClicked(object sender, EventArgs e) {
      await Task.Run(() => {
        _workbookService.EditExamScore(highestScoresPage.Exam);
      });
      highestScoresPage.Exam = _workbookService.Exam;
      highestScoresPage.SetExamBtnEnabled(false);
    }

    private async void HighestScores_SaveWrittenWorksClicked(object sender, EventArgs e) {

      await Task.Run(() => {
        _workbookService.EditHighestPossibleScore(highestScoresPage.WrittenWorks.ToList(), true);
      });

      highestScoresPage.WrittenWorks.Clear();
      _workbookService.WrittenWorks.ForEach(score => highestScoresPage.WrittenWorks.Add(score));

    }

    private async void HighestScores_SavePerformanceTasksClicked(object sender, EventArgs e) {

      await Task.Run(() => {
        _workbookService.EditHighestPossibleScore(highestScoresPage.PerformanceTasks.ToList(), false);
      });

      highestScoresPage.PerformanceTasks.Clear();
      _workbookService.PerformanceTasks.ForEach(score => highestScoresPage.PerformanceTasks.Add(score));

    }

    private async void StudentsPage_NameClicked(object sender, (uint row, string name) tuple) {
      studentDetails.OriginalWrittenWork.Clear();
      studentDetails.OriginalPerformanceTask.Clear();
      studentDetails.WrittenWorks.Clear();
      studentDetails.PerformanceTasks.Clear();
      (List<string> writtenWorks, List<string> performanceTasks, string exam, string grade, int transmutedGrade) = await Task.Run(() => {
        var (writtenWorks, performanceTasks, exam, grade) = _workbookService.ReadStudentScores(tuple.row);
        var transmuted = _workbookService.GetComputedGrade(writtenWorks, performanceTasks, exam);
        return (writtenWorks, performanceTasks, exam, grade, transmuted);
      });

      studentDetails.StudentRow = tuple.row;
      studentDetails.StudentName = tuple.name;
      studentDetails.Exam = exam;
      studentDetails.SetGrade(transmutedGrade.ToString());
      for (int i = 0; i < _workbookService.WrittenWorks.Count; i++) {

        studentDetails.WrittenWorks.Add(new FieldDefinition {
          Label = _workbookService.WrittenWorks[i],
          Value = writtenWorks[i]
        });
        studentDetails.OriginalWrittenWork.Add(writtenWorks[i]);

      }

      for (int i = 0; i < _workbookService.PerformanceTasks.Count; i++) {

        studentDetails.PerformanceTasks.Add(new FieldDefinition {
          Label = _workbookService.PerformanceTasks[i],
          Value = performanceTasks[i]
        });
        studentDetails.OriginalPerformanceTask.Add(performanceTasks[i]);

      }
      rbtnStudents.IsChecked = false;
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

    private void StudentsPage_CLiked(object sender, EventArgs e) {
      MainContent.Content = studentsPage;
      rbtnStudents.IsChecked = true;
    }

    private async void HomePage_FileChosen(object sender, string[] file) {
      homePage.SetLoading(true);
      
      (bool isValid, string fileName, List<string> maleNames, List<string> femaleNames) = await Task.Run(() => {
        _workbookService.LoadWorkbook(file[0]);

        if (!_workbookService.IsFileECR()) {
          return (
            false,
            "",
            [],
            []
          );
        }

        var (maleNames, femaleNames) = _workbookService.ReadAllNames();

        return (
          true,
          file[1],
          maleNames?.ToList(),
          femaleNames?.ToList()
        );
      });
      
      if (!isValid) {
        homePage.ShowError("Missing Sheets", "Missing Sheets: " + _workbookService.MissingSheets);
        homePage.SetLoading(false);
        return;
      }

      homePage.SetFileName(fileName);      
      this.Title = fileName;

      studentsPage.MaleNames.Clear();
      maleNames.ForEach(name => studentsPage.MaleNames.Add(name));

      studentsPage.FemaleNames.Clear();
      femaleNames.ForEach(name => studentsPage.FemaleNames.Add(name));

      highestScoresPage.WrittenWorks.Clear();
      _workbookService.WrittenWorks.ForEach(score => highestScoresPage.WrittenWorks.Add(score));

      highestScoresPage.PerformanceTasks.Clear();
      _workbookService.PerformanceTasks.ForEach(score => highestScoresPage.PerformanceTasks.Add(score));

      highestScoresPage.Exam = _workbookService.Exam;

      rbtnScores.IsEnabled = rbtnStudents.IsEnabled = true;
      rbtnFile.IsChecked = false;
      rbtnStudents.IsChecked = true;

      homePage.SetLoading(false);
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
