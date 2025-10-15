
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window {

    private readonly WorkbookService _workbookService = new WorkbookService();

    private readonly HomePage homePage = new HomePage();
    private readonly SamplePage studentsPage = new SamplePage();
    private readonly HighestScores highestScoresPage = new HighestScores();
    private readonly StudentDetails studentDetails = new StudentDetails();
    private readonly ScoreOfStudents scoreOfStudents = new ScoreOfStudents();

    public MainWindow() {
      InitializeComponent();
      MyProgressBar.Visibility = Visibility.Collapsed;
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
      highestScoresPage.HighestScoreItemClicked += HighestScores_HighestScoreItemClicked;

      studentDetails.SaveExamClicked += StudentDetails_SaveExamClicked;
      studentDetails.SaveWrittenWorksClicked += StudentDetails_SaveWrittenWorksClicked;
      studentDetails.SavePerformanceTasksClicked += StudentDetails_SavePerformanceTasksClicked;

      scoreOfStudents.SaveScores += ScoreOfStudents_SaveScoresClicked;
    }

    private async void ScoreOfStudents_SaveScoresClicked(object sender, EventArgs e) {
      try {
        SetLoading(true);
        (List<string> maleStudents, List<string> femaleStudents) = await Task.Run(() => {
          _workbookService.SetScoresOfItem(            
            scoreOfStudents.ColumnIndex,
            scoreOfStudents.MaleStudentsWithScores.Select(s => s.Score).ToList(),
            scoreOfStudents.FemaleStudentsWithScores.Select(s => s.Score).ToList(),
            scoreOfStudents.Type == "Written Works"
            );
          (List<string> maleStudentsWS, List<string> femaleStudentsWS) = _workbookService.GetScoresOfItem(scoreOfStudents.ColumnIndex, scoreOfStudents.Type == "Written Works");
          return (maleStudentsWS, femaleStudentsWS);
        });

        List<StudentWithScore> maleStudentWithScores = scoreOfStudents.MaleStudentsWithScores.ToList();
        List<StudentWithScore> femaleStudentWithScores = scoreOfStudents.FemaleStudentsWithScores.ToList();

        for (int i = 0; i < maleStudents.Count; i++) {
          maleStudentWithScores[i].Score = maleStudents[i];
        }

        for (int i = 0; i < femaleStudents.Count; i++) {
          femaleStudentWithScores[i].Score = femaleStudents[i];
        }

        scoreOfStudents.MaleStudentsWithScores.Clear();
        scoreOfStudents.FemaleStudentsWithScores.Clear();

        if (scoreOfStudents.Type == "Written Works")
          scoreOfStudents.Type = "Written Works";
        else
          scoreOfStudents.Type = "Performance Tasks";

        scoreOfStudents.HighestScore = scoreOfStudents.HighestScore;

        scoreOfStudents.ColumnIndex = scoreOfStudents.ColumnIndex;

        maleStudentWithScores.ForEach(s => {
          scoreOfStudents.MaleStudentsWithScores.Add(s);
          scoreOfStudents.InitialMaleStudentsScores.Add(s.Score);
        });
        femaleStudentWithScores.ForEach(s => {
          scoreOfStudents.FemaleStudentsWithScores.Add(s);
          scoreOfStudents.InitialFemaleStudentsScores.Add(s.Score);
        });

        MainContent.Content = scoreOfStudents;

        rbtnScores.IsChecked = false;

        SetLoading(false);

        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void HighestScores_HighestScoreItemClicked(uint index, bool isWrittenWork, int highestScore) {
      try {
        SetLoading(true);
        (List<string> maleStudents, List<string> femaleStudents) = await Task.Run(() => {
          (List<string> maleStudentsWS, List<string> femaleStudentsWS) = _workbookService.GetScoresOfItem(index, isWrittenWork);
          return (maleStudentsWS, femaleStudentsWS);
        });

        List<StudentWithScore> maleStudentWithScores = scoreOfStudents.MaleStudentsWithScores.ToList();
        List<StudentWithScore> femaleStudentWithScores = scoreOfStudents.FemaleStudentsWithScores.ToList();

        for (int i = 0; i < maleStudents.Count; i++) {
          maleStudentWithScores[i].Score = maleStudents[i];
        }

        for (int i = 0; i < femaleStudents.Count; i++) {
          femaleStudentWithScores[i].Score = femaleStudents[i];
        }

        scoreOfStudents.MaleStudentsWithScores.Clear();
        scoreOfStudents.FemaleStudentsWithScores.Clear();

        if (isWrittenWork)
          scoreOfStudents.Type = "Written Works";
        else
          scoreOfStudents.Type = "Performance Tasks";

        scoreOfStudents.HighestScore = highestScore;

        scoreOfStudents.ColumnIndex = index;

        maleStudentWithScores.ForEach(s => {
          scoreOfStudents.MaleStudentsWithScores.Add(s);
          scoreOfStudents.InitialMaleStudentsScores.Add(s.Score);
        });
        femaleStudentWithScores.ForEach(s => {
          scoreOfStudents.FemaleStudentsWithScores.Add(s);
          scoreOfStudents.InitialFemaleStudentsScores.Add(s.Score);
        });

        MainContent.Content = scoreOfStudents;

        rbtnScores.IsChecked = false;

        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void HomePage_Track_Chosen(object sender, int index) {
      try {
        if (WorkbookService.tracks[index] != _workbookService.Track) {
          SetLoading(true);
          await Task.Run(() => {
            Debug.WriteLine("Track: " + WorkbookService.tracks[index]);
            _workbookService.ChangeTrack(index);
          });
          highestScoresPage.SetWrittenWorksPercentage((_workbookService.WeightedScores[0] * 100) + "%");

          highestScoresPage.SetPerformancePercentage((_workbookService.WeightedScores[1] * 100) + "%");

          highestScoresPage.SetExamPercentage((_workbookService.WeightedScores[2] * 100) + "%");
          SetLoading(false);
        }
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void Homepage_QuarterChosen(object sender, bool e) {
      if (e != _workbookService.Quarter1) {
        SetLoading(true);
        try {
          await Task.Run(() => {
            _workbookService.ChangeQuarter();
          });

          highestScoresPage.Exam = _workbookService.Exam;

          highestScoresPage.WrittenWorks.Clear();

          _workbookService.WrittenWorks.ForEach(score => highestScoresPage.WrittenWorks.Add(score));

          highestScoresPage.PerformanceTasks.Clear();

          _workbookService.PerformanceTasks.ForEach(score => highestScoresPage.PerformanceTasks.Add(score));



          homePage.ChangeQuarter(_workbookService.Quarter1);
        } catch (Exception ex) {
          ShowError("Error", ex.Message);
        }
        SetLoading(false);
      }
    }

    private async void StudentDetails_SaveExamClicked(object sender, EventArgs e) {
      try {
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
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void StudentDetails_SaveWrittenWorksClicked(object sender, EventArgs e) {
      try {
        SetLoading(true);
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
        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void StudentDetails_SavePerformanceTasksClicked(object sender, EventArgs e) {
      try {
        SetLoading(true);
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
        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void HighestScores_SaveExamClicked(object sender, EventArgs e) {
      try {
        SetLoading(true);
        await Task.Run(() => {
          _workbookService.EditExamScore(highestScoresPage.Exam);
        });
        highestScoresPage.Exam = _workbookService.Exam;
        highestScoresPage.SetExamBtnEnabled(false);
        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void HighestScores_SaveWrittenWorksClicked(object sender, EventArgs e) {
      try {
        SetLoading(true);
        await Task.Run(() => {
          _workbookService.EditHighestPossibleScore(highestScoresPage.WrittenWorks.ToList(), true);
        });

        highestScoresPage.WrittenWorks.Clear();
        _workbookService.WrittenWorks.ForEach(score => highestScoresPage.WrittenWorks.Add(score));
        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void HighestScores_SavePerformanceTasksClicked(object sender, EventArgs e) {
      try {
        SetLoading(true);
        await Task.Run(() => {
          _workbookService.EditHighestPossibleScore(highestScoresPage.PerformanceTasks.ToList(), false);
        });

        highestScoresPage.PerformanceTasks.Clear();
        _workbookService.PerformanceTasks.ForEach(score => highestScoresPage.PerformanceTasks.Add(score));
        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void StudentsPage_NameClicked(object sender, (uint row, string name) tuple) {
      try {
        SetLoading(true);
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
        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
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
      SetLoading(true);
      try {
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
          ShowError("Missing Sheets", "Missing Sheets: " + _workbookService.MissingSheets);
          homePage.SetLoading(false);
          SetLoading(false);
          return;
        }

        homePage.SetFileName(fileName);
        this.Title = fileName;

        studentsPage.MaleNames.Clear();
        maleNames.ForEach(name => studentsPage.MaleNames.Add(name));

        studentsPage.FemaleNames.Clear();
        femaleNames.ForEach(name => studentsPage.FemaleNames.Add(name));

        scoreOfStudents.MaleStudentsWithScores.Clear();
        maleNames.ForEach(name => scoreOfStudents.MaleStudentsWithScores.Add(new StudentWithScore { Name = name, Score = "0" }));

        scoreOfStudents.FemaleStudentsWithScores.Clear();
        femaleNames.ForEach(name => scoreOfStudents.FemaleStudentsWithScores.Add(new StudentWithScore { Name = name, Score = "0" }));

        highestScoresPage.WrittenWorks.Clear();
        _workbookService.WrittenWorks.ForEach(score => highestScoresPage.WrittenWorks.Add(score));

        highestScoresPage.PerformanceTasks.Clear();
        _workbookService.PerformanceTasks.ForEach(score => highestScoresPage.PerformanceTasks.Add(score));

        highestScoresPage.Exam = _workbookService.Exam;

        highestScoresPage.SetWrittenWorksPercentage((_workbookService.WeightedScores[0] * 100) + "%");

        highestScoresPage.SetPerformancePercentage((_workbookService.WeightedScores[1] * 100) + "%");

        highestScoresPage.SetExamPercentage((_workbookService.WeightedScores[2] * 100) + "%");

        rbtnStudents.IsEnabled = rbtnScores.IsEnabled = true;

        homePage.SetTrack(Array.IndexOf(WorkbookService.tracks, _workbookService.Track));

      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }

      SetLoading(false);
      homePage.SetLoading(false);
    }


    private async void StudentsPage_AddMale(object sender, string newName) {
      try {
        SetLoading(true);
        studentsPage.EnableAddButton(false);
        studentsPage.MaleNames.Clear();
        await Task.Run(() => {
          _workbookService.AppendAndSortNames(newName).ForEach(name => studentsPage.MaleNames.Add(name));
        });
        studentsPage.EnableAddButton();
        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private async void StudentsPage_AddFemale(object sender, string newName) {
      try {
        SetLoading(true);
        studentsPage.EnableAddButton(false);
        studentsPage.FemaleNames.Clear();
        await Task.Run(() => {
          _workbookService.AppendAndSortNames(newName, false).ForEach(name => studentsPage.FemaleNames.Add(name));
        });
        studentsPage.EnableAddButton();
        SetLoading(false);
      } catch (Exception ex) {
        ShowError("Error", ex.Message);
      }
    }

    private void ShowError(string title, string error) {
      ErrorTitle.Text = title;
      ErrorDescription.Text = error;
      RootDialog.IsOpen = true;
    }

    private void SetLoading(bool isLoading) {
      if (isLoading) {
        MainContent.IsEnabled = false;     
        MyProgressBar.Visibility = Visibility.Visible;
      } else {
        MainContent.IsEnabled = true;       
        MyProgressBar.Visibility = Visibility.Collapsed;
      }
      MyProgressBar.IsIndeterminate = isLoading;
    }
  }
}
