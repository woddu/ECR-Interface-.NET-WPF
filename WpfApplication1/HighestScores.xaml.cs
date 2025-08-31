using System.Collections.ObjectModel;
using System.Windows.Controls;

namespace WpfApplication1 {
  /// <summary>
  /// Interaction logic for HighestScores.xaml
  /// </summary>
  public partial class HighestScores : UserControl {

    public ObservableCollection<string> WrittenWorks { get; set; } = new ObservableCollection<string>();

    public ObservableCollection<string> PerformanceTasks { get; set; } = new ObservableCollection<string>();

    public HighestScores() {
      InitializeComponent();
      DataContext = this;
    }
  }
}
