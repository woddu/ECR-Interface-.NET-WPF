using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace WpfApplication1 {
  public class ScoreToEnabledConverter : IValueConverter {
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
      // value is the TextBox.Text (string)
      if (string.IsNullOrWhiteSpace(value as string))
        return false;

      if (int.TryParse(value.ToString(), out int score))
        return score != 0;

      return false;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
  }
}
