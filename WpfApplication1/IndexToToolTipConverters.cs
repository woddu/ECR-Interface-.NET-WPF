using System;
using System.Globalization;
using System.Windows.Data;

namespace WpfApplication1 {  
  public class WWIndexToToolTipConverter : IValueConverter {
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
      if (value == null)
        return string.Empty;

      return $"View Scores of Written Works #{uint.Parse(value.ToString()) + 1u}";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
      // Not needed for one-way binding
      throw new NotImplementedException();
    }
  }

  public class PTIndexToToolTipConverter : IValueConverter {
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
      if (value == null)
        return string.Empty;

      return $"View Scores of Performance Tasks #{uint.Parse(value.ToString()) + 1u}";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
      // Not needed for one-way binding
      throw new NotImplementedException();
    }
  }

}
