using System;
using System.Globalization;
using System.Windows.Data;

namespace XnrgyEngineeringAutomationTools;

public class ProgressToWidthConverter : IMultiValueConverter
{
	public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
	{
		if (values != null && values.Length >= 3 && values[0] is int num && values[1] is int num2 && values[2] is double num3 && num2 > 0)
		{
			double num4 = (double)num / (double)num2;
			return num3 * num4;
		}
		return 0.0;
	}

	public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
	{
		throw new NotImplementedException();
	}
}
