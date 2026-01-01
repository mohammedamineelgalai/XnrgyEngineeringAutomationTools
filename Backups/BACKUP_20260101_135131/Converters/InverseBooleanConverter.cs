using System;
using System.Globalization;
using System.Windows.Data;

namespace XnrgyEngineeringAutomationTools;

public class InverseBooleanConverter : IValueConverter
{
	public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
	{
		if (value is bool flag)
		{
			return !flag;
		}
		return false;
	}

	public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
	{
		if (value is bool flag)
		{
			return !flag;
		}
		return false;
	}
}
