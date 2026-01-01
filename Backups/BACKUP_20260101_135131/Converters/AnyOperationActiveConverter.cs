using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace XnrgyEngineeringAutomationTools;

public class AnyOperationActiveConverter : IMultiValueConverter
{
	public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
	{
		if (values != null && values.Length >= 2)
		{
			object obj = values[0];
			bool flag = default(bool);
			int num;
			if (obj is bool)
			{
				flag = (bool)obj;
				num = 1;
			}
			else
			{
				num = 0;
			}
			bool flag2 = (byte)((uint)num & (flag ? 1u : 0u)) != 0;
			obj = values[1];
			bool flag3 = default(bool);
			int num2;
			if (obj is bool)
			{
				flag3 = (bool)obj;
				num2 = 1;
			}
			else
			{
				num2 = 0;
			}
			bool flag4 = (byte)((uint)num2 & (flag3 ? 1u : 0u)) != 0;
			return (!(flag2 || flag4)) ? Visibility.Collapsed : Visibility.Visible;
		}
		return Visibility.Collapsed;
	}

	public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
	{
		throw new NotImplementedException();
	}
}
