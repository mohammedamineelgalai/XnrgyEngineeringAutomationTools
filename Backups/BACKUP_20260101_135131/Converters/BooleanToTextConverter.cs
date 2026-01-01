using System;
using System.Globalization;
using System.Windows.Data;

namespace XnrgyEngineeringAutomationTools;

public class BooleanToTextConverter : IValueConverter
{
	public string TrueText { get; set; } = "Connecté";

	public string FalseText { get; set; } = "Déconnecté";

	public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
	{
		if (value is bool flag)
		{
			return flag ? TrueText : FalseText;
		}
		return FalseText;
	}

	public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
	{
		throw new NotImplementedException();
	}
}
