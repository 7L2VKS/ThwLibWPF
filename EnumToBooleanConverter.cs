using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ThwLib
{
    public class EnumToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null || !Enum.IsDefined(value.GetType(), value))
            {
                return DependencyProperty.UnsetValue;
            }

            if (Enum.TryParse(value.GetType(), parameter.ToString(), out object? result))
            {
                return Equals(result, value);
            }
            return DependencyProperty.UnsetValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
            {
                return DependencyProperty.UnsetValue;
            }

            if (Enum.TryParse(targetType, parameter.ToString(), out object? result))
            {
                if (value is bool isChecked)
                {
                    if (isChecked)
                    {
                        return result;
                    }
                    else
                    {
                        return Binding.DoNothing;
                    }
                }
            }
            return DependencyProperty.UnsetValue;
        }
    }

}
