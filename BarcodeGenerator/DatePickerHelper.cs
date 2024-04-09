using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace BarcodeGenerator
{
    public class DatePickerHelper
    {
        public static readonly DependencyProperty EnableFastInputProperty =
         DependencyProperty.RegisterAttached("EnableFastInput", typeof(bool), typeof(DatePickerHelper),
                new FrameworkPropertyMetadata((bool)false,
                new PropertyChangedCallback(OnEnableFastInputChanged)));

        public static bool GetEnableFastInput(DependencyObject d)
        {
            return (bool)d.GetValue(EnableFastInputProperty);
        }

        public static void SetEnableFastInput(DependencyObject d, bool value)
        {
            d.SetValue(EnableFastInputProperty, value);
        }

        private static void OnEnableFastInputChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var datePicker = d as DatePicker;
            if (datePicker != null)
            {
                if ((bool)e.NewValue)
                {
                    datePicker.DateValidationError += DatePickerOnDateValidationError;
                }
                else
                {
                    datePicker.DateValidationError -= DatePickerOnDateValidationError;
                }
            }
        }

        private static void DatePickerOnDateValidationError(object sender, DatePickerDateValidationErrorEventArgs e)
        {
            var datePicker = sender as DatePicker;
            if (datePicker != null)
            {
                var text = e.Text;
                DateTime dateTime;
                if (DateTime.TryParseExact(text, "yyyyMMdd", CultureInfo.CurrentUICulture, DateTimeStyles.None, out dateTime))
                {
                    datePicker.SelectedDate = dateTime;
                }
            }
        }

    }
}
