using Microsoft.Xaml.Behaviors;
using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ThwLib
{
    public class DateTimePickerBehavior : Behavior<WindowsFormsHost>
    {
        public DateTime Value
        {
            get => (DateTime)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        public static readonly DependencyProperty ValueProperty = DependencyProperty.Register(
            nameof(Value),
            typeof(DateTime),
            typeof(DateTimePickerBehavior),
            new FrameworkPropertyMetadata(
                DateTime.Now,
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                OnValueChanged));

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var behavior = (DateTimePickerBehavior)d;
            if (behavior.AssociatedObject?.Child is DateTimePicker dateTimePicker)
            {
                if (!dateTimePicker.Value.Equals((DateTime)e.NewValue))
                    dateTimePicker.Value = (DateTime)e.NewValue;
            }
        }

        private void DateTimePicker_ValueChanged(object? sender, EventArgs e)
        {
            if (sender is DateTimePicker dateTimePicker)
            {
                Value = dateTimePicker.Value;
            }
        }

        protected override void OnAttached()
        {
            base.OnAttached();
            if (AssociatedObject.Child is DateTimePicker dateTimePicker)
            {
                dateTimePicker.ValueChanged += DateTimePicker_ValueChanged;
                dateTimePicker.Value = Value;

            }
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            if (AssociatedObject.Child is DateTimePicker dateTimePicker)
            {
                dateTimePicker.ValueChanged -= DateTimePicker_ValueChanged;
            }
        }
    }
}
