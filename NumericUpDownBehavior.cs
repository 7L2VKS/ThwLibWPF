using Microsoft.Xaml.Behaviors;
using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ThwLib
{
    public class NumericUpDownBehavior : Behavior<WindowsFormsHost>
    {
        public decimal Value
        {
            get => (decimal)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        public static readonly DependencyProperty ValueProperty = DependencyProperty.Register(
            nameof(Value),
            typeof(decimal),
            typeof(NumericUpDownBehavior),
            new FrameworkPropertyMetadata(
                default(decimal),
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                OnValueChanged));

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var behavior = (NumericUpDownBehavior)d;
            if (behavior.AssociatedObject?.Child is NumericUpDown numericUpDown)
            {
                var newValue = (decimal)e.NewValue;
                if (newValue >= numericUpDown.Minimum && newValue <= numericUpDown.Maximum
                 && numericUpDown.Value != newValue)
                {
                    numericUpDown.Value = newValue;
                }
            }
        }

        private void NumericUpDown_ValueChanged(object? sender, EventArgs e)
        {
            if (sender is NumericUpDown numericUpDown)
            {
                Value = numericUpDown.Value;
            }
        }

        protected override void OnAttached()
        {
            base.OnAttached();
            if (AssociatedObject.Child is NumericUpDown numericUpDown)
            {
                numericUpDown.ValueChanged += NumericUpDown_ValueChanged;
                numericUpDown.Value = Value;
            }
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            if (AssociatedObject.Child is NumericUpDown numericUpDown)
            {
                numericUpDown.ValueChanged -= NumericUpDown_ValueChanged;
            }
        }
    }
}
