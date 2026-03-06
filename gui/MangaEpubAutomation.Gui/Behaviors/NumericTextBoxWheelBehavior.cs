using System.Globalization;
using System.Windows;
using System.Windows.Input;
using TextBox = System.Windows.Controls.TextBox;

namespace MangaEpubAutomation.Gui.Behaviors;

public static class NumericTextBoxWheelBehavior
{
    public static readonly DependencyProperty IsEnabledProperty =
        DependencyProperty.RegisterAttached(
            "IsEnabled",
            typeof(bool),
            typeof(NumericTextBoxWheelBehavior),
            new PropertyMetadata(false, OnIsEnabledChanged));

    public static readonly DependencyProperty MinimumProperty =
        DependencyProperty.RegisterAttached(
            "Minimum",
            typeof(int),
            typeof(NumericTextBoxWheelBehavior),
            new PropertyMetadata(int.MinValue));

    public static readonly DependencyProperty MaximumProperty =
        DependencyProperty.RegisterAttached(
            "Maximum",
            typeof(int),
            typeof(NumericTextBoxWheelBehavior),
            new PropertyMetadata(int.MaxValue));

    public static readonly DependencyProperty StepProperty =
        DependencyProperty.RegisterAttached(
            "Step",
            typeof(int),
            typeof(NumericTextBoxWheelBehavior),
            new PropertyMetadata(1));

    public static bool GetIsEnabled(DependencyObject obj) => (bool)obj.GetValue(IsEnabledProperty);
    public static void SetIsEnabled(DependencyObject obj, bool value) => obj.SetValue(IsEnabledProperty, value);

    public static int GetMinimum(DependencyObject obj) => (int)obj.GetValue(MinimumProperty);
    public static void SetMinimum(DependencyObject obj, int value) => obj.SetValue(MinimumProperty, value);

    public static int GetMaximum(DependencyObject obj) => (int)obj.GetValue(MaximumProperty);
    public static void SetMaximum(DependencyObject obj, int value) => obj.SetValue(MaximumProperty, value);

    public static int GetStep(DependencyObject obj) => (int)obj.GetValue(StepProperty);
    public static void SetStep(DependencyObject obj, int value) => obj.SetValue(StepProperty, value);

    private static void OnIsEnabledChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not TextBox textBox) return;

        if ((bool)e.NewValue) textBox.PreviewMouseWheel += OnPreviewMouseWheel;
        else textBox.PreviewMouseWheel -= OnPreviewMouseWheel;
    }

    private static void OnPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (sender is not TextBox textBox) return;

        var currentValue = ParseInt(textBox.Text, 0);
        var step = Math.Max(1, GetStep(textBox));
        var min = GetMinimum(textBox);
        var max = GetMaximum(textBox);
        var nextValue = currentValue + (e.Delta > 0 ? step : -step);

        if (nextValue < min) nextValue = min;
        if (nextValue > max) nextValue = max;

        textBox.Text = nextValue.ToString(CultureInfo.InvariantCulture);
        textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
        e.Handled = true;
    }

    private static int ParseInt(string? value, int defaultValue)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)) return parsed;
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out parsed)) return parsed;
        return defaultValue;
    }
}
