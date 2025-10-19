using System.Windows;
using static ThwLib.IDialogService;

namespace ThwLib
{
    public class DialogService : IDialogService
    {
        public void ShowMessage(string message, Severity severity)
        {
            MessageBoxImage icon = severity switch
            {
                Severity.Information => MessageBoxImage.Information,
                Severity.Warning     => MessageBoxImage.Warning,
                Severity.Error       => MessageBoxImage.Error,
                _                    => MessageBoxImage.None
            };
            MessageBox.Show(message, "ThwLib " + severity.ToString(), MessageBoxButton.OK, icon);
        }
    }
}
