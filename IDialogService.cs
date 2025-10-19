namespace ThwLib
{
    public interface IDialogService
    {
        void ShowMessage(string message, Severity severity);
        public enum Severity { Information, Warning, Error }
    }
}
