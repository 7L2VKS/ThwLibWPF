using System.Windows;

namespace ThwLib
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.DataContext = new MainWindowViewModel(new DialogService());
            InitializeComponent();
        }
    }
}