using System.Windows;

namespace Group4338
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonAuthor_Click(object sender, RoutedEventArgs e)
        {
            var authorWindow = new GurkinaArina4338();
            authorWindow.ShowDialog();
        }
    }
}