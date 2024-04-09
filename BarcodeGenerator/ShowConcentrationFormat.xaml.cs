using System.Windows;

namespace BarcodeGenerator
{
    /// <summary>
    /// ShowConcentrationFormat.xaml 的交互逻辑
    /// </summary>
    public partial class ShowConcentrationFormat : Window
    {
        public bool IsCloseApp;
        public ShowConcentrationFormat()
        {
            InitializeComponent();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!IsCloseApp)
            {
                e.Cancel = true;
                WindowState = WindowState.Normal;
                Hide();
            }
        }
    }
}
