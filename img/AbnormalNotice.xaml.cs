using System.Windows;
using System.Windows.Controls;


namespace ReportPro.img
{
    /// <summary>
    /// AbnormalNotice.xaml 的交互逻辑
    /// </summary>
    public partial class AbnormalNotice : UserControl
    {

        public AbnormalNotice()
        {
            InitializeComponent();
        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {


            warn.Visibility = Visibility.Hidden;
        }
        public void close()
        {
            this.IsEnabled = false;
        }
    }
}
