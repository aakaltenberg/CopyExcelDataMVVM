using System.Windows;
using System.Windows.Controls;

namespace CopyExcelData
{
    public partial class ExcelCopyWindow : Window
    {
        public ExcelCopyWindow()
        {
            InitializeComponent();
        }

        private void ControlRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && rb.DataContext is ExcelCopyViewModel model && model.DroughtChecked)
            {
                model.DroughtChecked = !model.ControlChecked;
            }
        }

        private void DroughtRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && rb.DataContext is ExcelCopyViewModel model && model.ControlChecked)
            {
                model.ControlChecked = !model.DroughtChecked;
            }
        }
    }
}