
using System.Windows;
using System.Windows.Forms;

namespace xamlTool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow 
    {
        public MainWindow()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
        }

        private void SimpleButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择文件";
            openFileDialog.Filter = "Excel文件|*.xls; *.xlsx|所有文件|*.*";
            openFileDialog.FileName = string.Empty;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "zip";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            string fileName = openFileDialog.FileName;
            this.showBox.Text = fileName;
        }

        private void Click_Translate(object sender, RoutedEventArgs e)
        {
            string filePath = this.showBox.Text;
            string changedName = null;
            bool isSuccess = false;
            if (filePath != "请选择需要转换格式的Excel文件")
            {
                ExcelReader reader = new ExcelReader();
                isSuccess = reader.TranslateFunction(filePath, out changedName);
                if (isSuccess)
                {
                    System.Windows.MessageBox.Show(System.Windows.Application.Current.MainWindow, "转换成功！新文件为:" + changedName + "(位于同级目录下)");
                }
                else if (isSuccess == false && changedName != null)
                {
                    System.Windows.MessageBox.Show(System.Windows.Application.Current.MainWindow, "转换失败！未能读取文件:" + changedName );
                }
                else 
                {
                    System.Windows.MessageBox.Show(System.Windows.Application.Current.MainWindow, "转换失败！");
                }
            }
            else
            {
                System.Windows.MessageBox.Show(System.Windows.Application.Current.MainWindow, "请选择目标文件");
            }
        }
    }
}
