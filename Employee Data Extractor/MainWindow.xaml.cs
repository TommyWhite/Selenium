using EmployeeInfoGrabber;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Resources;
using System.Windows;
using Forms = System.Windows.Forms;

namespace Employee_Data_Extractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {
        private HtmlHandler htmlHandler;
        private ExcelHandler xlsHanlder;
        private DataGrabber grabber;

        public MainWindow()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
            InitCoreLogic();

            var bs = Properties.Resources.chromedriver;
            File.WriteAllBytes("chromedriver.exe", bs);
        }

        private void InitCoreLogic()
        {
            htmlHandler = new HtmlHandler();
            xlsHanlder = new ExcelHandler();
            //TODO: Remove in production.
            if (Debugger.IsAttached)
            {

                txtBoxToInputExcel.Text = @"C:\Users\artemm\Desktop\EmployeeInfoGrabber\InfoGrabber\bin\Debug\Resources\Input\input.xlsx";
                txtBoxToHtmlReport.Text = @"C:\Users\artemm\Desktop\EmployeeInfoGrabber\InfoGrabber\bin\Debug\Resources\Output";
            }
        }

        private void RaiseErorrMessageBox(string message, string caption)
        {
            MessageBox.Show(message, caption, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void btnBrowseInputExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = @"Excel files (*.xlsx;*xls)|*.xlsx;*xls";
                if (openFileDialog.ShowDialog() == true)
                {
                    txtBoxToInputExcel.Text = openFileDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                RaiseErorrMessageBox("Failed to open excel file.", "Failed");
            }
        }

        private void btnBrowseOutpuDir_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Forms.FolderBrowserDialog browseDialog = new Forms.FolderBrowserDialog();
                browseDialog.ShowDialog();
                txtBoxToHtmlReport.Text = browseDialog.SelectedPath;
            }
            catch (Exception ex)
            {
                RaiseErorrMessageBox("Failed to specify choosen directory.", "Failed");
            }
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void btnGenerateReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var data = htmlHandler.ReadReportData(txtBoxToHtmlReport.Text);

                using (xlsHanlder)
                    xlsHanlder.CreateXlsDoc(txtBoxToHtmlReport.Text, "A2", $"P{data.GetLength(0)}", data);
            }
            catch (Exception ex)
            {
                RaiseErorrMessageBox("Report generation is failed. Check if you have specified the folder with html data.", "Failure");
            }
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                InitCoreLogic();
                grabber = null;
            }
            catch (Exception ex)
            {
                RaiseErorrMessageBox("Failed to stop the program.", "Failure");
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                if (!File.Exists(txtBoxToInputExcel.Text) || !Directory.Exists(txtBoxToHtmlReport.Text))
                {
                    throw new FileNotFoundException();
                }
                using (grabber = new DataGrabber())
                {
                    grabber.Run(txtBoxToInputExcel.Text, txtBoxToHtmlReport.Text);
                }
            }
            catch (Exception ex)
            {
                RaiseErorrMessageBox(!string.IsNullOrEmpty(ex.Message) ? ex.Message : "Check if all file specified.", "Failure");
            }
        }

        public void Dispose()
        {
            xlsHanlder.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}