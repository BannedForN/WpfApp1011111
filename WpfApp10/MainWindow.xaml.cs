using System;
using System.Data;
using System.IO;
using System.Net.Mail;
using System.Windows;
using System.Windows.Controls;
using Spire.Xls;

namespace WpfApp10
{
    public partial class MainWindow : Window
    {
        public static string currentFilePath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void NewFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Create a new Excel file
            dataGrid.ItemsSource = null;
            dataGrid.Columns.Clear();
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Open an existing Excel file
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                currentFilePath = openFileDialog.FileName;
                Workbook excelFile = new Workbook();
                excelFile.LoadFromFile(currentFilePath);
                Worksheet sheet = excelFile.Worksheets[0];
                CellRange locatedRange = sheet.AllocatedRange;
                var dataTable = sheet.ExportDataTable(locatedRange, true);
                dataGrid.ItemsSource = dataTable.DefaultView;
            }
        }

        private void SaveFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Save the current Excel file
            if (string.IsNullOrEmpty(currentFilePath))
            {
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx"
                };
                if (saveFileDialog.ShowDialog() == true)
                {
                    currentFilePath = saveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }

            var dataTable = dataGrid.ItemsSource as DataView;

            Workbook excelFile = new Workbook();
            excelFile.Worksheets.Clear();
            Worksheet sheet = excelFile.Worksheets.Add("Лист 1");
            sheet.InsertDataView(dataTable, true, 1,1);
            excelFile.SaveToFile(currentFilePath, Spire.Xls.FileFormat.Version2016);
        }

        private void SendEmailButton_Click(object sender, RoutedEventArgs e)
        {
            SendWindow send = new SendWindow();
            send.ShowDialog();
        }
    }
}