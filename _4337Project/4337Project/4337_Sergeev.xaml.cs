using Microsoft.Win32;
using System;
using System.Windows;

namespace _4337Project
{
    public partial class _4337_Sergeev : Window
    {
        public _4337_Sergeev()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "JSON Files (*.json)|*.json"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                string connectionString = "Server=FOCSTER;Database=ISRPO;User Id=your_username;Integrated Security=True;";
                string tableName = "Orders";

                try
                {
                    Sergeev_Export_Import.ImportData(filePath, connectionString, tableName);
                    MessageBox.Show("Данные успешно импортированы!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте: {ex.Message}");
                }
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Word Documents (*.docx)|*.docx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string outputFilePath = saveFileDialog.FileName;
                string connectionString = "Server=FOCSTER;Database=ISRPO;User Id=your_username;Integrated Security=True;";
                string tableName = "Orders";

                try
                {
                    Sergeev_Export_Import.ExportData(connectionString, tableName, outputFilePath);
                    MessageBox.Show("Данные успешно экспортированы!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте: {ex.Message}");
                }
            }
        }
    }
}
