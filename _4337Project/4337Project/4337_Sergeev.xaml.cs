using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration.Provider;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace _4337Project
{
    /// <summary>
    /// Логика взаимодействия для _4337_Sergeev.xaml
    /// </summary>
    public partial class _4337_Sergeev : Window
    {
        public _4337_Sergeev()
        {
            InitializeComponent();
        }
        private void Import_Click(object sender, RoutedEventArgs e)
        {
            // Открытие диалога выбора файла для импорта
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                string connectionString = "Server=FOCSTER;Database=ISRPO;User Id=your_username;Integrated Security=True;";
                string tableName = "Orders";

                // Вызов метода импорта данных
                Sergeev_Export_Import.ImportData(filePath, connectionString, tableName);

                MessageBox.Show("Данные успешно импортированы!");
            }
        }

        // Обработчик для кнопки "Экспорт"
        private void Export_Click(object sender, RoutedEventArgs e)
        {
            // Открытие диалога выбора места для сохранения файла
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string outputFilePath = saveFileDialog.FileName;
                string connectionString = "Server=FOCSTER;Database=ISRPO;User Id=your_username;Integrated Security=True;";
                string tableName = "Orders";

                // Вызов метода экспорта данных
                Sergeev_Export_Import.ExportData(connectionString, tableName, outputFilePath);

                MessageBox.Show("Данные успешно экспортированы!");
            }
        }
    }
}
