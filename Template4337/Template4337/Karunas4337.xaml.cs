using System.IO;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Template4337.Migrations;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4337
{
    /// <summary>
    /// Логика взаимодействия для Karunas4337.xaml
    /// </summary>
    public partial class Karunas4337 : Window
    {
        public Karunas4337()
        {
            InitializeComponent();
        }

        private void ImportData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int _columns = (int)lastCell.Column;
                int _rows = (int)lastCell.Row;
                string[,] list = new string[_rows, _columns];

                for (int i = 0; i < _rows; i++)
                {
                    for (int j = 0; j < _columns; j++)
                    {
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();

                using (ServiceContext context = new ServiceContext())
                {
                    for (int i = 0; i < _rows; i++)
                    {
                        string priceString = list[i, 4];
                        if (int.TryParse(priceString, out int price))
                        {
                            context.Services.Add(new Service()
                            {
                                Id = Convert.ToInt32(list[i, 0]),
                                Name = list[i, 1],
                                Type = list[i, 2],
                                Code = list[i, 3],
                                Price = price
                            });
                        }
                        else
                        {
                            //MessageBox.Show($"Неверный формат цены в строке {i + 1}: '{priceString}'.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    context.SaveChanges();
                    MessageBox.Show("Импорт завершен", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            List<Service> services;

            using (ServiceContext service = new ServiceContext())
            {
                services = service.Services.OrderBy(s => s.Price).ToList();
            }


            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            Excel.Worksheet category1Sheet = app.Worksheets.Item[1];
            category1Sheet.Name = "Категория 1 (от 0 до 350)";

            Excel.Worksheet category2Sheet = app.Worksheets.Item[2];
            category2Sheet.Name = "Категория 2 (от 350 до 800)";

            Excel.Worksheet category3Sheet = app.Worksheets.Item[3];
            category3Sheet.Name = "Категория 3 (от 800 и выше)";

            var groupedByCategory = services.GroupBy(user =>
            {
                if (user.Price <= 350)
                    return "Категория 1";
                else if (user.Price <= 800)
                    return "Категория 2";
                else
                    return "Категория 3";
            });

            foreach (var group in groupedByCategory)
            {
                Excel.Worksheet worksheet = null;

                switch (group.Key)
                {
                    case "Категория 1":
                        worksheet = category1Sheet;
                        break;
                    case "Категория 2":
                        worksheet = category2Sheet;
                        break;
                    case "Категория 3":
                        worksheet = category3Sheet;
                        break;
                }

                int startRowIndex = 1;
                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "Вид услуги";
                worksheet.Cells[4][startRowIndex] = "Стоимость";
                startRowIndex++;
                
                foreach (Service service in group)
                {
                    worksheet.Cells[1][startRowIndex] = service.Id;
                    worksheet.Cells[2][startRowIndex] = service.Name;
                    worksheet.Cells[3][startRowIndex] = service.Type;
                    worksheet.Cells[4][startRowIndex] = service.Price;
                    startRowIndex++;
                }

                worksheet.Columns.AutoFit();
            }

            app.Visible = true;
            this.Close();
        }
    }
}