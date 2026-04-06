using Group4337.Models;
using Microsoft.Win32;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SecondWindow second = new SecondWindow();
            second.Show();
            this.Close();
        }

        private void ImportData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };

            if (!(openFileDialog.ShowDialog() == true))
                return;

            string[,] list;

            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(openFileDialog.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;

            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (Isrpolab3Context dbContext = new Isrpolab3Context())
            {
                for (int i = 1; i < _rows; i++)
                {
                    dbContext.Users.Add(new User()
                    {
                        Post = list[i, 0], 
                        UserLogin = list[i, 1]  
                    });
                }
                dbContext.SaveChanges();
            }

            MessageBox.Show("Импорт завершён.", "Сообщение", MessageBoxButton.OK);
        }

        private void ExportData_Click(object sender, RoutedEventArgs e)
        {
            using (Isrpolab3Context dbContext = new Isrpolab3Context())
            {
                var allUsers = dbContext.Users.ToList();
                var usersByPost = allUsers.GroupBy(u => u.Post).ToList();

                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;
                Excel.Workbook workbook = excelApp.Workbooks.Add();

                int sheetIndex = 1;

                foreach (var postGroup in usersByPost)
                {
                    Excel.Worksheet worksheet;
                    if (sheetIndex == 1)
                        worksheet = (Excel.Worksheet)workbook.Sheets[sheetIndex];
                    else
                    {
                        worksheet = (Excel.Worksheet)workbook.Sheets.Add();
                        worksheet.Move(After: workbook.Sheets[workbook.Sheets.Count]);
                    }

                    worksheet.Name = postGroup.Key;

                    worksheet.Cells[1, 1] = "ID";
                    worksheet.Cells[1, 2] = "Должность";
                    worksheet.Cells[1, 3] = "Логин пользователя";

                    int row = 2;
                    foreach (var user in postGroup)
                    {
                        worksheet.Cells[row, 1] = user.Id;
                        worksheet.Cells[row, 2] = user.Post;
                        worksheet.Cells[row, 3] = user.UserLogin;
                        row++;
                    }

                    sheetIndex++;
                }

                if (usersByPost.Count == 0)
                {
                    Excel.Worksheet emptySheet = (Excel.Worksheet)workbook.Sheets[1];
                    emptySheet.Name = "Нет данных";
                    emptySheet.Cells[1, 1] = "Нет пользователей";
                }
            }
        }
    }
}