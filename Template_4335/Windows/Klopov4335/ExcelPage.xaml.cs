using Microsoft.Win32;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4335.Windows.Klopov4335
{
    /// <summary>
    /// Логика взаимодействия для ExcelPage.xaml
    /// </summary>
    public partial class ExcelPage : Page
    {
        public ExcelPage()
        {
            InitializeComponent();
            DBGridModel.ItemsSource = IsrpoEntities.GetContext().sotrydniki.AsEnumerable().ToList();
        }
        private void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "Файл Excel|*.xlsx",
                Title = "Выберите файл"
            };
            if (!openFileDialog.ShowDialog() == true)
                return;
            ImportData(openFileDialog.FileName);
            DBGridModel.ItemsSource = IsrpoEntities.GetContext().sotrydniki.AsEnumerable().ToList();
        }

        private static void ImportData(string path)
        {
            try
            {
                var objWorkExcel = new Excel.Application();
                var objWorkBook = objWorkExcel.Workbooks.Open(path);
                var objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
                var lastCell = objWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                var columns = lastCell.Column;
                var rows = lastCell.Row;
                var list = new string[rows, columns];
                for (var j = 0; j < columns; j++)
                {
                    for (var i = 1; i < rows; i++)
                    {
                        list[i, j] = objWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }

                objWorkBook.Close(false, Type.Missing, Type.Missing);
                objWorkExcel.Quit();
                GC.Collect();
                using (var db = new IsrpoEntities())
                {
                    for (var i = 1; i < 11; i++)
                    {
                        var uslugi = new sotrydniki
                        {
                            role_e = list[i, 0].ToString(),
                            fio_e = list[i, 1].ToString(),
                            login_e = list[i, 2].ToString(),
                            pass_e = GetHashString(list[i, 3].ToString())
                        };
                        db.sotrydniki.Add(uslugi);
                    }
                    try
                    {
                        db.SaveChanges();
                        MessageBox.Show("Данные импортированы!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Внимание", MessageBoxButton.OK);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Внимание1", MessageBoxButton.OK);
            }
        }

        private static string GetHashString(string s)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(s);
            MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();
            byte[] byteHash = CSP.ComputeHash(bytes);
            string hash = "";
            foreach (byte b in byteHash)
            {
                hash += string.Format("{0:x2}", b);
            }
            return hash;
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            #region Объявление листов
            var first = new List<sotrydniki>();
            var second = new List<sotrydniki>();
            var third = new List<sotrydniki>();
            var categoriesPriceCount = 3;
            #endregion

            using (var isrpoEntities = new IsrpoEntities())
            {
                #region Сортировка
                first = isrpoEntities.sotrydniki.ToList().Where(sR => sR.role_e == "Менеджер").ToList();
                second = isrpoEntities.sotrydniki.ToList().Where(sR => sR.role_e == "Администратор").ToList();
                third = isrpoEntities.sotrydniki.ToList().Where(sR => sR.role_e == "Клиент").ToList();
                #endregion

                var app = new Excel.Application { SheetsInNewWorkbook = categoriesPriceCount };
                var book = app.Workbooks.Add(Type.Missing);

                #region Создание листов в Excel
                var startRowIndex = 1;
                var sheet1 = app.Worksheets.Item[1];
                sheet1.Name = "Менеджеры";
                var sheet2 = app.Worksheets.Item[2];
                sheet2.Name = "Администраторы";
                var sheet3 = app.Worksheets.Item[3];
                sheet3.Name = "Клиенты";
                #endregion

                #region Создание колонок в Excel
                sheet1.Cells[1][startRowIndex] = "Логин";
                sheet1.Cells[2][startRowIndex] = "Пароль";

                sheet2.Cells[1][startRowIndex] = "Логин";
                sheet2.Cells[2][startRowIndex] = "Пароль";

                sheet3.Cells[1][startRowIndex] = "Логин";
                sheet3.Cells[2][startRowIndex] = "Пароль";
                startRowIndex++;
                #endregion

                #region Заполнение первого листа
                for (var i = 0; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in first)
                    {
                        sheet1.Cells[1][startRowIndex] = item.login_e;
                        sheet1.Cells[2][startRowIndex] = item.pass_e;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение второго листа
                for (var i = 1; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in second)
                    {
                        sheet2.Cells[1][startRowIndex] = item.login_e;
                        sheet2.Cells[2][startRowIndex] = item.pass_e;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение третьего листа
                for (var i = 2; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in third)
                    {
                        sheet3.Cells[1][startRowIndex] = item.login_e;
                        sheet3.Cells[2][startRowIndex] = item.pass_e;
                        startRowIndex++;
                    }
                }
                #endregion

                app.Visible = true;
            }
        }

        private void DeleteDataBtn_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Очистить данные?", "Внимание!", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                using (var isrpoEntities = new IsrpoEntities())
                {
                    isrpoEntities.sotrydniki.RemoveRange(isrpoEntities.sotrydniki.ToList());
                    isrpoEntities.SaveChanges();
                    IsrpoEntities.GetContext().sotrydniki.AsEnumerable().ToList().Clear();
                    foreach (var uslugi in isrpoEntities.sotrydniki.AsEnumerable().ToList())
                    {
                        IsrpoEntities.GetContext().sotrydniki.AsEnumerable().ToList().Add(uslugi);
                    }
                    DBGridModel.ItemsSource = IsrpoEntities.GetContext().sotrydniki.AsEnumerable().ToList();
                }
            }
        }
    }
}
