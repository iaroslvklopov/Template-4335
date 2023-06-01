using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
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

using Word = Microsoft.Office.Interop.Word;

namespace Template_4335.Windows.Klopov4335
{
    /// <summary>
    /// Логика взаимодействия для WordPage.xaml
    /// </summary>
    public partial class WordPage : System.Windows.Controls.Page
    {
        public WordPage()
        {
            InitializeComponent();
            DBGridModel.ItemsSource = IsrpoEntities.GetContext().sotrydniki.AsEnumerable().ToList();
        }
        private async void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Template_4335", "Windows", "Klopov4335", "5.json");
            using (var db = new IsrpoEntities())
            {
                var emp = await JsonSerializer.DeserializeAsync<List<sotrydniki>>(new FileStream(path, FileMode.Open));
                foreach (sotrydniki item in emp)
                {
                    var employee = new sotrydniki
                    {
                        role_e = item.role_e,
                        fio_e = item.fio_e,
                        login_e = item.login_e,
                        pass_e = GetHashString(item.pass_e)
                    };

                    db.sotrydniki.Add(employee);
                }
                try
                {
                    db.SaveChanges();
                    MessageBox.Show("Данные импортированы!");
                    DBGridModel.ItemsSource = IsrpoEntities.GetContext().sotrydniki.AsEnumerable().ToList();
                }
                catch (DbEntityValidationException ex)
                {
                    MessageBox.Show(ex.Message);
                }
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

        private void ExportBtn_Click(object sender, object e)
        {
            #region Объявление листов
            var first = new List<sotrydniki>();
            var second = new List<sotrydniki>();
            var third = new List<sotrydniki>();
            #endregion

            using (var isrpoEntities = new IsrpoEntities())
            {
                #region Сортировка
                first = isrpoEntities.sotrydniki.ToList().Where(sR => sR.role_e == "Менеджер").ToList();
                second = isrpoEntities.sotrydniki.ToList().Where(sR => sR.role_e == "Администратор").ToList();
                third = isrpoEntities.sotrydniki.ToList().Where(sR => sR.role_e == "Клиент").ToList();
                #endregion

                var app = new Word.Application();
                var document = app.Documents.Add();

                #region Заполнение первой таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Менеджеры";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, first.Count() + 1, 2);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Логин";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Пароль";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in first)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.login_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.pass_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение второй таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Администраторы";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, second.Count() + 1, 2);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Логин";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Пароль";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in second)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.login_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.pass_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение третьей таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Клиенты";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, third.Count() + 1, 2);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Логин";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Пароль";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in third)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.login_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.pass_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
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
