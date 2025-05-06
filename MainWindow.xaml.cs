using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.EMMA;
using System.Xml;

using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

//https://learn.microsoft.com/ru-ru/office/open-xml/spreadsheet/overview

/*
 Возможные проблемы:
1. двойные кавычки нужно экранировать
2. строки с переносом
3. замена части строки
4. графический интерфейс
 */

namespace ExcelTextReplacer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string path = @"book.xlsx";
        //string path = @"93-24-2030_РКМ_Койда_1_безопасность.xlsx";
        int index = 0; //курсор искомой строки
        int writed = 0; //записанные символы
        int counter = 0;
        public MainWindow()
        {
            InitializeComponent();

            string oldTxt = "abcd";
            string newTxt = @"xyzzzz";
            

            replaceWhat.Text = oldTxt;
            replaceWith.Text = newTxt;

            //string? val = GetCellValue(path, "Плановая2", "A1");
            //bool res = CheckCellString(path, oldTxt, newTxt);
            //bool res1 = CheckCellString(path, "Цена продукции (без НДС)", "helloo");
            //Debug.WriteLine(res);
        }
        bool CheckCellString(string filepath, string oldTxt, string newTxt)
        {
            string sheetName = "Плановая2"; //пока ищем только на этом листе
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(filepath, true)) //открываем файл ecxel
                {
                    WorkbookPart? wbPart = excelDoc.WorkbookPart; //получаем часть
                    //Sheets? sheets = wbPart?.Workbook.Sheets; //получаем страницы
                    //Sheet? sheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name.Value.StartsWith(sheetName)).FirstOrDefault(); //поиск листа по имени
                    //WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(sheet?.Id!); //часть листа

                    //далее получаем все узлы на листе
                    SharedStringTablePart? ssPart = wbPart.SharedStringTablePart;

                    //Debug.WriteLine(ssPart.SharedStringTable.OuterXml);

                    foreach (SharedStringItem ssItem in ssPart.SharedStringTable) //перебираем строки (элементы) в таблице строк
                    {
                        string unitedtxt = ssItem.InnerText; //запишем в переменную текст из ячейки
                        unitedtxt = unitedtxt.ReplaceLineEndings(); //заменяем в тексте из ячейки переносы на нормальные \r\n для корректного сравнения с искомой строкой

                        if (unitedtxt == oldTxt) //если текст совпадает с искомой строкой
                        {
                            string str = ssItem.InnerText;

                            //Debug.WriteLine($"Заменяем {unitedtxt} на {newTxt}");

                            //если ssItem.InnerText (txt) совпдадает, то начинается поиск в узле
                            //1. текст может быть без форматирования, тогда он просто в узле <x:t>. Требуется проверка совпадает ли он
                            //2. текст с форматированием будет вложен в отдельных узлах <x:r>. Форматирование тоже, но его не трогаем
                            //3. последовательно проходися по узлам <x:r><x:t></x:t></x:r> выискивая совпадения с искомым текстом

                            foreach (DocumentFormat.OpenXml.OpenXmlElement el in ssItem) //проходим по элементам si
                            {
                                if (el is Text)
                                {
                                    string txt = ((Text)el).Text;

                                    for (int i = 0; i < txt.Length; i++) //перебираем символы текста
                                    {
                                        if(txt[i] == oldTxt[index]) //если символ совпадает с символом из искомой строки
                                        {
                                            Debug.WriteLine(txt[i]);
                                            index++; //продвигаем каретку в искомой строке
                                        }
                                    }

                                    //если количество символов равно искомому количеству, то весь текст здесь
                                    //можно присваивать текст. в любом случае перезаписываем новые символы

                                    if(index > newTxt.Length)
                                    {
                                        ((Text)el).Text = newTxt.Substring(writed, newTxt.Length); //записываем пройденное количество символов
                                    }
                                    else
                                    {
                                        ((Text)el).Text = newTxt.Substring(writed, index); //записываем пройденное количество символов
                                    }
                                        
                                }
                                if (el is DocumentFormat.OpenXml.Spreadsheet.Run)
                                {
                                    //Debug.WriteLine($"Run!");
                                }

                                //Debug.WriteLine(el.GetType());
                                //el.InnerText = new DocumentFormat.OpenXml.Spreadsheet.Text("");
                            }

                            //ssItem.Text = new DocumentFormat.OpenXml.Spreadsheet.Text(newTxt); //добавилось в начало...
                            counter++;
                        }
                    }
                    //excelDoc.Save(); //сохраняет документ excel, но таблица строк сохраняется сама
                }
                return true;
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        private void userWindow_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                Close();
            }
        }
        private void replaceBtn_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                Close();
            }
        }
        private void replaceBtn_Click(object sender, RoutedEventArgs e)
        {
            bool res = CheckCellString(path, replaceWhat.Text, replaceWith.Text);

            if (res)
            {
                //Debug.WriteLine("Сделано замен: " + counter + "!");
                //MessageBox.Show("Сделано замен: " + counter + "!");
            }
        }
    }
}