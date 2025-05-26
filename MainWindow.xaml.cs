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
<<<<<<< HEAD
=======
using System.Windows.Controls.Primitives;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.ComponentModel;

//https://learn.microsoft.com/ru-ru/office/open-xml/spreadsheet/overview
>>>>>>> progressBar-Changes

using Text = DocumentFormat.OpenXml.Spreadsheet.Text;
using Run = DocumentFormat.OpenXml.Spreadsheet.Run;

//https://learn.microsoft.com/ru-ru/office/open-xml/spreadsheet/overview
//git@github.com:Kloshar/ExcelTextReplacer.git
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

            string oldTxt = "xyzqwe";
            string newTxt = @"ab";
            

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
                        if (ssItem.InnerText == oldTxt) //если текст совпадает с искомой строкой
                        {
                            string str = ssItem.InnerText;

                            //Debug.WriteLine($"Заменяем {ssItem.InnerText} на {newTxt}");

                            //если ssItem.InnerText (txt) совпдадает, то начинается поиск в узле
                            //1. текст может быть без форматирования, тогда он просто в узле <x:t>. Требуется проверка совпадает ли он
                            //2. текст с форматированием будет вложен в отдельных узлах <x:r>. Форматирование тоже, но его не трогаем
                            //3. последовательно проходися по узлам <x:r><x:t></x:t></x:r> выискивая совпадения с искомым текстом

                            foreach (DocumentFormat.OpenXml.OpenXmlElement el in ssItem) //проходим по элементам si
                            {
                                if (el is Text) //если это текст без форматирования
                                {
                                    string txt = ((Text)el).Text;

                                    if (txt.Length == oldTxt.Length) //если все искомые символы в этом элементе
                                    {
                                        ((Text)el).Text = newTxt; //то просто заменяем текст в элементе на новую строку
                                        counter++; //обновляем счётчик
                                        return true; //и выходим из метода
                                    }

                                    if (txt.Length < oldTxt.Length) //если в элементе только часть искомых символов
                                    {
                                        index += txt.Length; //продвигаем индекс на количество заменяемых символов
                                        ((Text)el).Text = newTxt.Substring(writed, index - writed); //то перезаписываем текст в элементе частью новой строки
                                        writed += index; //сохраняем число записанных символов
                                        //переходим к следующему элементу el
                                    }
                                }
                                if (el is Run)
                                {
                                    Text? t = el.Descendants<Text>().FirstOrDefault();

                                    string txt = ((Text)t).Text;

                                    if (txt.Length == oldTxt.Length) //если все искомые символы в этом элементе
                                    {
                                        ((Text)t).Text = newTxt; //то просто заменяем текст в элементе на новую строку
                                        counter++; //обновляем счётчик
                                        return true; //и выходим из метода
                                    }

                                    //нужна проверка, что символы новой строки не израсходованы
                                    //если кончились, то нужно удалить все элементы старой строки

                                    if (txt.Length < oldTxt.Length) //если в элементе только часть искомых символов
                                    {
                                        if(txt.Length >= newTxt.Substring(writed).Length) //если осталось заменить больше, чем осталось в новой строке
                                        {
                                            Debug.WriteLine($"больше - {newTxt.Substring(writed).Length}");
                                            index += newTxt.Substring(writed).Length; //или продвигаем на оставшееся количество символов
                                            
                                        }
                                        else
                                        {
                                            Debug.WriteLine($"меньше - {newTxt.Substring(writed).Length}");
                                            index += txt.Length; //продвигаем индекс на количество заменяемых символов
                                        }                                           

                                        string strPart = newTxt.Substring(writed, index - writed);

                                        Debug.WriteLine($"{newTxt.Substring(writed, index - writed)}");

                                        ((Text)t).Text = newTxt.Substring(writed, index - writed); //то перезаписываем текст в элементе частью новой строки
                                        writed += index; //сохраняем число записанных символов
                                        //переходим к следующему элементу el
                                    }
                                }
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

<<<<<<< HEAD
            if (res)
            {
                Debug.WriteLine("Сделано замен: " + counter + "!");
                //MessageBox.Show("Сделано замен: " + counter + "!");
            }
=======
            oldTxt = replaceWhat.Text;
            newTxt = replaceWith.Text;

            progress.Value = 0;
            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += doWork;
            worker.RunWorkerAsync(10000);            
        }
        void doWork(object sender, DoWorkEventArgs e)
        {
            string[] files = Directory.GetFiles(Environment.CurrentDirectory, "*.xlsx");
            foreach (string file in files)
            {
                //Debug.WriteLine(System.IO.Path.GetFileName(file));

                //ReplaceSymbolsInSharedStringTable(file, oldTxt, newTxt);

                Dispatcher.Invoke(() =>
                {
                    progress.Value += 100/files.Length;
                    progressText.Text = file;
                });
                //Thread.Sleep(500);

            }
            MessageBox.Show("Сделано замен: " + counter + "!");
>>>>>>> progressBar-Changes
        }
    }
}