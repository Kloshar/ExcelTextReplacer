using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Xml;
using System.Windows.Controls.Primitives;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.ComponentModel;
//https://learn.microsoft.com/ru-ru/office/open-xml/spreadsheet/overview

namespace ExcelTextReplacer
{
    public partial class MainWindow : Window
    {
        string path = @"book.xlsx";
        //string path = @"93-24-2030_РКМ_Койда_1_безопасность.xlsx";
        int counter = 0;
        string oldTxt = "Барышева";
        string newTxt = "!!!";
        int filesNumber = 0;

        BackgroundWorker worker;

        public MainWindow()
        {
            InitializeComponent();

            replaceWhat.Text = oldTxt;
            replaceWith.Text = newTxt;

            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += doWork;

            //replaceBtn.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent)); //автоматическое нажатие кнопки начала замены
        }

        void ReplaceSymbolsInSharedStringTable(string filepath, string oldTxt, string newTxt)
        {
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(filepath, true)) //открываем файл ecxel
                {
                    WorkbookPart? wbPart = excelDoc.WorkbookPart; //получаем часть книги
                    SharedStringTablePart? ssPart = wbPart.SharedStringTablePart; //далее получаем все узлы на листе

                    //Debug.WriteLine(ssPart.SharedStringTable.OuterXml); //выводим xml узлов

                    foreach (SharedStringItem siItem in ssPart.SharedStringTable) //перебираем строки в таблице строк
                    {
                        Debug.WriteLine($"Текст: {siItem.InnerText}"); //выводим текст
                        Debug.WriteLine($"oldTxt: {oldTxt}"); //выводим текст

                        var stripped = from c in oldTxt.ToCharArray() where c != '\r' select c; //убираем из oldTxt все '\r'
                        oldTxt = new string(stripped.ToArray()); //сохраняем в переменную искомого текста

                        if (siItem.InnerText.Contains(oldTxt)) //если часть текста элемента si совпадает с искомой строкой
                        {
                            string txt = siItem.InnerText; //записываем текст из строки
                            int start = txt.IndexOf(oldTxt); //получаем индекс начала найденной строки
                            int len = oldTxt.Length; //количество заменяемых (в том числе удаляемых) символов
                            Queue<char> newTextChars = new Queue<char>(); //очередь символов из новой строки для замены
                            foreach (char c in newTxt) newTextChars.Enqueue(c); //добавляем новые символы в очередь

                            //Debug.WriteLine($"len = {len}");

                            for (int i = 0; i < siItem.Count(); i++) //перебор блоков с кусками форматированного текста
                            {
                                DocumentFormat.OpenXml.OpenXmlElement el = siItem.ChildElements[i]; //один из блоков с символами

                                //Debug.WriteLine($"Работаем с элементом: {el.InnerXml}");

                                Text? t = null; //переменная для хранения текста блока
                                if (el is Text) t = (Text?)el; //если это текст без форматирования
                                if (el is Run) t = el.Descendants<Text>().FirstOrDefault(); //если это прогон, получаем потомка типа текст

                                for (int j = 0; j < t.Text.Length; j++) //далее перебираем символы пока индекс не уменьшится до нуля
                                {
                                    if (start > 0) //если перебрали меньше символов, чем было в искомой строке
                                    {
                                        start--; //то уменьшаем на один
                                        //Debug.WriteLine($"Обрабатывается символ: {t.Text[j]}, start = {start}");
                                    }
                                    else
                                    {
                                        Debug.WriteLine($"Обрабатывается символ: {t.Text[j]}, len = {len}, newTextChars.Count = {newTextChars.Count}");

                                        //заменяем символ, если количество заменяемых символов осталось и есть чем заменять
                                        if (len > 0 && newTextChars.Count > 0)
                                        {
                                            t.Text = t.Text.Remove(j, 1).Insert(j, newTextChars.Dequeue().ToString());
                                            len--; //уменьшаем количество символов, которые нужно заменить, так как заменили символ
                                            //Debug.WriteLine($"Символы в блоке после замены: {t.Text}, len = {len}, newTextChars.Count = {newTextChars.Count}");

                                            //если после замены израсходованы все искомые символы, но остались новые символы, то дописываем их в этот же блок
                                            if (len == 0 && newTextChars.Count > 0)
                                            {
                                                while (newTextChars.Count > 0)
                                                {
                                                    j++; //увеличиваем итерацию символа
                                                    //Debug.WriteLine($"Дописываем символ: newTextChars.Peek = {newTextChars.Peek()}");
                                                    t.Text = t.Text.Insert(j, newTextChars.Dequeue().ToString());
                                                    //Debug.WriteLine($"Символы в блоке после дописывания: {t.Text}, len = {len}, newTextChars.Count = {newTextChars.Count}");
                                                }
                                            }
                                        }

                                        //удаляем символ, если количество искомых символов осталось, но нет новых символов
                                        else if (len > 0 && newTextChars.Count == 0)
                                        {
                                            t.Text = t.Text.Remove(j, 1);
                                            len--; //уменьшаем количество символов, которые нужно заменить, так как записали или удалил символ
                                            j--; //уменьшаем переменную итерации
                                            //Debug.WriteLine($"Символы в блоке после удаления: {t.Text}, len = {len}, newTextChars.Count = {newTextChars.Count}");
                                        }

                                        //это условие избыточно, но может показать, что что-то не так...
                                        else
                                        {
                                            //Debug.WriteLine($"Оставшееся условие: {t.Text}, len = {len}, newTextChars.Count = {newTextChars.Count}");
                                        }
                                    }
                                } //перебор символов 
                            } //перебор блоков с кусками форматированного текстая
                            Debug.WriteLine($"Новая строка si: {siItem.InnerText}");
                            counter++; //увеличиваем подсчёт
                        }//если часть текста всех блоков совпадает с искомой строкой
                    }
                    //excelDoc.Save(); //сохраняет документ excel, но таблица строк сохраняется сама
                }
            }
            catch (FileNotFoundException ex)
            {
                Debug.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
            }
        }

        private void replaceBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.Compare((string)replaceBtn.Content, "Заменить") == 0 && worker.IsBusy != true)
            {
                //File.Copy("bookOld.xlsx", "book.xlsx", true);
                oldTxt = replaceWhat.Text;
                newTxt = replaceWith.Text;
                progress.Value = 0;
                replaceBtn.Content = "Отмена";

                worker.ProgressChanged += worker_ProgressChanged;
                worker.RunWorkerCompleted += worker_RunWorkerCompleted;

                worker.RunWorkerAsync();
            }
            else
            {
                worker.CancelAsync();
            }
        }
        void doWork(object? sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            string[] files = Directory.GetFiles(Environment.CurrentDirectory, "*.xlsx");

            filesNumber = files.Length;

            foreach (string file in files)
            {
                if(worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    //ReplaceSymbolsInSharedStringTable(file, oldTxt, newTxt);                    
                    worker.ReportProgress(100 / filesNumber, file);                    
                    Thread.Sleep(100);
                }
            }
        }
        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progress.Value += 100 / filesNumber;
            progressText.Text = e.UserState.ToString();
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(e.Cancelled == true)
            {

            }
            else if(e.Error != null)
            {

            }
            else
            {
                progress.Value = 100;
                MessageBox.Show("Сделано замен: " + counter + "!");
                replaceBtn.Content = "Заменить";
            }
        }
    }
}