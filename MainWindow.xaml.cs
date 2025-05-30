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
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Collections.Specialized;
//https://learn.microsoft.com/ru-ru/office/open-xml/spreadsheet/overview

namespace ExcelTextReplacer
{
    public class fileObject //класс для получения имени файлов
    {
        public fileObject(string filePath) //конструктор
        {
            FullPath = filePath; //устанавливаем свойство
            Name = Path.GetFileName(filePath); //устанавливаем свойство
        }
        public string FullPath { get; set; } //свойство FullPath
        public string Name { get; } //свойство Name
        public override string ToString() //переопределяем, чтобы можно было отобразить имя в списке
        {
            return Name;
        }
    }
    public partial class MainWindow : Window
    {
        string path = @"book.xlsx";
        //string path = @"93-24-2030_РКМ_Койда_1_безопасность.xlsx";
        int counter = 0;
        string oldTxt = "";
        string newTxt = "";
        int filesNumber = 0;
        string[] filesInCurrentFolder;
        public ObservableCollection<fileObject> files { get; set; } //свойство класса MainWindow. В нём располагается коллекция объектов типа 'fileObject'

        BackgroundWorker worker;

        public MainWindow()
        {
            files = new ObservableCollection<fileObject>(); //коллекция, посылающая уведомления об изменении

            filesInCurrentFolder = Directory.GetFiles(Environment.CurrentDirectory, "*.xlsx");
            foreach (string f in filesInCurrentFolder) files.Add(new fileObject(f));

            InitializeComponent();
            DataContext = this;

            //files.CollectionChanged += (object? sender, NotifyCollectionChangedEventArgs e) => Debug.WriteLine($"Коллекция изменилась: {e.Action}");

            replaceWhat.Text = oldTxt;
            replaceWith.Text = newTxt;

            worker = new BackgroundWorker(); //класс для асинхронного выполнения в отдельном потоке
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += doWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;

            //replaceBtn.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent)); //автоматическое нажатие кнопки начала замены
        }

        string[] makeFilesList(string path)
        {
            string[] files = Directory.GetFiles(Environment.CurrentDirectory, "*.xlsx");
            return files;
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
        void replaceBtn_Click(object sender, RoutedEventArgs e)
        {
            //File.Copy("bookOld.xlsx", "book.xlsx", true);
            oldTxt = replaceWhat.Text;
            newTxt = replaceWith.Text;
            progress.Value = 0;

            if (string.Compare((string)replaceBtn.Content, "Заменить") == 0 && worker.IsBusy != true)
            {
                replaceBtn.Content = "Отмена";
                worker.RunWorkerAsync();
            }
            else
            {
                if(worker.IsBusy == true) worker.CancelAsync();
            }
        }
        void doWork(object? sender, DoWorkEventArgs e)
        {
            BackgroundWorker? worker = sender as BackgroundWorker;

            filesNumber = files.Count;

            foreach (fileObject file in files)
            {
                //Debug.WriteLine($"worker.CancellationPending = {worker.CancellationPending}");

                if (worker.CancellationPending == true)
                {                    
                    e.Cancel = true;
                    break;
                }
                else
                {
                    //ReplaceSymbolsInSharedStringTable(file.path, oldTxt, newTxt);
                    worker.ReportProgress(100 / filesNumber, file);                    
                    Thread.Sleep(500);
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
            Debug.WriteLine("In worker_RunWorkerCompleted...");
            if(e.Cancelled == true)
            {
                MessageBox.Show("Прервано пользователем.\nСделано замен: " + counter + "!");
            }
            else if(e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else
            {
                progress.Value = 100;
                MessageBox.Show("Сделано замен: " + counter + "!");
            }
            replaceBtn.Content = "Заменить";
            progressText.Text = "";

        }

        void addFileBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dia = new OpenFileDialog();
            dia.Multiselect = true;
            if(dia.ShowDialog() == true)
            {
                foreach (string f in dia.FileNames) files.Add(new fileObject(f));
            }
        }
        void removeFileBtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in lstView.SelectedItems) Debug.WriteLine(s);

            Debug.WriteLine($"{lstView.SelectedItems}");


        }
    }
}