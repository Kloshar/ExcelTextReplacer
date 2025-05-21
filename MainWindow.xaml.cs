using System;
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
using Run = DocumentFormat.OpenXml.Spreadsheet.Run;
using System.Windows.Controls.Primitives;

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
    public partial class MainWindow : Window
    {
        string path = @"book.xlsx";
        //string path = @"93-24-2030_РКМ_Койда_1_безопасность.xlsx";
        int index = 0; //курсор искомой строки
        int writed = 0; //посчёт записанных символов (новых)

        int counter = 0;
        string oldTxt = "bc";
        string newTxt = "new";
        bool usedUp = false;
        int option = 1;
        public MainWindow()
        {
            InitializeComponent();

            replaceWhat.Text = oldTxt;
            replaceWith.Text = newTxt;

            replaceBtn.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent)); //автоматическое нажатие кнопки начала замены
        }

        void StartWorking(string filepath, string oldTxt, string newTxt)
        {
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(filepath, true)) //открываем файл ecxel
                {
                    WorkbookPart? wbPart = excelDoc.WorkbookPart; //получаем часть книги
                    SharedStringTablePart? ssPart = wbPart.SharedStringTablePart; //далее получаем все узлы на листе

                    Debug.WriteLine(ssPart.SharedStringTable.OuterXml); //выводим xml узлов

                    foreach (SharedStringItem siItem in ssPart.SharedStringTable) //перебираем строки в таблице строк
                    {
                        if (siItem.InnerText.Contains(oldTxt) && option == 2) //если часть текста элемента si совпадает с искомой строкой
                        {
                            string txt = siItem.InnerText; //записываем текст из строки
                            int start = txt.IndexOf(oldTxt); //получаем индекс начала найденной строки
                            int len = oldTxt.Length; //количество заменяемых (в том числе удаляемых) символов
                            Queue<char> newTextChars = new Queue<char>(); //очередь символов из новой строки для замены
                            foreach(char c in newTxt) newTextChars.Enqueue(c); //добавляем новые символы в очередь
                                 
                            for (int i = 0; i < siItem.Count(); i++) //перебор блоков с кусками форматированного текста
                            {
                                DocumentFormat.OpenXml.OpenXmlElement el = siItem.ChildElements[i]; //один из блоков с символами

                                Debug.WriteLine($"Работаем с элементом: {el.InnerXml}");

                                Text? t = null; //переменная для хранения текста блока
                                if (el is Text) t = (Text?)el; //если это текст без форматирования
                                if (el is Run) t = el.Descendants<Text>().FirstOrDefault(); //если это прогон, получаем потомка типа текст

                                for (int j = 0; j < t.Text.Length; j++) //далее перебираем символы пока индекс не уменьшится до нуля
                                {
                                    if (start > 0) //если перебрали меньше символов, чем было в искомой строке
                                    {
                                        start--; //то уменьшаем на один
                                        Debug.WriteLine($"Обрабатывается символ: {t.Text[j]}, start = {start}");
                                    }
                                    else
                                    {
                                        Debug.WriteLine($"Обрабатывается символ: {t.Text[j]}, искомая строка = {oldTxt}");

                                        //заменяем символ, если количество заменяемых символов осталось и есть чем заменять
                                        if (len > 0 && newTextChars.Count > 0)
                                        {
                                            t.Text = t.Text.Remove(j, 1).Insert(j, newTextChars.Dequeue().ToString());
                                            len--; //уменьшаем количество символов, которые нужно заменить, так как записали или заменили символ
                                            Debug.WriteLine($"Символы в блоке после замены: {t.Text}, len = {len}");
                                            
                                        }

                                        else if (len > 0 && newTextChars.Count == 0) //нужно удалить символ
                                        {
                                            t.Text = t.Text.Remove(j, 1);                                            
                                            len--; //уменьшаем количество символов, которые нужно заменить, так как записали или удалил символ
                                            Debug.WriteLine($"Символы в блоке после удаления: {t.Text}, len = {len}");
                                        }

                                        else if (len == 0 && newTextChars.Count > 0) //нужно дописать символ в тот же блок
                                        {
                                            t.Text = t.Text.Insert(j, newTextChars.Dequeue().ToString());
                                            Debug.WriteLine($"Символы в блоке после дописывания: {t.Text}, len = {len}");
                                        }
                                    }
                                } //перебор символов 
                            } //перебор блоков с кусками форматированного текстая
                            Debug.WriteLine(siItem.InnerText);
                        }//если часть текста всех блоков совпадает с искомой строкой
                    }
                    //excelDoc.Save(); //сохраняет документ excel, но таблица строк сохраняется сама
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }        

        private void replaceBtn_Click(object sender, RoutedEventArgs e)
        {
            File.Copy("bookOld.xlsx", "book.xlsx", true);

            oldTxt = replaceWhat.Text;
            newTxt = replaceWith.Text;

            StartWorking(path, oldTxt, newTxt);

            //if (res)
            //{
            //    Debug.WriteLine("Сделано замен: " + counter + "!");
            //    MessageBox.Show("Сделано замен: " + counter + "!");
            //}
        }
    }
}