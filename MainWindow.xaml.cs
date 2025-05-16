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
        int substituted = 0; //подсчёт замешённых символов (старых)
        int counter = 0;
        string oldTxt = "qw";
        string newTxt = "r";
        bool usedUp = false;
        int option = 1;
        public MainWindow()
        {
            InitializeComponent();

            replaceWhat.Text = oldTxt;
            replaceWith.Text = newTxt;

            replaceBtn.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent)); //автоматическое нажатие кнопки начала замены
        }
        bool CheckCellString(string filepath, string oldTxt, string newTxt, int option)
        {
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(filepath, true)) //открываем файл ecxel
                {
                    WorkbookPart? wbPart = excelDoc.WorkbookPart; //получаем часть                    
                    SharedStringTablePart? ssPart = wbPart.SharedStringTablePart; //далее получаем все узлы на листе

                    //Debug.WriteLine(ssPart.SharedStringTable.OuterXml);

                    foreach (SharedStringItem ssItem in ssPart.SharedStringTable) //перебираем строки (элементы) в таблице строк
                    {                        
                        if (ssItem.InnerText == oldTxt && option == 1) //если текст всех блоков совпадает с искомой строкой
                        {
                            for(int i = 0; i < ssItem.Count(); i++) //перебор блоков с кусками форматированного текста
                            {
                                DocumentFormat.OpenXml.OpenXmlElement el = ssItem.ChildElements[i]; //один из потомков

                                Debug.WriteLine($"Работаем с элементом: {el.InnerXml}");

                                Text? t = null;
                                if (el is Text) t = (Text?)el; //если это текст без форматирования
                                if (el is Run) t = el.Descendants<Text>().FirstOrDefault(); //если это прогон, получаем потомка типа текст

                                if (usedUp == false) //если не все символы новой строки израсходованы
                                {
                                    Debug.WriteLine($"Заменяем текст в элементе: {el.InnerXml}");
                                    replaceTextInBlock(t);
                                }
                                else
                                {
                                    Debug.WriteLine($"Удаляем элемент: {el.InnerXml}");
                                    el.Remove(); //удаляем оставшиеся блоки с текстом и форматированием
                                    i--;
                                }
                            }
                        }//если текст всех блоков совпадает с искомой строкой
                        if (ssItem.InnerText.Contains(oldTxt) && option == 2) //если часть текста всех блоков совпадает с искомой строкой
                        {
                            string txt = ssItem.InnerText;
                            int start = txt.IndexOf(oldTxt); //start = 3

                            for (int i = 0; i < ssItem.Count(); i++) //перебор блоков с кусками форматированного текста
                            {
                                DocumentFormat.OpenXml.OpenXmlElement el = ssItem.ChildElements[i]; //один из потомков

                                //Debug.WriteLine($"Работаем с элементом: {el.InnerXml}");

                                Text? t = null;
                                if (el is Text) t = (Text?)el; //если это текст без форматирования
                                if (el is Run) t = el.Descendants<Text>().FirstOrDefault(); //если это прогон, получаем потомка типа текст
                                
                                for (int j = 0; j < t.Text.Length; j++) //далее перебираем символы пока индекс не уменьшится до нуля
                                {
                                    if(start > 0) //если перебрали меньше символов, чем было в искомой строке
                                    {
                                        start--; //то уменьшаем на один
                                        //Debug.WriteLine($"Обрабатывается символ: {t.Text[j]}, start = {start}"); //xyzqwe
                                    }
                                    else
                                    {
                                        if (usedUp == false) //если не все символы новой строки израсходованы
                                        {
                                            Debug.WriteLine($"Заменяем текст в элементе: {el.InnerXml}");
                                            replaceTextInBlock(t);
                                        }
                                        else
                                        {
                                            //в этом варианте не удаляем оставшиеся блоки с текстом
                                        }
                                    }
                                }
                            }
                        }//если часть текста всех блоков совпадает с искомой строкой
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
        void replaceTextInBlock(Text? t)
        {  
            string txt = t.Text;
            if (txt.Length >= oldTxt.Length) //если все искомые символы в этом элементе
            {
                Debug.WriteLine($"Все искомые символы с этом теге!");
                t.Text = newTxt; //то просто заменяем текст в элементе на новую строку
                index += txt.Length; //продвигаем индекс на количество символов в элементе
                writed += index; //сохраняем число записанных символов
                if (writed >= newTxt.Length) usedUp = true; //если записаны все символы новой строки
                substituted += txt.Length; //сколько символов перезаписано
            }
            if (txt.Length < oldTxt.Length) //если в элементе только часть искомых символов //q.Length < r.Length
            {
                Debug.WriteLine($"index = {index}");

                //если осталось заменить больше или равное, чем осталось в новой строке, то продвигаем на оставшееся количество символов или продвигаем индекс на количество заменяемых символов
                index += txt.Length >= newTxt.Substring(writed).Length ? newTxt.Substring(writed).Length : txt.Length; //

                //Debug.WriteLine($"{newTxt.Substring(writed, index - writed)}");

                t.Text = newTxt.Substring(writed, index - writed); //то перезаписываем текст в элементе частью новой строки
                writed += index; //сохраняем число записанных символов
                if (writed >= newTxt.Length) usedUp = true; //если записаны все символы новой строки
                //substituted += 
            }
            counter++; //обновляем счётчик
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
            oldTxt = replaceWhat.Text;
            newTxt = replaceWith.Text;

            option = (bool)optionsBtn01.IsChecked ? 1 : 2;
            option = (bool)optionsBtn02.IsChecked ? 2 : 1;

            bool res = CheckCellString(path, replaceWhat.Text, replaceWith.Text, option);

            if (res)
            {
                Debug.WriteLine("Сделано замен: " + counter + "!");
                //MessageBox.Show("Сделано замен: " + counter + "!");
            }
            File.Copy("bookOld.xlsx", "book.xlsx", true);
        }
    }
}