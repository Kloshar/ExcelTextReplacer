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
        string path = @"93-24-2030_РКМ_Койда_1_безопасность.xlsx";
        int counter = 0;
        public MainWindow()
        {
            InitializeComponent();
            
            string oldTxt = @"hello";
            string newTxt = @"Good morning!";            

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

                    foreach(SharedStringItem ssItem in ssPart.SharedStringTable)
                    {
                        //Debug.WriteLine(ssItem.InnerText);
                        if (ssItem.InnerText.StartsWith(oldTxt))
                        {
                            string str = ssItem.InnerText;

                            Debug.WriteLine(str.Contains('\n'));

                            Debug.WriteLine($"Заменяем {ssItem.InnerText} на {newTxt}");
                            ssItem.Text = new DocumentFormat.OpenXml.Spreadsheet.Text(newTxt);
                            counter++;
                        }
                    }                    
                    excelDoc.Save();
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
        private void replaceBtn_Click(object sender, RoutedEventArgs e)
        {
            bool res = CheckCellString(path, replaceWhat.Text, replaceWith.Text);

            if (res)
            {
                MessageBox.Show("Сделано замен: " + counter + "!");
            }
        }
    }
}