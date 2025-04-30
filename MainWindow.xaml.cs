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

namespace ExcelTextReplacer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            //сначала нужно сделать перебор ячеек на одной странице
            //дальше попробовать замену

            InitializeComponent();

            string path = @"93-24-2030_РКМ_Койда_1_безопасность.xlsx";
            string txt = @"ИТОГО с НДС 0%";

            string? val = GetCellValue(path, "Плановая2", "B57");
            bool res = CheckCellString(path, txt);

            Debug.WriteLine(res);
        }
        bool CheckCellString(string filepath, string str)
        {
            string sheetName = "Плановая2"; //пока ищем только на этом листе
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(filepath, false)) //открываем файл ecxel
                {
                    WorkbookPart? wbPart = excelDoc.WorkbookPart; //получаем часть
                    Sheets? sheets = wbPart?.Workbook.Sheets; //получаем страницы
                    Sheet? sheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name.Value.StartsWith(sheetName)).FirstOrDefault(); //поиск листа по имени
                    WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(sheet?.Id!); //часть листа

                    Cell? cell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == address).FirstOrDefault(); //поиск ячейки по адресу


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
    }
}