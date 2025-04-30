using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text.RegularExpressions;

//https://learn.microsoft.com/ru-ru/office/open-xml/spreadsheet/overview

namespace ExcelTextReplacer
{
    public partial class MainWindow //вынесено в часть класса для исключения конфликта имён между Spreadsheet и Wordprocessing
    {
        static string? SetCellValue(string filepath, string sheetName, string address, string newVal)
        {
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(filepath, true)) //открываем файл ecxel
                {
                    excelDoc.Save();

                    WorkbookPart? wbPart = excelDoc.WorkbookPart; //получаем часть
                    Sheets? sheets = wbPart?.Workbook.Sheets; //получаем страницы
                    Sheet? sheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name.Value.StartsWith(sheetName)).FirstOrDefault(); //поиск листа по имени

                    //Sheet? sheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault(); //поиск листа по имени
                    WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(sheet?.Id!);
                    Cell? cell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == address).FirstOrDefault(); //поиск ячейки по адресу
                    string? newValue = string.Empty; //для возвращаемого значения

                    //Console.WriteLine(cell.InnerXml);

                    //последовательно проверяем существует ли ячека, есть ли у неё тип данных, если есть, то какой именно
                    if (cell is null || cell.InnerText.Length <= 0) //обработка, если ячейка не найдена
                    {
                        Console.WriteLine($"Ячейка не найдена или её значение имеет длину меньше нуля!");
                        return string.Empty;
                    }

                    if (cell.DataType != null) //если тип данных не равен нулю, значит там строка или булево значение
                    {
                        if (cell.DataType == CellValues.SharedString) //если тип данных - строка
                        {
                            //Console.WriteLine(cell?.OuterXml);
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            newValue = stringTable?.SharedStringTable.ElementAt(int.Parse(cell?.CellValue?.InnerText)).InnerText;
                            //Console.WriteLine(newValue);
                            return newValue;
                        }
                        else if (cell.DataType == CellValues.Boolean) //если тип данных - булево значение
                        {
                            newValue = cell?.CellValue?.InnerText == "1" ? "true" : "false";
                            //Console.WriteLine(newValue);
                            return newValue;
                        }
                        else //другие случаи DataType, мало ли
                        {
                            newValue = cell?.CellValue?.InnerText;
                            return newValue;
                        }
                    }
                    else //в остальных случаях - число
                    {
                        //newValue = PrecisionAdjusting(cell?.CellValue?.InnerText);
                        //Console.WriteLine($"New value: {newValue}");

                        //Console.WriteLine($"Текущее значение: {cell.CellValue.Text}");

                        cell.CellValue = new CellValue(newVal);

                        //Console.WriteLine($"Новое значение: {cell.CellValue.Text}");

                        return newValue.ToString();
                        //return cell.CellValue.InnerText; //заглушка
                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return string.Empty;
            }

        } //end SetCellValue
        static string? DelCellValue(string filepath, string sheetName, string address)
        {
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(filepath, true)) //открываем файл ecxel
                {
                    WorkbookPart? wbPart = excelDoc.WorkbookPart; //получаем часть
                    Sheets? sheets = wbPart?.Workbook.Sheets; //получаем страницы
                    Sheet? sheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name.Value.StartsWith(sheetName)).FirstOrDefault(); //поиск листа по имени

                    //Console.WriteLine($"Страница; {sheet.Name}, шаблон: {sheetName}");

                    //Sheet? plan2 = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheet).FirstOrDefault(); //поиск листа по имени
                    WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(sheet?.Id!);
                    Cell? cell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == address).FirstOrDefault(); //поиск ячейки по адресу
                    string? newValue = string.Empty; //для возвращаемого значения

                    //последовательно проверяем существует ли ячека, есть ли у неё тип данных, если есть, то какой именно
                    if (cell is null || cell.InnerText.Length <= 0) //обработка, если ячейка не найдена
                    {
                        Console.WriteLine($"Ячейка не найдена или её значение имеет длину меньше нуля!");
                        return string.Empty;
                    }

                    if (cell.DataType != null) //если тип данных не равен нулю, значит там строка или булево значение
                    {
                        if (cell.DataType == CellValues.SharedString) //если тип данных - строка
                        {
                            //Console.WriteLine(cell?.OuterXml);
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            newValue = stringTable?.SharedStringTable.ElementAt(int.Parse(cell?.CellValue?.InnerText)).InnerText;
                            //Console.WriteLine(newValue);
                            return newValue;
                        }
                        else if (cell.DataType == CellValues.Boolean) //если тип данных - булево значение
                        {
                            newValue = cell?.CellValue?.InnerText == "1" ? "true" : "false";
                            //Console.WriteLine(newValue);
                            return newValue;
                        }
                        else //другие случаи DataType, мало ли
                        {
                            newValue = cell?.CellValue?.InnerText;
                            return newValue;
                        }
                    }
                    else //в остальных случаях - число
                    {
                        //newValue = PrecisionAdjusting(cell?.CellValue?.InnerText);
                        //Console.WriteLine($"New value: {newValue}");
                        if (cell.CellValue != null)
                        {
                            cell.CellValue.Remove();
                        }

                        return newValue.ToString();
                        //return cell.CellValue.InnerText; //заглушка
                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return string.Empty;
            }

        } //end DelCellValue
        string? GetCellValue(string filepath, string sheetName, string address)
        {
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(filepath, false)) //открываем файл ecxel
                {
                    WorkbookPart? wbPart = excelDoc.WorkbookPart; //получаем часть
                    Sheets? sheets = wbPart?.Workbook.Sheets; //получаем страницы
                    Sheet? sheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name.Value.StartsWith(sheetName)).FirstOrDefault(); //поиск листа по имени

                    //foreach (Sheet s in sheets)
                    //{
                    //    Console.WriteLine($"{s.Name.Value}, {sheetName}, {s.Name.Value.StartsWith(sheetName)}");
                    //}

                    //Sheet? sheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault(); //поиск листа по имени
                    WorksheetPart wsPart = (WorksheetPart)wbPart!.GetPartById(sheet?.Id!);
                    Cell? cell = wsPart.Worksheet?.Descendants<Cell>()?.Where(c => c.CellReference == address).FirstOrDefault(); //поиск ячейки по адресу
                    string? newValue = string.Empty; //для возвращаемого значения

                    //последовательно проверяем существует ли ячека, есть ли у неё тип данных, если есть, то какой именно
                    if (cell is null || cell.InnerText.Length <= 0) //обработка, если ячейка не найдена
                    {
                        Console.WriteLine($"Ячейка не найдена или её значение имеет длину меньше нуля!");
                        return string.Empty;
                    }

                    if (cell.DataType != null) //если тип данных не равен нулю, значит там строка или булево значение
                    {
                        if (cell.DataType == CellValues.SharedString) //если тип данных - строка
                        {
                            //Console.WriteLine(cell?.OuterXml);
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            newValue = stringTable?.SharedStringTable.ElementAt(int.Parse(cell?.CellValue?.InnerText)).InnerText;
                            //Console.WriteLine(newValue);
                            return newValue;
                        }
                        else if (cell.DataType == CellValues.Boolean) //если тип данных - булево значение
                        {
                            newValue = cell?.CellValue?.InnerText == "1" ? "true" : "false";
                            //Console.WriteLine(newValue);
                            return newValue;
                        }
                        else //другие случаи DataType, мало ли
                        {
                            newValue = cell?.CellValue?.InnerText;
                            return newValue;
                        }
                    }
                    else //в остальных случаях - число
                    {
                        newValue = PrecisionAdjusting(cell?.CellValue?.InnerText);
                        //Console.WriteLine($"New value: {newValue}");
                        return newValue.ToString();
                        //return cell.CellValue.InnerText; //заглушка
                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return string.Empty;
            }

        } //end GetCellValue
        static string PrecisionAdjusting(string value)
        {
            string oldValue = value;

            Regex regInteger = new Regex(@"^\d+$"); //шаблон для целых чисел
            Regex regDouble = new Regex(@"^(\d*)(\.|\,){1}(\d*)$"); //шаблон для числа с точкой

            //Console.WriteLine(oldValue);

            if (regInteger.IsMatch(oldValue)) //сначала проверяем на целое число по шаблону @"^\d+$"
            {
                //Console.WriteLine($"Целое: {oldValue}");
                return oldValue;
            }
            else if (regDouble.IsMatch(oldValue)) //далее проверяем на соответствие десятичному числу по шаблону @"^(\d*)(\.|\,){1}(\d*)$"
            {
                double doubleValue = Convert.ToDouble(oldValue.Replace('.', ',')); //заменяем точеку на запятую и конвертим в десятичное
                doubleValue = Math.Round(doubleValue, 2);
                string newValue = doubleValue.ToString("#,###.##");
                //Console.WriteLine($"Десятичное: {newValue}");
                return newValue;
            }
            else //остальные значения считаем строками (не должно быть, на всякий случай)
            {
                //Console.WriteLine($"Строка: {oldValue}");
                return oldValue; //если строка
            }
        } //округление до двух знаков и отделение пробелом
    }
}