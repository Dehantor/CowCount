using System;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;

namespace CowCount
{
    class CowExcel
    {
        /// <summary>
        /// Агрегация данных с csv
        /// </summary>
        /// <param name="file"></param>
        public void Start(FileInfo file)
        {
            try
            {
                //проблема с кодировкой
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                //открывем ексель куда будем сохранять все данные
                using (var package = new ExcelPackage(file))
                {
                    ExcelWorksheet WS;
                    //копируем данные с шаблона
                    try
                    {
                        using (var temporary = new ExcelPackage(new FileInfo("Шаблон.xlsx")))
                        {
                            WS = package.Workbook.Worksheets.Add(DateTime.Now.ToString("dd.MM.yyyy"), temporary.Workbook.Worksheets["шаблон"]);
                        }
                    }
                    catch
                    {
                        Console.WriteLine("Лист с таками именем уже существует");
                        return;
                    }
                    //проходимся по всем файлам CSV и копируем данные из них
                    string direct = Environment.CurrentDirectory;
                    DirectoryInfo directoryInfo = new DirectoryInfo(direct);
                    foreach (FileInfo item in directoryInfo.GetFiles())
                    {
                        if (item.Name.Contains(".CSV"))
                        {
                            int qual;
                            var lines = File.ReadAllLines(item.Name, Encoding.GetEncoding("windows-1251"));
                            if (lines[0].Split(";")[1] == "\"р  ДСП\"")
                                qual = Convert.ToInt32(lines[1].Split(";")[1]);
                            else
                                qual = Convert.ToInt32(lines[1].Split(";")[0]);
                            string cell = item.Name.Replace(".CSV", "");
                            WS.Cells[cell].Value = qual;
                        }
                    }
                    //сохраняем файл
                    package.Save();
                    //открываем ексель файл
                    var proc = new Process();
                    proc.StartInfo = new ProcessStartInfo(file.Name)
                    {
                        UseShellExecute = true
                    };
                    proc.Start();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
    }
}
