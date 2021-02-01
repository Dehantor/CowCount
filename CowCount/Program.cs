using System;
using System.IO;
using System.Threading.Tasks;
/// <summary>
///1.	Написать приложение, которое при запуске берет указанный  из командной строки файл .xlsx
///2.   В файле есть лист с именем «шаблон»
///3.	Создаем лист в файле с названием текущая дата например 19.12.2020, копируем в него содержимое шаблона.
///4.	Заполняем новый лист цифрами из файлов (брать из папки откуда запустится приложение). Имена файлов соответствуют ячейкам (в том числе и С51, в задании по этой ячейке в скобках пояснения, поэтому прочитав задание надо файл переименовать в С51.csv) 
///5.После выполнения оставить сформированный документ открытым
/// </summary>
namespace CowCount
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo fileName;
            while (true)
            {
                Console.WriteLine("Введите название файлас расширением: ");
                string file = Console.ReadLine();
                fileName = new FileInfo(file);
                if (fileName.Exists)
                    break;
                else
                    Console.WriteLine("Файл не существеует! попробуйте еще раз.");
            }
            CowExcel cowExcel = new CowExcel();
            cowExcel.Start(fileName);
            Console.WriteLine("Готово");
        }
    }
}
