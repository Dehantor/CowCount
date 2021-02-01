using System;
using System.IO;
using System.Threading.Tasks;

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
