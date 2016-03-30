using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader2;
using Models;
using System.Linq.Expressions;

namespace ConsoleApp
{
    class Program
    {

        static void Main(string[] args)
        {
            Reader reader = new Reader();
            var models = reader.Read<Model>(@"1.xlsx");
            models.ForEach(x => Console.WriteLine(x.Name + "    " +x.Album2));
            Console.ReadLine();
        }
    }
}
