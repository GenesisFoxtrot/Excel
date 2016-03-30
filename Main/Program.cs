using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader2;
using Models;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {

            List<OptionViewModel> options = new List<OptionViewModel>();
            options.Add(new OptionViewModel("Name", "Name"));
            options.Add(new OptionViewModel("Album", "Album"));
            OptionsViewModel optionModel = new OptionsViewModel(options);

            Reader reader = new Reader();
            var vms = reader.Read(optionModel);

            List<Model> models = vms.Select(x => new Model() //We can use reflection for mapping here 
            {
                Album = x.Properties.FirstOrDefault(y => y.Option.Name == "Album").Value,
                Name = x.Properties.FirstOrDefault(y => y.Option.Name == "Name").Value
            }).ToList();
            Console.ReadLine();
        }
    }
}
