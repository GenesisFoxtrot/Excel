using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class OptionViewModel
    {
        public String Name { get; set; }

        public string ExcelName { get; set; }

        public int Column { get; set; }

        public OptionViewModel(String name, string excelName)
        {
            Name = name;
            ExcelName = excelName;
        }
    }
}
