using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader2.Models
{
    public class OptionViewModel
    {
        public PropertyInfo Property { get; set; }

        public string ExcelName { get; set; }

        public int Column { get; set; }

        public OptionViewModel(PropertyInfo property, string excelName)
        {
            Property = property;
            ExcelName = excelName;
        }
    }
}
