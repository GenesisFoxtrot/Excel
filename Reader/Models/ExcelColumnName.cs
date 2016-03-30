using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader2.Models
{

    [System.AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnName : System.Attribute
    {
        public string Name;
        public ExcelColumnName(string name)
        {
            Name = name;
        }
    }
}
