using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader2.Models
{
    public class ModelViewModel
    {
        public List<Property> Properties;
        public int Row { get; set; }

        public ModelViewModel()
        {
            Properties = new List<Property>();
        }
    }
}
