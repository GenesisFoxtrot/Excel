using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader2.Models
{
    public class OptionsViewModel
    {
        public OptionsViewModel(List<OptionViewModel> properties)
        {
            Properties = properties.Select(x => new OptionViewModel(x.Property, x.ExcelName)).ToList();
        }

        public List<OptionViewModel> Properties; 
    }
}
