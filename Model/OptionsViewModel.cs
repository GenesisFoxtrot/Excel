using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class OptionsViewModel
    {
        public OptionsViewModel(List<OptionViewModel> properties)
        {
            Properties = properties.Select(x => new OptionViewModel(x.Name, x.ExcelName)).ToList();
        }

        public List<OptionViewModel> Properties; 
    }
}
