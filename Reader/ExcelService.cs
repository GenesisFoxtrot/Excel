using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader2.Models;

namespace ExcelReader2
{
    public class ExcelService
    {
        public OptionsViewModel GetOptions<T>()
        {
            List<OptionViewModel> options = new List<OptionViewModel>();

            var props = typeof(T).GetProperties().ToList();


            props.ForEach(p =>
            {
                var columName = (p.GetCustomAttributes(false).FirstOrDefault(x => x is ExcelColumnName) as ExcelColumnName)?.Name;
                options.Add(new OptionViewModel(p, columName ?? p.Name));
            });
            OptionsViewModel optionModel = new OptionsViewModel(options);

            return optionModel;
        }

        public T MapModel<T>(ModelViewModel vm) where T : new()
        {
            var result = new T();
            vm.Properties.ForEach(p => p.Option.Property.SetValue(result, p.Value));
            return result;

        }
        
    }
}
