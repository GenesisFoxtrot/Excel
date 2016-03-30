using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using Models;
using System.Drawing;
using System.Linq.Expressions;

namespace ExcelReader2
{
    public class Reader
    {
        delegate void Cast(Model a, string b);
        public List<T> Read<T>(string filePath) where T : new()
        {
            List<OptionViewModel> options = new List<OptionViewModel>();

            var props = typeof(T).GetProperties().ToList(); //Reflection
            props.ForEach(p => options.Add(new OptionViewModel(p, p.Name)));
            OptionsViewModel optionModel = new OptionsViewModel(options);

            Workbook workbook = new Workbook();

            workbook.LoadFromFile(filePath);

            Worksheet sheet = workbook.Worksheets[0];

            var c = sheet.Rows[0].Columns[1].Text;

            var properties = sheet.Range.Rows[0].Select(x => new { x.Text, x.Column }).ToList();
            optionModel.Properties.ForEach(x => x.Column = properties.FirstOrDefault(y => y.Text == x.ExcelName).Column);

            var models = sheet.Rows.ToList().Select(row =>
            {
                var model = new ModelViewModel();
                model.Row = row.Row;
                optionModel.Properties.ForEach(option =>
                {
                    var propery = new Property();
                    propery.Option = option;
                    propery.Value = row.Columns[option.Column - 1].Value;
                    model.Properties.Add(propery);
                });
                return model;
            }
            ).ToList();

            sheet.Range["A1:C1"].Style.Color = Color.SkyBlue;
            sheet.Range["A1:C1"].Comment.RichText.Text = "Some comment";

            workbook.SaveToFile(@"Sample.xlsx");

            List<T> result = models.Select(x =>
            {
                var r = new T();
                x.Properties.ForEach(p => p.Option.Property.SetValue(r, p.Value));
                return r;
            }).ToList();

            return result;
        }



    }
}
