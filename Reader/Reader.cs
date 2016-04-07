using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using System.Drawing;
using ExcelReader2.Models;
using System.Linq.Expressions;

namespace ExcelReader2
{
    public class Reader : IReader
    {
        public List<T> Read<T>(string filePath) where T : new()
        {

            var result = new List<T>();
            ExcelService service = new ExcelService();
   
            OptionsViewModel optionModel = service.GetOptions<T>();

            try
            {
                using (Workbook workbook = new Workbook())
                {
                    workbook.LoadFromFile(filePath);

                    using (Worksheet sheet = workbook.Worksheets[0])
                    {

                        var c = sheet.Rows[0].Columns[1].Text;

                        var properties = sheet.Rows[0].Select(x => new { x.Text, Column = x.Column -1 }).ToList();
                        optionModel.Properties.ForEach(x => x.Column = properties.FirstOrDefault(y => y.Text == x.ExcelName).Column);

                        var models = sheet.Rows.Where(x => x.Row > 1).ToList().Select(row =>
                        {
                            var model = new ModelViewModel();
                            model.Row = row.Row;
                            optionModel.Properties.ForEach(option =>
                            {
                                var propery = new Property();
                                propery.Option = option;
                                propery.Value = row.Columns[option.Column].Value;
                                model.Properties.Add(propery);
                            });
                            return model;
                        }).ToList();
                         
                        sheet.Range["A1:C1"].Style.Color = Color.SkyBlue;              //Sample
                        sheet.Range["A1:C1"].Comment.RichText.Text = "Some comment";   //Sample
                        workbook.SaveToFile(@"Sample.xlsx");                           //Sample

                        result = models.Select(service.MapModel<T>).ToList();
                    }
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return result;
        }
    }
}
