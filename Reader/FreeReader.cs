using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using ExcelReader2.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;


namespace ExcelReader2
{
    public class FreeReader : IReader
    {
        public List<T> Read<T>(string filePath) where T : new()
        {

            var result = new List<T>();
            ExcelService service = new ExcelService();

            OptionsViewModel optionModel = service.GetOptions<T>();


            MemoryStream ms = new MemoryStream();

            try
            {
                using (SLDocument sl = new SLDocument(filePath))
                {
                    var ingexes = sl.GetCells().ToList().Select(x => x.Key).Where(x => x.RowIndex == 1).ToList();
                    var properties =
                        ingexes.Select(
                            x => new {Text = sl.GetCellValueAsString(x.RowIndex, x.ColumnIndex), Column = x.ColumnIndex})
                            .ToList();
                    optionModel.Properties.ForEach(
                        x => x.Column = properties.FirstOrDefault(y => y.Text == x.ExcelName).Column);

                    var rows = sl.GetCells().ToList().Select(x => x.Key).GroupBy(c=>c.RowIndex).ToList();

                    var models = rows.Where(x => x.Key > 1).ToList().Select(row =>
                    {
                        var model = new ModelViewModel();
                        model.Row = row.Key;
                        optionModel.Properties.ForEach(option =>
                        {
                            var propery = new Property();
                            propery.Option = option;
                            propery.Value = sl.GetCellValueAsString(row.Key, option.Column);
                            model.Properties.Add(propery);
                        });
                        return model;
                    }).ToList();

                    result = models.Select(service.MapModel<T>).ToList();


                    var comm = sl.CreateComment(); //Sample
                    comm.SetText("Some comment");  //Sample
                    sl.InsertComment(3, 1, comm);  //Sample

                    var style = new SLStyle();              //Sample
                    style.Fill.SetPattern(PatternValues.DarkGrid,  System.Drawing.Color.Red,
                        System.Drawing.Color.Red);   //Sample
                    sl.SetCellStyle(3,1, style);           //Sample

                    sl.SaveAs(@"Sample.xlsx");            //Sample 

                }
                ms.Position = 0;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return result;
        }
    }
}
