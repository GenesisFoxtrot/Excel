using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Model;

namespace ExcelReaderLinq.Excel
{
    public class Reader
    {
        public void PrintArtistAlbums()
        {
            string pathToExcelFile = ""
            + @"D:\1.xlsx";

            string sheetName = "Sheet1";

        

            var excelFile = new LinqToExcel.ExcelQueryFactory(pathToExcelFile);
            excelFile.AddMapping("Name", "Name");
            excelFile.AddMapping("Album", "Album");
            var artistAlbums = from a in excelFile.Worksheet<Model.Model>(sheetName) select a;

            foreach (var a in artistAlbums)
            {
                string artistInfo = "Artist Name: {0}; Album: {1}";
                Console.WriteLine(artistInfo, a.Name, a.Album);
            }
        }
    }
}
