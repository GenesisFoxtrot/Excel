using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader2.Models;

namespace Models
{
    public class Model
    {
        public string Name { get; set; }

        [ExcelColumnName("Album")]
        public string Album2 { get; set; }
    }
}
