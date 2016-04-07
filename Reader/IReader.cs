using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader2
{
    public interface IReader
    {
       List<T> Read<T>(string filePath) where T : new();
    }
}
