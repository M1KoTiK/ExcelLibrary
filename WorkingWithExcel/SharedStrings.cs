using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace WorkingWithExcel
{
    public class SharedStrings
    {
        public int Count { get; private set; }
        public int UniqueCount { get; private set; }
        public List<string> SharedString { get; private set; }
        public SharedStrings(ExcelWorkspace workspace) 
        {

        }
        
    }
}
