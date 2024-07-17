using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class ExcelfileLocation
    {
        public FileInfo OriginalFileLocation { get; private set; }
        
        public ExcelfileLocation(FileInfo fileLocation)
        {
            OriginalFileLocation = fileLocation;
        }
        public ExcelFile Open()
        {
            return new ExcelFile(OriginalFileLocation);
        }
        public ExcelFile OpenDuplicate(FileInfo duplicateLocation)
        {
            return new ExcelFile(duplicateLocation);
        }

        public ExcelFile Create() 
        { 
            throw new NotImplementedException();
        }     
    }
}
