using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.IO.Enumeration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class ExcelFile
    {
        public FileInfo FileLocation { get; private set; }
        public ExcelDocumentInfo DocInfo { get; private set; }
        public IEnumerable<Sheet> Sheets { get; private set; }
        public ExcelFile(FileInfo fileLocation)
        {
            FileLocation = fileLocation;
            DocInfo = new ExcelDocumentInfo(FileLocation);
        }
        public void AddSheet(string name)
        {
            throw new NotImplementedException();
        }
        public void RemoveSheet(string name) 
        {
            throw new NotImplementedException();
        }
        
        private IEnumerable<Sheet> GetSheets()
        {
            throw new NotImplementedException();
            //var sheets = new List<Sheet>();
            //foreach (var el in GetXMLSheets())
            //{
            //    sheets.Add(new Sheet(this, el));
            //}
            //return sheets;
        }  
    }
}
