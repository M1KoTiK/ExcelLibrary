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
    public class ExcelWorkspace
    {

        public FileInfo FileLocation { get; private set; }
        
        public ExcelWorkspace(FileInfo fileLocation)
        {
            FileLocation = fileLocation;
        }

        private IEnumerable<XElement> GetXMLSheets()
        {
            using var fileStream = File.Open(FileLocation.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var workbookArchive = archive.GetEntry("xl/workbook.xml");
            if (workbookArchive == null) throw new NullReferenceException("Файл xl/workbook.xml отсутствует");
            var stream = workbookArchive.Open();
            XDocument doc = XDocument.Load(stream);
            stream.Dispose();
            return (from el in doc.Descendants() where el.Name.LocalName == "sheet" select el).AsEnumerable();
        }
        
        public IEnumerable<Sheet> GetSheets()
        {
            var sheets = new List<Sheet>();
            foreach(var el in GetXMLSheets())
            {
                sheets.Add(new Sheet(this, el));
            }
            return sheets;
        }

        public Sheet? SheetByName(string sheetName)
        {
            foreach(var sheet in GetSheets())
            {
                if(sheet.Name == sheetName)
                {
                    return sheet;
                }
            }
            return null;
        }

        



        
    }
}
