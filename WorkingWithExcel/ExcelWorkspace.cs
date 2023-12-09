using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class ExcelWorkspace
    {
        public List<XElement> XMLSheets { get; private set; }
        public List<Sheet> Sheets { get; private set; }

        public FileInfo ExcelFile { get; private set; }
        

        public ExcelWorkspace(FileInfo excelFile)
        {
            ExcelFile = excelFile;
            Update();

        }

        private void Update()
        {
            XMLSheets = GetXMLSheets().ToList();
            Sheets = GetSheetsFromXMLSheets().ToList();

        }


        private IEnumerable<XElement> GetXMLSheets()
        {
            using var fileStream = File.Open(ExcelFile.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var workbookArchive = archive.GetEntry("xl/workbook.xml");
            if (workbookArchive == null) throw new NullReferenceException("Файл xl/workbook.xml отсутствует");
            var stream = workbookArchive.Open();
            XDocument doc = XDocument.Load(stream);
            stream.Dispose();
            return (from el in doc.Descendants() where el.Name.LocalName == "sheet" select el).AsEnumerable();
            
        }

        private IEnumerable<Sheet> GetSheetsFromXMLSheets()
        {
            var sheets = new List<Sheet>();
            foreach(var el in XMLSheets)
            {
                sheets.Add(new Sheet(this, el));
            }
            return sheets;
        }

        public IEnumerable<XElement> GetSharedStringXML()
        {

            using var fileStream = File.Open(ExcelFile.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var workbookArchive = archive.GetEntry("xl/sharedStrings.xml");
            if (workbookArchive == null) throw new NullReferenceException("Файл xl/sharedStrings.xml отсутствует");
            var stream = workbookArchive.Open();
            XDocument doc = XDocument.Load(stream);
            stream.Dispose();
            return (from el in doc.Descendants() where el.Name.LocalName == "si" select el).AsEnumerable();
        }

        public int FindSharedString(string str)
        {
            var xmlSharedStringList = GetSharedStringXML().ToList();
            for (int i = 0; i < xmlSharedStringList.Count(); i++)
            {
                if (xmlSharedStringList[i].Value == str)
                {
                    return i;
                }
            }
            return -1;
        }

        public void AddSharedString(string sharedString)
        {
            if (FindSharedString(sharedString) != -1)
            {
                return;
            }

            using var fileStream = File.Open(ExcelFile.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var workbookArchive = archive.GetEntry("xl/sharedStrings.xml");
            if (workbookArchive == null) throw new NullReferenceException("Файл xl/sharedStrings.xml отсутствует");

            var stream = workbookArchive.Open();
            XDocument doc = XDocument.Load(stream);
            doc.Root.Add(ExcelXML.CreateSharedStringXElement(sharedString, doc.Root.GetDefaultNamespace() ));

            //Очищаем стрим чтобы данные в файле не дублировались
            stream.Position = 0;
            stream.SetLength(0);

            doc.Save(stream);
            stream.Dispose();
            
        }



        
    }
}
