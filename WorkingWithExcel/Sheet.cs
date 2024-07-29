using System;
using System.Collections.Generic;
using System.Data;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class Sheet
    {
        private string name;
        public string Name
        {
            get => name;
            set
            {
                name = value;
                SetSheetName(value);
            }
        }

        public string SheetId { get; private set; }
        public string RId { get; private set; }
        public string SheetFileName { get; private set; }

        FileInfo fileLocation;

        private XElement xmlSheetElement;
        
        private Sheet(FileInfo fileLocation, XElement xmlSheetElement)
        {
            this.xmlSheetElement = xmlSheetElement;
            this.fileLocation = fileLocation;
            var nameAttr = xmlSheetElement.Attribute("name");

            if (nameAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут name");
            name = nameAttr.Value;

            var sheetIdAttr = xmlSheetElement.Attribute("sheetId");
            if(sheetIdAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут sheetId");
            SheetId = sheetIdAttr.Value;

            var ns = xmlSheetElement.GetNamespaceOfPrefix("r");
            var rIdAttr = xmlSheetElement.Attribute(ns + "id");
            if (rIdAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут r:id");
            RId = rIdAttr.Value;

            using var fileStream = File.Open(fileLocation.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var relationshipArchive = archive.GetEntry("xl/_rels/workbook.xml.rels");
            if (relationshipArchive == null) throw new NullReferenceException("Файл workbook.xml.rels отсутствует");
            var stream = relationshipArchive.Open();
            XDocument doc = XDocument.Load(stream);

            var relationshipsElements = (from el in doc.Descendants() where el.Name.LocalName == "Relationships" select el).FirstOrDefault();
            if( relationshipsElements == null) throw new NullReferenceException("при попытике спарсить workbook.xml.rels не найден атрибут Relationships");
            stream.Dispose();

            var relationshipElement = (from el  in relationshipsElements.Descendants() where el.Attribute("Id")?.Value == RId select el).FirstOrDefault();
            if (relationshipElement == null) throw new NullReferenceException("При попытки парсить workbook.xml.rels не найден элемент соответствующий данной Sheet");

            var targetAttr = relationshipElement.Attribute("Target");
            if(targetAttr == null) throw new NullReferenceException("При попытки парсить workbook.xml.rels relationship (sheet) не найден атрибут Target");
            SheetFileName = new FileInfo(targetAttr.Value).Name;
        }
        public static IEnumerable<Sheet> GetSheets(FileInfo excelFileLocation)
        {
            var sheets = new List<Sheet>();
            foreach (var el in GetXMLSheets(excelFileLocation))
            {
                sheets.Add(new Sheet(excelFileLocation, el));
            }
            return sheets;
        }

        private static IEnumerable<XElement> GetXMLSheets(FileInfo excelFileLocation)
        {
            using var fileStream = File.Open(excelFileLocation.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var workbookArchive = archive.GetEntry("xl/workbook.xml");
            if (workbookArchive == null) throw new NullReferenceException("Файл xl/workbook.xml отсутствует");
            var stream = workbookArchive.Open();
            XDocument doc = XDocument.Load(stream);
            stream.Dispose();
            return (from el in doc.Descendants() where el.Name.LocalName == "sheet" select el).AsEnumerable();
        }

        //static public IEnumerable<Sheet> GetSheets(FileInfo fileLocation)
        //{
        //    var sheets = new List<Sheet>();
        //    foreach (var el in GetXMLSheets())
        //    {
        //        sheets.Add(new Sheet());
        //    }
        //    return sheets;
        //}

        //public Sheet? SheetByName(string sheetName)
        //{
        //    foreach (var sheet in GetSheets())
        //    {
        //        if (sheet.Name == sheetName)
        //        {
        //            return sheet;
        //        }
        //    }
        //    return null;
        //}
        private void SetSheetName(string Name)
        {

        }

        public void WriteString(string value, int row, string col)
        {
            var sharedStringNumber = SharedStringsHelper.AddSharedString(fileLocation, value);
            using var fileStream = File.Open(fileLocation.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var path = "xl/worksheets/" + SheetFileName;
            var relationshipArchive = archive.GetEntry(path);
            if (relationshipArchive == null) throw new NullReferenceException("Файл " + SheetFileName + "отсутствует");
            var stream = relationshipArchive.Open();
            XDocument doc = XDocument.Load(stream);
            var rows = (from r in doc.Descendants() where r.Name.LocalName == "row" select r).AsEnumerable();
        }

        public void WriteValue(string value, int row, string col)
        {
            using var fileStream = File.Open(fileLocation.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var path = "xl/worksheets/" + SheetFileName;
            var relationshipArchive = archive.GetEntry(path);
            if (relationshipArchive == null) throw new NullReferenceException("Файл " + SheetFileName + "отсутствует");
            var stream = relationshipArchive.Open();
            XDocument doc = XDocument.Load(stream);
            var rows = (from r in doc.Descendants() where r.Name.LocalName == "row" select r).ToList();
            var root = doc.Root;
            if (root == null) throw new NullReferenceException("в файле " + path + " отсутсвует элемент root");

            var rowEl = ExcelXMLHelper.FindRowByRowNumber(root, row);
            if (rowEl == null)
            {

            }
            else
            {
                var colEl = ExcelXMLHelper.FindColumnByColumnName(rowEl, col);
                if(colEl == null)
                {
                    
                }
                else
                {
                    var valueEl = (from el in colEl.Descendants() where el.Name.LocalName == "v" select el).FirstOrDefault();
                    if(valueEl != null)
                    {
                        valueEl.Value = value;
                        stream.Position = 0;
                        stream.SetLength(0);
                        doc.Save(stream);
                        stream.Dispose();
                        return;
                    }
                }
                
            }
            foreach (var r in rows)
            {
                var rowAttr = r.Attribute("r");
                if (rowAttr == null) throw new NullReferenceException("При парсе" + SheetFileName + "не найден атрибут r для row");
                if(rowAttr.Value == row.ToString())
                {
                    foreach(var c in r.Descendants())
                    {
                        var colAttr = c.Attribute("r");
                        if (colAttr == null) break;
                        if(colAttr.Value == col)
                        {
                            var valueElement = (from el in c.Descendants() where el.Name.LocalName == "v" select el).FirstOrDefault();
                            if (valueElement == null)
                            {
                                
                            }
                            else
                            {
                                valueElement.Value = value;
                                stream.Position = 0;
                                stream.SetLength(0);
                                doc.Save(stream);
                                stream.Dispose();
                            }
                            return;
                        }
                        else
                        {
                            ExcelXMLHelper.CreateColumnXElement(col, value, doc.Root?.GetDefaultNamespace());
                            
                        }
                    }
                }
                else
                {

                }
            }
        }

    }
}
