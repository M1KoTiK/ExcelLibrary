using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class Sheet
    {
        public string Name { get; private set; }
        public string SheetId { get; private set; }
        public string RId { get; private set; }
        public string SheetFileName { get; private set; }


        ExcelWorkspace workspace;

        private XElement xmlSheetElement;

        public Sheet(ExcelWorkspace workspace, XElement XMLSheetElement) 
        {
            this.workspace = workspace;
            this.xmlSheetElement = XMLSheetElement;

            var nameAttr = xmlSheetElement.Attribute("name");
            if (nameAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут name");
            Name = nameAttr.Value;

            var sheetIdAttr = xmlSheetElement.Attribute("sheetId");
            if(sheetIdAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут sheetId");
            SheetId = sheetIdAttr.Value;

            var ns = xmlSheetElement.GetNamespaceOfPrefix("r");
            var rIdAttr = xmlSheetElement.Attribute(ns + "id");
            if (rIdAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут r:id");
            RId = rIdAttr.Value;

            using var fileStream = File.Open(workspace.ExcelFile.FullName, FileMode.Open);
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
        public void WriteValue(string value, int row, int col)
        {

        }

    }
}
