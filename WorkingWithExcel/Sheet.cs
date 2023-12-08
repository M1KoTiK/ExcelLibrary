using System;
using System.Collections.Generic;
using System.ComponentModel.Design.Serialization;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class Sheet
    {
        public string Name 
        {
            get
            {
                var nameAttr = xmlSheetElement.Attribute("name");
                if (nameAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут name");
                return nameAttr.Value;
            }
        }
        public string SheetId 
        { 
            get 
            {
                var sheetIdAttr = xmlSheetElement.Attribute("sheetId");
                if (sheetIdAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут sheetId");
                return sheetIdAttr.Value;
            } 
        }
        public string RId
        {
            get
            {
                var ns = xmlSheetElement.GetNamespaceOfPrefix("r");
                var rIdAttr = xmlSheetElement.Attribute(ns + "id");
                if (rIdAttr == null) throw new NullReferenceException("При попытки парсить XmlSheet не найден атрибут r:id");
                return rIdAttr.Value;
            }
        }

        public string Target 
        { 
            get 
            {
                using var fileStream = File.Open(workspace.ExcelFile.FullName, FileMode.Open);
                using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
                var relationshipArchive = archive.GetEntry("xl/_rels/workbook.xml.rels");
                if (relationshipArchive == null) throw new NullReferenceException("Файл workbook.xml.rels отсутствует");
                var stream = relationshipArchive.Open();
                XDocument doc = XDocument.Load(stream);

                var relationshipsElements = (from el in doc.Descendants() where el.Name.LocalName == "Relationships" select el).FirstOrDefault();
                if (relationshipsElements == null) throw new NullReferenceException("при попытике спарсить workbook.xml.rels не найден атрибут Relationships");

                var relationshipElement = (from el in relationshipsElements.Descendants() where el.Attribute("Id")?.Value == RId select el).FirstOrDefault();
                if (relationshipElement == null) throw new NullReferenceException("При попытки парсить workbook.xml.rels не найден элемент соответствующий данной Sheet");

                var targetAttr = relationshipElement.Attribute("Target");
                if (targetAttr == null) throw new NullReferenceException("При попытки парсить workbook.xml.rels не найден атрибут Target");
                stream.Dispose();
                return targetAttr.Value;               
            } 
        }
        private XElement relationship 
        { 
            get
            { 

            } 
        }
        private ExcelWorkspace workspace;

        private XElement xmlSheetElement;
        public Sheet(ExcelWorkspace WS, XElement XMLSheetElement) 
        {
            this.xmlSheetElement= XMLSheetElement;
            this.workspace = WS;                       
        }

        public WriteAtIndex()
        {

        }

    }
}
