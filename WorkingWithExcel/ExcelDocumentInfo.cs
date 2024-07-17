using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace WorkingWithExcel
{
    public class ExcelDocumentInfo
    {

        private string author;
        public string Author
        {
            get => author;
            set
            {
                author = value;
                //SetAuthor(value);
            }
        }

        private DateTime creationDate;
        public DateTime CreationDate
        {
            get => creationDate;
            set
            {
                creationDate = value;
               // SetCreationDate(value);
            }
        }

        private DateTime modifyDate;
        public DateTime ModifyDate
        {
            get => modifyDate;
            set
            {
                modifyDate = value;
                //SetModifyDate(value);
            }
        }

        private string lastModifyAuthor;
        public string LastModifyAuthor
        {
            get => lastModifyAuthor;
            set
            {
                lastModifyAuthor = value;
                //SetlastModifyAuthor(value);
            }
        }

        private FileInfo FileLocation { get; }


        private XDocument propsCoreXMLDoc;
        public ExcelDocumentInfo(FileInfo fileLocation)
        {
            FileLocation = fileLocation;
            RefreshData();
        }
        public void RefreshData()
        {
            propsCoreXMLDoc = GetCorePropsXML();
            author = FindXElementInDocCorePropsByName("creator").Value;
            lastModifyAuthor = FindXElementInDocCorePropsByName("lastModifiedBy").Value;
            creationDate = DateTime.Parse(FindXElementInDocCorePropsByName("created").Value.Replace("T", " "));
            modifyDate = DateTime.Parse(FindXElementInDocCorePropsByName("modified").Value.Replace("T", " "));
        }

        private XElement FindXElementInDocCorePropsByName(string elementName)
        {
            var els = propsCoreXMLDoc.Descendants();
            foreach (var el in els)
            {
                if (el.Name.LocalName == elementName)
                {
                    return el;
                }
            }
            throw new NullReferenceException($"В файле {docPropsPath} отсутствует элемент {elementName}");
        }
        private XDocument GetCorePropsXML()
        {
            using (var fileStream = File.Open(FileLocation.FullName, FileMode.Open))
            {
                using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
                var entry = archive.GetEntry(docPropsPath);
                if (entry == null) throw new NullReferenceException($"Файл {docPropsPath} отсутствует");
                var stream = entry.Open();
                XDocument doc = XDocument.Load(stream);
                stream.Dispose();
                return doc;
            }
        }

        private string docPropsPath = "docProps/core.xml";
    }
}
