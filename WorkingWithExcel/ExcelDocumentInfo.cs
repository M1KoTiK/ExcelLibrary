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
                SetAttributeValueByName("creator", value);
            }
        }

        private DateTime creationDate;
        public DateTime CreationDate
        {
            get => creationDate.ToLocalTime();
            set
            {
                creationDate = value;
                SetAttributeValueByName("created", value.ToUniversalTime().ToString("s"));
            }
        }

        private DateTime modifyDate;
        public DateTime ModifyDate
        {
            get => modifyDate.ToLocalTime();
            set
            {
                modifyDate = value;
                SetAttributeValueByName("modified", value.ToUniversalTime().ToString("s"));
            }
        }

        private string lastModifyAuthor;
        public string LastModifyAuthor
        {
            get => lastModifyAuthor;
            set
            {
                lastModifyAuthor = value;
                SetAttributeValueByName("modified", value);
            }
        }

        private FileInfo FileLocation { get; }


        private XDocument propsCoreXMLDoc;
        public ExcelDocumentInfo(FileInfo fileLocation)
        {
            FileLocation = fileLocation;
            propsCoreXMLDoc = GetCorePropsXML();
            author = FindAttributeByName("creator").Value;
            lastModifyAuthor = FindAttributeByName("lastModifiedBy").Value;
            creationDate = DateTime.Parse(FindAttributeByName("created").Value);
            modifyDate = DateTime.Parse(FindAttributeByName("modified").Value);
        }
        
        /// <summary>
        /// В случае стороннего изменения файла во время 
        /// работы библиотеки, может понадобится
        /// вручную обновить данные
        /// </summary>
        public void RefreshData()
        {
            propsCoreXMLDoc = GetCorePropsXML();
            author = FindAttributeByName("creator").Value;
            lastModifyAuthor = FindAttributeByName("lastModifiedBy").Value;
            creationDate = DateTime.Parse(FindAttributeByName("created").Value);
            modifyDate = DateTime.Parse(FindAttributeByName("modified").Value);
        }

        private XElement FindAttributeByName(string elementName)
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
        private void SetAttributeValueByName(string elementName, string value)
        {
            using (var fileStream = File.Open(FileLocation.FullName, FileMode.Open))
            {
                using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
                var entry = archive.GetEntry(docPropsPath);
                if (entry == null) throw new NullReferenceException($"Файл {docPropsPath} отсутствует");
                var stream = entry.Open();
                XDocument doc = XDocument.Load(stream);
                foreach (var el in doc.Descendants())
                {
                    if (el.Name.LocalName == elementName)
                    {
                        el.Value = value;
                    }
                }
                stream.Position = 0;
                stream.SetLength(0);
                doc.Save(stream);
                stream.Dispose();
            }
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
