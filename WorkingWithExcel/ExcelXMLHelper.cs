using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class ExcelXMLHelper
    {
        public static XElement CreateRelationshipXElement(string id, string imageNameWithExt, XNamespace? ns = null)
        {
            return new XElement(ns + "Relationship",
                new XAttribute("Id", id),
                new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/image"),
                new XAttribute("Target", "../media/" + imageNameWithExt));            
        }

        public static XElement CreateSharedStringXElement(string sharedString, XNamespace? ns = null)
        {
            return new XElement(ns + "si", new XElement(ns + "t", sharedString));
        }

        public static XElement CreateColumnXElement(string col, string value, XNamespace? ns = null)
        {
            return new XElement(ns + "с",new XAttribute("r", col), new XElement(ns + "м", value));
        }

        public static XElement CreateRowXElement(string row, string col, string value, XNamespace? ns = null)
        {
            return new XElement(ns + "row", new XAttribute("r",row), new XElement(ns + "с", new XAttribute("r", col), new XElement(ns + "v", value)));
        }
       
        public static XElement? FindRowByRowNumber(XElement RowRootElement, int rowNumber)
        {
            foreach(var el in RowRootElement.Descendants())
            {
                var rowAttr = el.Attribute("r");
                if (rowAttr == null) break;
                if (rowAttr.Value == rowNumber.ToString());
                {
                    return el;
                }
            }
            return null;
        }

        public static XElement? FindColumnByColumnName(XElement ColumnRootElement, string colName)
        {
            foreach (var el in ColumnRootElement.Descendants())
            {
                var rowAttr = el.Attribute("r");
                if (rowAttr == null) break;
                if (rowAttr.Value == colName.ToString());
                {
                    return el;
                }
            }
            return null;
        }

    }
    //public static void setRelationshipForMedia(FileInfo excelFile, FileInfo imageName, string idForRelationship)
    //{
    //    using var fileStream = File.Open(excelFile.FullName, FileMode.Open);
    //    using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
    //    var relationshipArchive = archive.GetEntry("xl/drawings/_rels/drawing1.xml.rels");

    //    if (relationshipArchive == null)
    //    {
    //        archive.CreateEntry("xl/drawings/_rels/drawing1.xml.rels");
    //        relationshipArchive = archive.GetEntry("xl/drawings/_rels/drawing1.xml.rels");
    //    }

    //    var stream = relationshipArchive.Open();

    //    XDocument doc;
    //    XNamespace xmlns = "http://schemas.openxmlformats.org/package/2006/relationships";

    //    try
    //    {
    //        doc = XDocument.Load(stream);
    //    }
    //    catch(Exception e)
    //    {

    //        doc = new XDocument();
    //        doc.Add(new XElement("Relationships"));
    //        doc.Root.Name = xmlns + doc.Root.Name.LocalName;
    //    }

    //    var rootElement = (from el in doc.Descendants() where el.Name.LocalName == "Relationships" select el).FirstOrDefault();
    //    doc.Root.Add(CreateRelationshipImageXElement(doc.Root.Name, idForRelationship, imageName.Name));

    //    stream.Position = 0;
    //    stream.SetLength(0);
    //    doc.Save(stream);
    //    stream.Dispose();

    //    Console.WriteLine(doc);

    //}
    //public static XElement CreateRelationshipImageXElement(XName ns, string id, string imageNameWithExt)
    //{
    //    XElement newRelationShipElement = new XElement(ns + "Relationship",
    //        new XAttribute("Id", id),
    //        new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/image"),
    //        new XAttribute("Target", "../media/" + imageNameWithExt));
    //    return newRelationShipElement;
    //}
    //public static void WriteImageInExcelMedia(FileInfo excelFile, FileInfo imageFile, string? customImageName = null )
    //{
    //    using var fileStream = File.Open(excelFile.FullName, FileMode.Open);
    //    using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
    //    var imageFileName = imageFile.Name;
    //    if(customImageName != null )
    //    {
    //        imageFileName = customImageName;
    //    }
    //    string archiveEntryName = "xl/media/" + imageFileName + imageFile.Extension;

    //    archive.CreateEntryFromFile(imageFile.FullName, archiveEntryName);
    //}
}
