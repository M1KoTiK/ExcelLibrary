using Microsoft.VisualBasic;
using System;
using System.IO.Compression;
using System.Net.Mime;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    class Program
    {
        public static void Main(string[] args)
        {
            var imageFile = new FileInfo("C:\\Users\\User\\Pictures\\image_kaiji.jpg");

            var inputFilePath = new FileInfo("C:\\Users\\User\\Desktop\\input.xlsx");
            var outputFilePath = "C:\\Users\\User\\Desktop\\output.xlsx";
            var ex = new ExcelWorkspace(inputFilePath);
            foreach(var el in ex.Sheets)
            {
                Console.WriteLine(el.Name);
            }
            //setRelationshipForMedia(inputFilePath, imageFile, "image1");
        }
        public static void setRelationshipForMedia(FileInfo excelFile, FileInfo imageName, string idForRelationship)
        {

            using var fileStream = File.Open(excelFile.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var relationshipArchive = archive.GetEntry("xl/drawings/_rels/drawing1.xml.rels");
            
            if (relationshipArchive == null)
            {
                archive.CreateEntry("xl/drawings/_rels/drawing1.xml.rels");
                relationshipArchive = archive.GetEntry("xl/drawings/_rels/drawing1.xml.rels");
            }

            var stream = relationshipArchive.Open();

            XDocument doc;
            XNamespace xmlns = "http://schemas.openxmlformats.org/package/2006/relationships";

            try
            {
                doc = XDocument.Load(stream);
            }
            catch(Exception e)
            {
                
                doc = new XDocument();
                doc.Add(new XElement("Relationships"));
                doc.Root.Name = xmlns + doc.Root.Name.LocalName;
            }

            var rootElement = (from el in doc.Descendants() where el.Name.LocalName == "Relationships" select el).FirstOrDefault();
            doc.Root.Add(CreateRelationshipImageXElement(doc.Root.Name, idForRelationship, imageName.Name));

            stream.Position = 0;
            stream.SetLength(0);
            doc.Save(stream);
            stream.Dispose();

            Console.WriteLine(doc);




        }
        public static XElement CreateRelationshipImageXElement(XName ns, string id, string imageNameWithExt)
        {
            XElement newRelationShipElement = new XElement(ns + "Relationship",
                new XAttribute("Id", id),
                new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/image"),
                new XAttribute("Target", "../media/" + imageNameWithExt));
            return newRelationShipElement;
        }
        public static void WriteImageInExcelMedia(FileInfo excelFile, FileInfo imageFile, string? customImageName = null )
        {
            using var fileStream = File.Open(excelFile.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var imageFileName = imageFile.Name;
            if(customImageName != null )
            {
                imageFileName = customImageName;
            }
            string archiveEntryName = "xl/media/" + imageFileName + imageFile.Extension;

            archive.CreateEntryFromFile(imageFile.FullName, archiveEntryName);
        }
    }
}
