using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class ExcelXML
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
    }
}
