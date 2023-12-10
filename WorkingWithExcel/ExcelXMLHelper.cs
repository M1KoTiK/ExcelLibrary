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
}
