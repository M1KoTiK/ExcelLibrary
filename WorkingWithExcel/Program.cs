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
            var inputFilePath = new FileInfo("C:\\Users\\User\\Desktop\\Проекты\\ExcelLibrary\\testExcel.xlsx");
            var ex = new ExcelWorkspace(inputFilePath).Open();
            
            Console.WriteLine(ex.DocInfo.Author);
            Console.WriteLine(ex.DocInfo.ModifyDate);
            ex.DocInfo.Author = "New Author";
            Console.WriteLine(ex.DocInfo.Author);
            foreach (var el in ex.Sheets)
            {
                Console.WriteLine(el.Name);
            }
            ex.Sheets.First().WriteString("test1", 1, "A1");

            //setRelationshipForMedia(inputFilePath, imageFile, "image1");
        }

        
    }
}
