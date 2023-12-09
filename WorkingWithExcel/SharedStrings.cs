using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class SharedStrings
    {
        public int Count { get; private set; }
        public int UniqueCount { get; private set; }


        public SharedStrings(ExcelWorkspace workspace) 
        {
            using var fileStream = File.Open(workspace.ExcelFile.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var sharedStringsArchive = archive.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsArchive == null) throw new NullReferenceException("Файл sharedStrings.xml отсутствует");
            var stream = sharedStringsArchive.Open();
            XDocument doc = XDocument.Load(stream);
            var sharedStringArray = new List<string>();


        }
        
    }
}
