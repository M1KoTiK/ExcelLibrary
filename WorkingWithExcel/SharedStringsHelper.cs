﻿using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    internal class SharedStringsHelper
    {
        static public IEnumerable<XElement> GetSharedStringXML(ExcelWorkspace workspace)
        {

            using var fileStream = File.Open(workspace.ExcelFile.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var workbookArchive = archive.GetEntry("xl/sharedStrings.xml");
            if (workbookArchive == null) throw new NullReferenceException("Файл xl/sharedStrings.xml отсутствует");
            var stream = workbookArchive.Open();
            XDocument doc = XDocument.Load(stream);
            stream.Dispose();
            return (from el in doc.Descendants() where el.Name.LocalName == "si" select el).AsEnumerable();
        }

        static public int FindSharedString(ExcelWorkspace workspace, string str)
        {
            var xmlSharedStringList = GetSharedStringXML(workspace).ToList();
            for (int i = 0; i < xmlSharedStringList.Count(); i++)
            {
                if (xmlSharedStringList[i].Value == str)
                {
                    return i;
                }
            }
            return -1;
        }

        static public int AddSharedString(ExcelWorkspace workspace, string sharedString)
        {
            if (FindSharedString(workspace, sharedString) != -1)
            {
                return -1;
            }

            using var fileStream = File.Open(workspace.ExcelFile.FullName, FileMode.Open);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
            var workbookArchive = archive.GetEntry("xl/sharedStrings.xml");
            if (workbookArchive == null) throw new NullReferenceException("Файл xl/sharedStrings.xml отсутствует");

            var stream = workbookArchive.Open();
            XDocument doc = XDocument.Load(stream);
            doc.Root.Add(ExcelXMLHelper.CreateSharedStringXElement(sharedString, doc.Root.GetDefaultNamespace()));

            //Очищаем стрим чтобы данные в файле не дублировались
            stream.Position = 0;
            stream.SetLength(0);

            doc.Save(stream);
            stream.Dispose();
            return FindSharedString(workspace, sharedString);

        }
    }
}