﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.IO.Enumeration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class ExcelFile
    {
        public FileInfo FileLocation { get; private set; }
        
        public ExcelDocumentInfo DocInfo { get; private set; }
        public ExcelFile(FileInfo fileLocation)
        {
            FileLocation = fileLocation;
            DocInfo = new ExcelDocumentInfo(FileLocation);
        }

        private void TryLoadInfoFromFile()
        {

        }       
    }
}