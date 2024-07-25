using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WorkingWithExcel
{
    public class ExcelWorkSpace
    {
        public FileInfo OriginalFileLocation { get; private set; }
        public FileInfo? DuplicateFileLocation { get; private set; } = null;
        
        public ExcelWorkSpace(FileInfo fileLocation)
        {
            OriginalFileLocation = fileLocation;
        }
        public ExcelFile Open()
        {
            return new ExcelFile(OriginalFileLocation);
        }

        /// <summary>
        /// Открывает файл Excel по указанному пути и создает 
        /// в этой же директории дупликат в котором и будут 
        /// осуществляться изменения
        /// </summary>
        /// <param name="duplicateFileName">
        /// Имя для файла-дупликата
        /// </param>
        /// <returns>
        /// Возвращает класс для работы с Excel файлом
        /// </returns>
        public ExcelFile OpenDuplicate(FileInfo duplicateLocation)
        {
            OriginalFileLocation.CopyTo(duplicateLocation.FullName);
            DuplicateFileLocation = duplicateLocation;
            return new ExcelFile(duplicateLocation);
        }

        /// <summary>
        /// Открывает файл Excel по указанному пути и создает 
        /// в этой же директории дупликат в котором и будут 
        /// осуществляться изменения
        /// </summary>
        /// <param name="duplicateFileName">
        /// Имя для файла-дупликата
        /// </param>
        /// <returns>
        /// Возвращает класс для работы с Excel файлом
        /// </returns>
        public ExcelFile OpenDuplicate(string duplicateFileName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Открывает файл Excel по указанному пути и создает 
        /// в этой же директории дупликат в котором и будут 
        /// осуществляться изменения. Для имени дупликата
        /// используется уникальный идентификатор
        /// <returns>
        /// Возвращает класс для работы с Excel файлом
        /// </returns>
        public ExcelFile OpenDuplicate()
        {
            throw new NotImplementedException();
        }

        public ExcelFile Create() 
        { 
            throw new NotImplementedException();
        }     
    }
}
