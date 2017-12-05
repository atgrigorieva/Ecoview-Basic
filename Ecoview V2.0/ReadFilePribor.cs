using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Ecoview_V2._0{
    
    class ReadFilePribor
    {
        EcoviewProfessional1 _Analis;
        string filepathRead;
        public ReadFilePribor(string filepath, EcoviewProfessional1 parent)
        {
            this.filepathRead = filepath;
            this._Analis = parent;
            FileReadPribor();
        }
        public void FileReadPribor()
        {
            _Analis.filereadpribor = File.ReadAllLines(filepathRead, Encoding.UTF8);

            
        }
    }
}
