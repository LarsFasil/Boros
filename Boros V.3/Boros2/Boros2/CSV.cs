using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace Boros2
{
    class CSV
    {
        public string[] getData(string path)
        {
            if (!File.Exists(path))
            {
                //Create file and close the process right away
                File.CreateText(path).Close();
            }
            return File.ReadLines(path).ToArray();
        }

    
        public void UpdateData(string[] data, string path)
        {
             using (StreamWriter stream = File.CreateText(path))
            {
                foreach (var item in data)
                {
                    stream.WriteLine(item);
                }
                
            }
        }
    }
}
