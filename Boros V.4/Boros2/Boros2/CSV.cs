using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;


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

        public string[] GETIT(string path, int column)
        {
            using (TextFieldParser csvParser = new TextFieldParser(path))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                // Skip the row with the column names
                csvParser.ReadLine();

                List<string> StringList = new List<string>();

                while (!csvParser.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    string[] fields = csvParser.ReadFields();
                    //string Name = fields[0];
                    //string Address = fields[1];
                    StringList.Add(fields[column]);
                }
                return StringList.ToArray();
            }
        }
    }
}
