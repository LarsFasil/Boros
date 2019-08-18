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
    class Excel
    {
        #region excel vars
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        #endregion

        public Excel(String path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public void WriteCell(int row, int col, string data)
        {
            //Alsof excel bij 0 begint
            row++;
            col++;
            ws.Cells[row, col] = data;
            wb.Save();
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
            {

                return ws.Cells[i, j].Value2;
            }
            else
            {
                return "";
            }
        }

        public void Close()
        {
            wb.Close(0);
        }
    }
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

        //string GetPath()
        //{
        //    if (!Directory.Exists(@Directory.GetCurrentDirectory() + "\\Dict1") )
        //    {
        //        Directory.CreateDirectory(@Directory.GetCurrentDirectory() + "\\Dict1");
        //    }
        //    return @Directory.GetCurrentDirectory() + "\\Dict1.csv";
        //}
    }
}
