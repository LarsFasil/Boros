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
        

        public string[] getData()
        {
            return File.ReadLines(@Directory.GetCurrentDirectory() + "\\Dict1.csv").ToArray();
        }

        //public void UpdateFile()
        //{
        //    string[][] output = new string[rowData.Count][];

        //    for (int i = 0; i < output.Length; i++)
        //    {
        //        output[i] = rowData[i];
        //    }

        //    int length = output.GetLength(0);
        //    string delimiter = ",";

        //    StringBuilder sb = new StringBuilder();

        //    for (int index = 0; index < length; index++)
        //        sb.AppendLine(string.Join(delimiter, output[index]));

        //    string s_filePath = GetPath();

        //    StreamWriter outStream = System.IO.File.CreateText(s_filePath);
        //    outStream.WriteLine(sb);
        //    outStream.Close();
        //}
        public void UpdateData(StringBuilder data)
        {
            //using (StreamWriter sr = new StreamWriter()
            //{
            //    sr.WriteLine(data);
            //}

            //StreamWriter outStream = File.CreateText(GetPath());
            //    outStream.WriteLine(data);
            //    outStream.Close();

            using (var stream = File.CreateText(GetPath()))
            {
                stream.WriteLine(data);
            }
        }

        string GetPath()
        {
            if (!Directory.Exists(@Directory.GetCurrentDirectory() + "\\Dict1") )
            {
                Directory.CreateDirectory(@Directory.GetCurrentDirectory() + "\\Dict1");
            }
            return @Directory.GetCurrentDirectory() + "\\Dict1.csv";
        }
    }
}
