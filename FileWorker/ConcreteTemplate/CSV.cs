using System;
using System.Data;
using System.IO;

namespace FileWorker.ConcreteTemplate
{
    public class CSV : FileTemplate
    {
        public override DataTable ExtractData(string path)
        {
            return GetDataFromCSV(path);
        }

        public DataTable GetDataFromCSV(string path)
        {
            DataTable dtable = new DataTable();
            using (StreamReader sr = new StreamReader(path))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dtable.Columns.Add(header);
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    DataRow dr = dtable.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dtable.Rows.Add(dr);
                }
            }

            return dtable;
        }
    }
}
