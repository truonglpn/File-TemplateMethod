using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FileWorker.ConcreteTemplate
{
    public class Excel : FileTemplate
    {
   
        public override DataTable ExtractData(string path)
        {
            return GetDataFromExcel(path);
        }

        /// <summary>
        /// open dialog, all option choose file and filter 
        /// </summary>
        /// <param name="theDialog"></param>
        /// <param name="titleDialog"></param>
        /// <param name="nameFilter"></param>
        /// <param name="filterOption"></param>
        /// <param name="initialDirectionOpenFile"></param>
        /// <returns></returns>
        public override OpenFileDialog OpenDialogStream(OpenFileDialog theDialog, string titleDialog, string nameFilter, string filterOption, string initialDirectionOpenFile)
        {
            return base.OpenDialogStream(theDialog, titleDialog, nameFilter, filterOption, initialDirectionOpenFile);
        }


        /// <summary>
        /// Get data from excel path 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="hasHeader"></param>
        /// <returns></returns>
        public  DataTable GetDataFromExcel(string path, bool hasHeader = true)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }
                    var ws = pck.Workbook.Worksheets.First();

                    TrimLastEmptyRows(ws);

                    DataTable tbl = new DataTable();
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    }
                    var startRow = hasHeader ? 2 : 1;
                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = tbl.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }
                    return tbl;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// get only data column on excel
        /// </summary>
        /// <param name="path"></param>
        /// <param name="hasHeader"></param>
        /// <returns></returns>
        public  DataTable GetDataColumnFromExcel(string path, bool hasHeader = true)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }
                    var ws = pck.Workbook.Worksheets.First();


                    DataTable tbl = new DataTable();
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    }

                    return tbl;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Trim last empty row in excel
        /// </summary>
        /// <param name="worksheet"></param>
        public void TrimLastEmptyRows(ExcelWorksheet worksheet)
        {
            while (IsLastRowEmpty(worksheet))
                worksheet.DeleteRow(worksheet.Dimension.End.Row);
        }

        /// <summary>
        /// check last row is Empty?
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public  bool IsLastRowEmpty(ExcelWorksheet worksheet)
        {
            var empties = new List<bool>();

            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            {
                var rowEmpty = worksheet.Cells[worksheet.Dimension.End.Row, i].Value == null ? true : false;
                empties.Add(rowEmpty);
            }

            return empties.All(e => e);
        }


    }
}
