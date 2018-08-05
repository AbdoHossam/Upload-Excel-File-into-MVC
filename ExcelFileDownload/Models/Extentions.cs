using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Web;
using System.Web.Mvc;
using System.Web.Mvc.Html;
using OfficeOpenXml;

namespace ExcelFileDownload.Models
{
    public static class Extentions
    {
        public static DataTable ExcelFileToDataTable(this HttpPostedFileBase file, string[] cols, out bool valid, string directorPath)
        {
            ExcelPackage package;

            string path = null;
            FileInfo newPath = null;
            if (file.FileName.EndsWith(".xls"))
            {
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory + directorPath, Path.GetFileName(file.FileName));

                //check if already exist
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                //save the File
                file.SaveAs(path);

                ExcelConvert(path);
                newPath = new FileInfo(path + "x");
                package = new ExcelPackage(newPath);
            }
            else
            {
                package = new ExcelPackage(file.InputStream);
            }

            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();

            //check if the parameters are valid
            valid = workSheet.Cells.Select(c => c.Text).Intersect(cols).Count() == cols.Length;
            if (!valid) return null;

            DataTable table = new DataTable();
            foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
            {
                table.Columns.Add(firstRowCell.Text);
            }

            for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                var newRow = table.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                table.Rows.Add(newRow);
            }

            //delete the excel files in temp folder
            if (path != null && File.Exists(path))
            {
                File.Delete(path);
            }
            if (newPath != null && File.Exists(newPath.ToString()))
            {
                File.Delete(newPath.ToString());
            }
            //dispose the object
            package.Dispose();
            return table;
        }



        private static void ExcelConvert(string fileName)
        {
            if (fileName == null) throw new ArgumentNullException(nameof(fileName));
            var app = new Microsoft.Office.Interop.Excel.Application();
            var wb = app.Workbooks.Open(fileName);
            wb.SaveAs(fileName + "x", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
        }
    }
}