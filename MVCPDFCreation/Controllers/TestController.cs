using HtmlAgilityPack;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using MVCPDFCreation.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat;

namespace MVCPDFCreation.Controllers
{
    public class TestController : Controller
    {
        // GET: Test
        public ActionResult Index()
        {
            using (FirstDatabaseEntities db = new FirstDatabaseEntities())
            {
                var emp = db.Employees.Take(10).ToList();
                return View(emp);
            }
        }
        [HttpPost]
        [ValidateInput(false)]
        public FileResult ExportPDF(string GridHtml)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                StringReader sr = new StringReader(GridHtml);
                Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 100f, 0f);
                PdfWriter writer = PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                XMLWorkerHelper.GetInstance().ParseXHtml(writer, pdfDoc, sr);
                pdfDoc.Close();
                return File(stream.ToArray(), "application/pdf", "Grid.pdf");
            }
        }
        [HttpPost]
        [ValidateInput(false)]
        public FileResult ExportExcel(string HtmlGrid)
        {
            string TableCell = "";
            DataTable dt = new DataTable();
            HtmlDocument htmldoc = new HtmlDocument();
            htmldoc.LoadHtml(@"<html><body>" + HtmlGrid + "</body></html>");
            foreach (HtmlNode table in htmldoc.DocumentNode.SelectNodes("//table"))
            {
                foreach (HtmlNode body in table.SelectNodes("tbody"))
                {
                    foreach (HtmlNode row in body.SelectNodes("tr"))
                    {
                        if (row.SelectSingleNode("th") != null)
                        {
                            foreach (HtmlNode cell in row.SelectNodes("th"))
                            {
                                dt.Columns.Add(cell.InnerText);
                            }
                        }
                        if (row.SelectSingleNode("td") != null)
                        {
                            DataRow datarow = dt.NewRow();
                            int count = 0;
                            foreach (HtmlNode cell in row.SelectNodes("td"))
                            {
                                datarow[count++] = cell.InnerText;
                            }
                            dt.Rows.Add(datarow);
                        }
                    }
                }
            }
            WriteExcel(dt);
            //  you can give any dynamic path here
            return File("D:\\EmployeeDetails.xlsx", "application/ms-excel", "Grid.xlsx");
        }


        public void WriteExcel(DataTable dt)
        {
            // Reference BY http://www.encodedna.com/2013/01/asp.net-export-to-excel.htm
            // ADD A WORKBOOK USING THE EXCEL APPLICATION.
            Excel.Application xlAppToExport = new Excel.Application();
            xlAppToExport.Workbooks.Add("");

            // ADD A WORKSHEET.
            Excel.Worksheet xlWorkSheetToExport = default(Excel.Worksheet);
            xlWorkSheetToExport = (Excel.Worksheet)xlAppToExport.Sheets["Sheet1"];

            // ROW ID FROM WHERE THE DATA STARTS SHOWING.
            int iRowCnt = 4;

            // SHOW THE HEADER.
            xlWorkSheetToExport.Cells[1, 1] = "Employee Details";

            Excel.Range range = xlWorkSheetToExport.Cells[1, 1] as Excel.Range;
            range.EntireRow.Font.Name = "Calibri";
            range.EntireRow.Font.Bold = true;
            range.EntireRow.Font.Size = 20;

            xlWorkSheetToExport.Range["A1:E1"].MergeCells = true;       // MERGE CELLS OF THE HEADER.



            // SHOW COLUMNS ON THE TOP.
            int count = 1;
            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                xlWorkSheetToExport.Cells[iRowCnt - 1, count++] = dt.Columns[i].ToString();
            }

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                int flag = 0;
                int counter = 0;
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    xlWorkSheetToExport.Cells[iRowCnt, ++counter] = dt.Rows[i][dt.Columns[flag++].ToString()];
                }
                //xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i][dt.Columns[0].ToString()];
                //xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i][dt.Columns[1].ToString()];
                //xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i][dt.Columns[2].ToString()];
                //xlWorkSheetToExport.Cells[iRowCnt, 4] = dt.Rows[i][dt.Columns[3].ToString()];
                //xlWorkSheetToExport.Cells[iRowCnt, 5] = dt.Rows[i][dt.Columns[4].ToString()];
                iRowCnt = iRowCnt + 1;
            }

            // FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
            Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Excel.Range;
            range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);

            // SAVE THE FILE IN A FOLDER.
            xlWorkSheetToExport.SaveAs("D:\\EmployeeDetails.xlsx");

            // CLEAR.
            xlAppToExport.Workbooks.Close();
            xlAppToExport.Quit();
            xlAppToExport = null;
            xlWorkSheetToExport = null;
        }
    }
}