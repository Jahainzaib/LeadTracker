using LeadTracker.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LeadTracker.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View("Create");
        }
        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Create(Data data)
        {
            var fileName = DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
            var filePath = Path.Combine(Server.MapPath("~/App_Data/"), fileName);

            Application excelApp = new Application();
            Workbook workbook = null;
            Worksheet worksheet = null;
            object misValue = System.Reflection.Missing.Value;
            int row = 0;
            try
            {
                excelApp.DisplayAlerts = false; // Disable alerts

                // Check if file exists
                if (System.IO.File.Exists(filePath))
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                    worksheet = workbook.Worksheets[1];
                }
                else
                {
                    workbook = excelApp.Workbooks.Add(misValue);
                    worksheet = (Worksheet)workbook.Worksheets[1];
                    worksheet.Name = "Data";
                    worksheet.Cells[1, 1].Value = "Agent ID Name";
                    worksheet.Cells[1, 2].Value = "Agent Name";
                    worksheet.Cells[1, 3].Value = "Client Name";
                    worksheet.Cells[1, 4].Value = "Client Email";
                    worksheet.Cells[1, 5].Value = "Product";
                    worksheet.Cells[1, 6].Value = "Proper Details";
                    worksheet.Cells[1, 7].Value = "Total Sale";
                    worksheet.Cells[1, 8].Value = "Upfront";
                    worksheet.Cells[1, 9].Value = "Remaining";
                    Range headerRange = worksheet.Range["A1", "I1"];
                    headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    headerRange.Borders.Weight = XlBorderWeight.xlMedium;
                    headerRange.Font.Size = 12;
                    headerRange.Font.Bold =true;
                    row = 2;
                }

                // Find the next empty row in the worksheet
                row = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row + 1;

                // Fill in data
                worksheet.Cells[row, 1].Value = data.Agent_ID_Name;
                worksheet.Cells[row, 2].Value = data.Agent_Name;
                worksheet.Cells[row, 3].Value = data.Client_Name;
                worksheet.Cells[row, 4].Value = data.Client_Email;
                worksheet.Cells[row, 5].Value = data.Product;
                worksheet.Cells[row, 6].Value = data.Proper_Details;
                worksheet.Cells[row, 7].Value = data.Total_Sales;
                worksheet.Cells[row, 8].Value = data.Upfront;
                worksheet.Cells[row, 9].Value = data.Remaining;

                // Save and close the workbook
                if (System.IO.File.Exists(filePath))
                {
                    workbook.Save();
                }
                else
                {
                    workbook.SaveAs(filePath);
                }

                workbook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                // Log exception
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Properly release COM objects
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // Return a response indicating success (could be a different response based on your application needs)
            return View();
        }
    }
}