using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OfficeOpenXml;

namespace excelProject
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            

            // Define the file path using Server.MapPath
            string folderPath = Server.MapPath("~/ExcelFiles/");  // Create ExcelFiles folder in the root directory
            string filePath = Path.Combine(folderPath, "my.xlsx");

            string userexcelfilePath = Server.MapPath("./Excel/10082024 - HourlyMWHRecord-MPCL.08-10-2024 (2).xlsx");

          
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            using (var package = new ExcelPackage())
            {
                
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

               
                worksheet.Cells["A1"].Value = "DATE Time 1";
                worksheet.Cells["B1"].Value = "(MWh) Previous";
                worksheet.Cells["C1"].Value = "(MWh) Present";
                worksheet.Cells["D1"].Value = "C-B";
                worksheet.Cells["E1"].Value = "(MWh) Diff-A";
                worksheet.Cells["F1"].Value = "(MWh) Previous";
                worksheet.Cells["G1"].Value = "(MWh) Present";
                worksheet.Cells["H1"].Value = "(MWh) Diff-B";
                worksheet.Cells["I1"].Value = "T.Dispatch A+B MWh";
                worksheet.Cells["J1"].Value = "DATE Time 2";
                worksheet.Cells["K1"].Value = "(MWh) Previous";
                worksheet.Cells["L1"].Value = "(MWh) Present";
                worksheet.Cells["M1"].Value = "L-K";
                worksheet.Cells["N1"].Value = "(MWh) Diff-A";
                worksheet.Cells["O1"].Value = "(MWh) Previous";
                worksheet.Cells["P1"].Value = "(MWh) Present";
                worksheet.Cells["Q1"].Value = "(MWh) Diff-B";
                worksheet.Cells["R1"].Value = "T.Dispatch A+B MWh";

                using (var packagee = new ExcelPackage(userexcelfilePath))
                {

                    var worksheett = packagee.Workbook.Worksheets[0];

                    for (int i = 13; i >= 1; i--)
                    {
                        worksheett.DeleteRow(i);
                    }
                    worksheett.Cells["J25:Q28"].Clear();
                    worksheett.Cells["A34:H37"].Clear();

                    int rowCount = 0;
                    for (int row = 1; row <= worksheett.Dimension.End.Row; row++)
                    {
                        if (worksheett.Cells[row, 1].Value != null)
                        {
                            rowCount++;
                        }
                    }

                    int rowCountt = 0;
                    for (int row = 1; row <= worksheett.Dimension.End.Row; row++)
                    {
                        if (worksheett.Cells[row, 10].Value != null)
                        {
                            rowCountt++;
                        }
                    }
                    int f = 2;
                    for (int j = 1; j <= rowCount; j++)
                    {
                        for (int i = 1; i <= rowCountt; i++)
                        {
                            string addressA = $"A{j}";
                            string addressJ = $"J{i}";

                            if (worksheett.Cells[addressA].Text == worksheett.Cells[addressJ].Value.ToString() && f < 25)
                            {
                                worksheet.Cells[$"A{f}:C{f}"].Value = worksheett.Cells[$"A{j}:C{j}"].Value;
                                worksheet.Cells[$"E{f}:I{f}"].Value = worksheett.Cells[$"D{j}:H{j}"].Value;
                                worksheet.Cells[$"J{f}:L{f}"].Value = worksheett.Cells[$"J{i}:L{i}"].Value;
                                worksheet.Cells[$"N{f}:R{f}"].Value = worksheett.Cells[$"M{i}:Q{i}"].Value;


                                DateTime dateValue = DateTime.Parse(worksheet.Cells[$"A{f}"].Value.ToString());
                                DateTime dateValuee = DateTime.Parse(worksheet.Cells[$"J{f}"].Value.ToString());
                                worksheet.Cells[$"A{f}"].Value = dateValue.ToString("HH:mm") + " 12/13/2024";
                                worksheet.Cells[$"J{f}"].Value = dateValuee.ToString("HH:mm") + " 12/13/2024";


                                worksheet.Cells[$"D{f}"].Value = ((double)worksheet.Cells[$"C{f}"].Value - (double)worksheet.Cells[$"B{f}"].Value).ToString();
                                worksheet.Cells[$"M{f}"].Value = ((double)worksheet.Cells[$"L{f}"].Value - (double)worksheet.Cells[$"K{f}"].Value).ToString();
                                f++;
                                Response.Write(worksheet.Cells[$"A{f}"].Value);
                                Response.Write(worksheet.Cells[$"A{f}"].Text);
                               
                            }

                        }

                    }

                }
               
                FileInfo fileInfo = new FileInfo(filePath);
                package.SaveAs(fileInfo);
            }

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AppendHeader("Content-Disposition", "attachment; filename=my.xlsx");
            Response.TransmitFile(filePath);
            Response.Flush();
            Response.End();
        }

       
    }
    }               