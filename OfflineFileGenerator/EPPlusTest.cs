using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfflineFileGenerator
{
    public class EPPlusTest
    {
        public void GetCellValue()
        {
            var color = System.Drawing.Color.FromName("LightGray");
            var testfile = @"D:\PRWork\Plumriver-Customers\BlackDiamond\BDEL\newoffline\F15_APPAREL_IGD_BDEL_Thailand_preseason_test.xlsx";
            using (ExcelPackage package = new ExcelPackage(new FileInfo(testfile)))
            {
                var sheet = package.Workbook.Worksheets["TabularOfflineOrderForm"];
                var cell = sheet.Cells["N8"];
                cell.Style.Numberformat.Format = "mm-dd-yy";
                cell = sheet.Cells["J8"];
                cell.Style.Numberformat.Format = "mm-dd-yy";
                package.Save();
            }


            //var existFile = @"D:\APPSetUp\FOOTWEAR FALL HOLIDAY 14_GB_GBP_Images.xlsm";
            var newFile = @"D:\APPSetUp\FOOTWEAR FALL HOLIDAY 14_GB_GBP_Images001.xlsm";
            InsertWorksheet("copyopenxml", newFile, 0);
            
            //File.Copy(existFile, newFile, true);
            using (ExcelPackage package = new ExcelPackage(new FileInfo(newFile)))
            {
                var dtSheet = package.Workbook.Worksheets.Copy("Order Template", "CopyTest01");
                package.Save();
            }

            var existingFile = @"D:\PRWork\Plumriver-Customers\Reebok\Offline\TabularOfflineOrderForm.xlsm";
            var newfile = @"D:\PRWork\Plumriver-Customers\Reebok\Offline\TabularOfflineOrderForm_00123.xlsm";
            File.Copy(existingFile, newfile, true);
            using (ExcelPackage package = new ExcelPackage(new FileInfo(newfile)))
            {
                //ExcelWorksheet worksheet = package.Workbook.Worksheets["CategoryTemplate"];
                //var cellvalue = worksheet.Cells["F8"].Value;

                ExcelWorksheet sheet = package.Workbook.Worksheets.Copy("CategoryTemplate", "abcdefg01");
                //sheet.Hidden = eWorkSheetHidden.Visible;

                package.Save();
            }
        }

        private void InsertWorksheet(string sheetName, string filePath, int idx)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
            {
                //get template
                WorksheetPart clonedSheet = null;
                var tempSheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == string.Format("Order Template", idx));

                var tempPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(tempSheets.First().Id.Value);
                // Add a blank WorksheetPart.
                //WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                //newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                //copy template to workbook
                var copySheet = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
                WorkbookPart copyWorkbookPart = copySheet.AddWorkbookPart();
                WorksheetPart copyWorksheetPart = copyWorkbookPart.AddPart<WorksheetPart>(tempPart);
                clonedSheet = spreadSheet.WorkbookPart.AddPart<WorksheetPart>(copyWorksheetPart);

                Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(clonedSheet);

                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
            }
        }

        public string HashPassword(string password)
        {
            byte[] passwordCharacters = System.Text.Encoding.ASCII.GetBytes(password);
            int hash = 0;
            if (passwordCharacters.Length > 0)
            {
                int charIndex = passwordCharacters.Length;

                while (charIndex-- > 0)
                {
                    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                    hash ^= passwordCharacters[charIndex];
                }
                // Main difference from spec, also hash with charcount
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters.Length;
                hash ^= (0x8000 | ('N' << 8) | 'K');
            }

            return Convert.ToString(hash, 16).ToUpperInvariant();
        }

        public void GetWorkbook()
        {
            try
            {
                var filePath = @"D:\PRWork\Plumriver-Customers\Reebok\Offline\enhancement\Pro_protect.xlsm";
                //using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
                //{
                //}

                using (ExcelPackage pck = new ExcelPackage(new FileInfo(filePath), "Plumriver"))
                {
                }
            }
            catch (Exception ex)
            {
                var a = ex.Message;
            }
            
        }
    }
}
