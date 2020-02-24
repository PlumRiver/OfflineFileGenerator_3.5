using System;
using System.IO;
using System.Collections.Generic;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.Text;
using System.Linq;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace OfflineFileGenerator
{
    public class OfflineGenerator
    {
        //This is the maximum department hierarchy levels in the generated file
        //Used to set global variables in this module
        //private int maxcols = 0;

        public bool OrderShipDate = false;

        public OfflineGenerator()
        {
        }

        private void WriteToLog(string msg)
        {
            //string templatefile = ConfigurationManager.AppSettings["templatefile"].ToString();
            var logFile = "OfflineLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            string savefolder = ConfigurationManager.AppSettings["savefolder"].ToString();
            StreamWriter sw = new StreamWriter(savefolder + logFile, true);
            sw.WriteLine(msg);
            sw.Close();
        }

        private string GetColumnLetter(int colnumber)
        {
            string colletter = "";
            switch (colnumber)
            {
                case 1:
                    colletter = "A";
                    break;
                case 2:
                    colletter = "B";
                    break;
                case 3:
                    colletter = "C";
                    break;
                case 4:
                    colletter = "D";
                    break;
                case 5:
                    colletter = "E";
                    break;
                case 6:
                    colletter = "F";
                    break;
                case 7:
                    colletter = "G";
                    break;
                case 8:
                    colletter = "H";
                    break;
                case 9:
                    colletter = "I";
                    break;
                case 10:
                    colletter = "J";
                    break;
                case 11:
                    colletter = "K";
                    break;
                case 12:
                    colletter = "L";
                    break;
                case 13:
                    colletter = "M";
                    break;
                case 14:
                    colletter = "N";
                    break;
            }
            return colletter;
        }

        public bool isPricingInclude(string templatePath, string variableName)
        {
            try
            {
                FileStream tempFile = new FileStream(templatePath, FileMode.Open, FileAccess.Read);
                HSSFWorkbook tempBook = new HSSFWorkbook(tempFile);
                HSSFSheet sheet = (HSSFSheet)tempBook.GetSheet("GlobalVariables");
                if (sheet == null)
                    WriteToLog("isPricingInclude -- sheet is null");
                HSSFRow rowa = (HSSFRow)sheet.GetRow(13);
                if (rowa != null)
                {
                    if (rowa.GetCell(0) != null && rowa.GetCell(0).StringCellValue.Trim().ToLower() == "includepricing")
                    {
                        if (rowa.GetCell(1) != null && rowa.GetCell(1).StringCellValue.Trim().ToLower() == "n")
                            return false;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToLog(DateTime.Now.ToString() + "--" + ex.Message + " details: " + ex.StackTrace);
                throw ex;
            }
            finally
            {
            }
            return true;
        }

        public void CopyForm(string source, string target, string saveas, int maxcols, SqlDataReader reader)
        {
            List<DateTime> DateSheetCollection = new List<DateTime>();
            bool TopDeptGroupStyle = (ConfigurationManager.AppSettings["TopDeptGroupStyle"] ?? "0") == "1";
            if (TopDeptGroupStyle)
            {
                try
                {
                    reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'DateSheetCollection'";
                    if (reader.GetSchemaTable().DefaultView.Count > 0 && reader["DateSheetCollection"] != null)
                    {
                        string[] datelist = reader["DateSheetCollection"].ToString().Split(new char[] { ',', ';', '|' });
                        foreach (string str in datelist)
                        {
                            DateTime dt;
                            if (DateTime.TryParse(str, out dt))
                            {
                                DateSheetCollection.Add(dt);
                            }
                        }
                    }
                }
                catch
                { }
            }

            try
            {
                WriteToLog(saveas);
                //string savefolder = ConfigurationManager.AppSettings["savefolder"].ToString();
                string uploadshipdate = ConfigurationManager.AppSettings["uploadshipdate"].ToString();
                string ExclusiveStyles = ConfigurationManager.AppSettings["ExclusiveStyles"] ?? "0";
                //Open a copy of the template
                //OpenFile(target, "", "target");
                FileStream sourceFileStream = new FileStream(source, FileMode.Open, FileAccess.Read);
                NPOI.POIFS.FileSystem.POIFSFileSystem sourceFile = new NPOI.POIFS.FileSystem.POIFSFileSystem(sourceFileStream);
                //FileStream sourceFile = new FileStream(source, FileMode.Open, FileAccess.Read);
                HSSFWorkbook sourceBook = new HSSFWorkbook(sourceFile);

                FileStream tempFileStream = new FileStream(target, FileMode.Open, FileAccess.Read);
                NPOI.POIFS.FileSystem.POIFSFileSystem tempFile = new NPOI.POIFS.FileSystem.POIFSFileSystem(tempFileStream);
                HSSFWorkbook tempBook = new HSSFWorkbook(tempFile);

                //Migrate the generated sheets into the copy of the template
                WriteToLog("begin move");
                var sheetNum = tempBook.NumberOfSheets;
                //tempBook.CreateSheet("WholesalePrice");
                //tempBook..CopySheets((HSSFSheet)sourceBook.GetSheetAt(2), (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                ((HSSFSheet)sourceBook.GetSheetAt(2)).CopyTo(tempBook, "WholesalePrice", true, true);
                sheetNum++;
                int tmpLogoIdx = 0;
                //HSSFSheet ordTmpSheet = (HSSFSheet)tempBook.CreateSheet("Order Template");
                //tempBook.CopySheets((HSSFSheet)sourceBook.GetSheetAt(0), (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                ((HSSFSheet)sourceBook.GetSheetAt(0)).CopyTo(tempBook, "Order Template", true, true);
                HSSFSheet ordTmpSheet = (HSSFSheet)tempBook.GetSheetAt(sheetNum);
                sheetNum++;
                List<HSSFSheet> dtSheets = new List<HSSFSheet>();
                for (int i = 0; i < DateSheetCollection.Count; i++)
                {
                    var logoPath = ConfigurationManager.AppSettings["LogoPath"];
                    byte[] bytes = System.IO.File.ReadAllBytes(logoPath);
                    int LogoIdx = tempBook.AddPicture(bytes, PictureType.JPEG);

                    //HSSFSheet dtSheet = (HSSFSheet)tempBook.CreateSheet(string.Format("{0:ddMMMyyyy}", DateSheetCollection[i]));
                    //tempBook.CopySheets((HSSFSheet)sourceBook.GetSheetAt(0), (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                    ((HSSFSheet)sourceBook.GetSheetAt(0)).CopyTo(tempBook, string.Format("{0:ddMMMyyyy}", DateSheetCollection[i]), true, true);
                    HSSFSheet dtSheet = (HSSFSheet)tempBook.GetSheetAt(sheetNum);

                    HSSFPatriarch patriarch = (HSSFPatriarch)dtSheet.CreateDrawingPatriarch();
                    HSSFClientAnchor anchor = new HSSFClientAnchor(50 * 1, 50 * 1, 500, 100, 0, 0, 1, 2);
                    HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, LogoIdx);

                    HSSFRow r = (HSSFRow)dtSheet.GetRow(0);
                    var dtcell = r.GetCell(0) == null ? (HSSFCell)r.CreateCell(0) : (HSSFCell)r.GetCell(0);
                    dtcell.SetCellValue(DateSheetCollection[i]);

                    dtSheets.Add(dtSheet);
                    sheetNum++;
                }
                //tempBook.CreateSheet("WebSKU");
                //tempBook.CopySheets((HSSFSheet)sourceBook.GetSheetAt(1), (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                ((HSSFSheet)sourceBook.GetSheetAt(1)).CopyTo(tempBook, "WebSKU", true, true);
                sheetNum++;
                //ISheet skuSheet = tempBook.CreateSheet("CellFormats");
                //tempBook.CopySheets((HSSFSheet)sourceBook.GetSheetAt(3), (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                ((HSSFSheet)sourceBook.GetSheetAt(3)).CopyTo(tempBook, "CellFormats", true, true);
                ISheet skuSheet = tempBook.GetSheetAt(sheetNum);
                sheetNum++;
                //HSSFSheet atpSheet = (HSSFSheet)tempBook.CreateSheet("ATPDate");
                //tempBook.CopySheets((HSSFSheet)sourceBook.GetSheetAt(4), (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                ((HSSFSheet)sourceBook.GetSheetAt(4)).CopyTo(tempBook, "ATPDate", true, true);
                ISheet atpSheet = tempBook.GetSheetAt(sheetNum);
                sheetNum++;
                var upcSheetName = ConfigurationManager.AppSettings["CatalogUPCSheetName"];
                var lastsheet = 4;
                int exclusiveStylesSheetIndex = 5;
                if (!string.IsNullOrEmpty(upcSheetName))
                {
                    exclusiveStylesSheetIndex = 6;
                    //tempBook.CreateSheet(upcSheetName);
                    //tempBook.CopySheets((HSSFSheet)sourceBook.GetSheetAt(5), (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                    ((HSSFSheet)sourceBook.GetSheetAt(5)).CopyTo(tempBook, upcSheetName, true, true);
                    sheetNum++; lastsheet++;
                }

                if (!string.IsNullOrEmpty(ExclusiveStyles) && ExclusiveStyles == "1")
                {
                    //tempBook.CreateSheet("ExclusiveStyles");
                    //tempBook.CopySheets((HSSFSheet)sourceBook.GetSheetAt(exclusiveStylesSheetIndex), (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                    ((HSSFSheet)sourceBook.GetSheetAt(exclusiveStylesSheetIndex)).CopyTo(tempBook, "ExclusiveStyles", true, true);
                    tempBook.SetSheetHidden(sheetNum, SheetState.HIDDEN);
                    sheetNum++;lastsheet++;
                }
                // Solo => Add validation and lock for cell on server side.
                PresetForProductSheet(ordTmpSheet, skuSheet);

                if (dtSheets.Count > 0)
                {
                    tempBook.SetSheetHidden(0, SheetState.HIDDEN);

                    string[] ATPLevelBackColors = (ConfigurationManager.AppSettings["ATPLevelBackColors"] ?? "").Split(new char[] { ',', '|' }, StringSplitOptions.RemoveEmptyEntries);
                    short LockedCellBGColor = short.Parse(ConfigurationManager.AppSettings["LockedCellBGColor"] ?? "22");
                    bool allThinBorder = ConfigurationManager.AppSettings["ThinBorder"] == null ? false : ConfigurationManager.AppSettings["ThinBorder"].ToString() == "1" ? true : false;
                    for (int i = 0; i < dtSheets.Count; i++)
                    {
                        HSSFSheet dtSheet = dtSheets[i];
                        Dictionary<int, ICellStyle> colors = new Dictionary<int, ICellStyle>();
                        ICellStyle lockedStyle = tempBook.CreateCellStyle();
                        lockedStyle.FillForegroundColor = LockedCellBGColor;
                        lockedStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
                        lockedStyle.BorderBottom = lockedStyle.BorderLeft = lockedStyle.BorderRight = lockedStyle.BorderTop = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM;
                        colors.Add(0, lockedStyle);
                        for (int j = 0; j < ATPLevelBackColors.Length / 2; j++)
                        {
                            ICellStyle qtyStyle = tempBook.CreateCellStyle();
                            qtyStyle.IsLocked = false;
                            qtyStyle.FillForegroundColor = short.Parse(ATPLevelBackColors[j * 2 + 1]);
                            qtyStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
                            qtyStyle.BorderBottom = qtyStyle.BorderLeft = qtyStyle.BorderRight = qtyStyle.BorderTop = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM;
                            colors.Add(int.Parse(ATPLevelBackColors[j * 2]), qtyStyle);
                        }
                        DateTime dt;
                        if (DateTime.TryParse(dtSheet.SheetName, out dt))
                        {
                            PresetForProductSheet(dtSheet, skuSheet, dt, atpSheet, colors);
                        }
                    }
                }

                for (int i = 3; i < 10 + DateSheetCollection.Count; i++)
                {
                    if (i > 6 && i < 7 + DateSheetCollection.Count)
                        continue;
                    tempBook.SetSheetHidden(i, SheetState.HIDDEN);
                }

                //tempBook.SetSheetHidden(3, SheetState.Hidden);
                //tempBook.SetSheetHidden(4, SheetState.Hidden);
                //tempBook.SetSheetHidden(5, SheetState.Hidden);
                //tempBook.SetSheetHidden(6, SheetState.Hidden);
                //tempBook.SetSheetHidden(7, SheetState.Hidden);
                //tempBook.SetSheetHidden(8, SheetState.Hidden);
                //tempBook.SetSheetHidden(9, SheetState.Hidden);

                var sourceBookNum = sourceBook.NumberOfSheets;
                for (int i = lastsheet + 1; i < sourceBookNum; i++)
                {
                    HSSFSheet sht = (HSSFSheet)sourceBook.GetSheetAt(i);
                    for (int k = 0; k < 7;k++ )
                        sht.AutoSizeColumn(k);
                    //tempBook.CreateSheet(sht.SheetName);
                    //tempBook.CopySheets(sht, (HSSFSheet)tempBook.GetSheetAt(sheetNum), true);
                    sht.CopyTo(tempBook, sht.SheetName, true, true);
                    sheetNum++;
                }

                WriteToLog("end move");

                //Insert the hard page breaks for printing
                WriteToLog("begin insert");
                InsertPageBreaks(tempBook, "Order Template");
                if (dtSheets.Count > 0)
                {
                    for (int i = 0; i < dtSheets.Count; i++)
                    {
                        HSSFSheet dtSheet = dtSheets[i];
                        InsertPageBreaks(tempBook, dtSheet.SheetName);
                    }
                }
                WriteToLog("end insert");
                //Set global variables for macro use
                WriteToLog("begin set");
                SetUploadShipDate(tempBook, uploadshipdate, maxcols, reader);
                WriteToLog("end set");
                //Set Cancel After Date
                SetCancelAfterDate(tempBook);
                FileStream orderFile = new FileStream(saveas, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                tempBook.Write(orderFile);

                sourceFileStream.Close();
                tempFileStream.Close();
                orderFile.Close();
                sourceBook = null;
                tempBook = null;
                System.GC.Collect();
                //if (File.Exists(@savefolder + "test.xls"))
                //    File.Delete(@savefolder + "test.xls");
                if (File.Exists(source))
                    File.Delete(source);
            }
            catch (Exception e)
            {
                string msg = e.Message;
                throw e;
            }
        }

        private void PresetForProductSheet(HSSFSheet prdSheet, ISheet skuSheet)
        {
            PresetForProductSheet(prdSheet, skuSheet, null, null, null);
        }

        private void PresetForProductSheet(HSSFSheet prdSheet, ISheet skuSheet, DateTime? sheetDate, ISheet atpSheet, Dictionary<int, ICellStyle> colors)
        {
            int? ThresholdQty = null; 
            if(ConfigurationManager.AppSettings["ThresholdQty"]!=null)
            {
                int qty;
                if(int.TryParse(ConfigurationManager.AppSettings["ThresholdQty"].ToString(), out qty))
                {
                    ThresholdQty = qty;
                }
            }
            for (int i = 1; i <= skuSheet.LastRowNum; i++)
            {
                HSSFRow row = (HSSFRow)skuSheet.GetRow(i);
                string cellPos = ((HSSFCell)row.Cells[0]).ToString();
                string multipleValue = ((HSSFCell)row.Cells[3]).ToString();
                string colIdxstring = string.Empty;
                int colIdx = -1;
                string rowIdxstring = string.Empty;
                int rowIdx = 0;
                for (int j = 0; j < cellPos.Length; j++)
                {
                    int ascii = (int)cellPos[j];
                    if (ascii > 57)
                    {
                        if (colIdx == -1)
                        {
                            colIdx = ascii - 65;
                        }
                        else
                        {
                            colIdx = (colIdx + 1) * 26 + (ascii - 65);
                        }
                        colIdxstring += cellPos[j];
                    }
                    else
                    {
                        rowIdxstring += cellPos[j];
                    }
                }
                rowIdx = int.Parse(rowIdxstring) - 1;
                ICell cell = prdSheet.GetRow(rowIdx).GetCell(colIdx);
                if (cell == null)
                {
                    WriteToLog(string.Format("Cell[{0}] is not created in sheet[{1}]", cellPos, prdSheet.SheetName));
                }
                else
                {
                    if (sheetDate.HasValue)
                    {

                        try
                        {
                            ICell atpCell = atpSheet.GetRow(rowIdx).GetCell(colIdx);
                            string dateValues = (atpCell == null ? string.Empty : atpSheet.GetRow(rowIdx).GetCell(colIdx).StringCellValue);
                            string[] arr = dateValues.Split(new char[] { '|', ',' }, StringSplitOptions.RemoveEmptyEntries);
                            int availbleQty = 0;
                            if (arr.Length == 1)
                            {
                                DateTime atpDate = DateTime.Parse(arr[0]);
                                if (atpDate <= sheetDate.Value)
                                {
                                    availbleQty = 999999;
                                }
                            }
                            else
                            {
                                for (int j = 0; j < arr.Length; j = j + 2)
                                {
                                    DateTime atpDate = DateTime.Parse(arr[j]);
                                    if (atpDate <= sheetDate.Value)
                                    {
                                        availbleQty += int.Parse(arr[j + 1]);
                                    }
                                }
                            }
                            var list = colors.Where(kvp => kvp.Key < availbleQty);
                            if (list.Count() > 0)
                            {
                                int colorIdx = list.LastOrDefault().Key;
                                cell.CellStyle = colors[colorIdx];
                            }
                            if (availbleQty > 0)
                            {
                                cell.CellStyle.IsLocked = false;
                                //cell.SetCellValue(availbleQty.ToString());
                                //if (ThresholdQty.HasValue && availbleQty > ThresholdQty)
                                //{
                                //    HSSFPatriarch patr = (HSSFPatriarch)prdSheet.CreateDrawingPatriarch();
                                //    HSSFComment comment = patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, colIdx, rowIdx, colIdx + 1, rowIdx + 1)) as HSSFComment;
                                //    comment.String = new HSSFRichTextString(string.Format("{0}+", ThresholdQty.Value));
                                //    comment.Author = "Plumriver";
                                //    cell.CellComment = comment;
                                //}
                            }
                        }
                        catch(Exception ex) {
                            WriteToLog(ex.Message);
                        }
                        
                    }
                    else
                    {
                        cell.CellStyle.IsLocked = false;
                    }
                    DVConstraint dvConstraint = DVConstraint.CreateCustomFormulaConstraint(string.Format("(MOD(indirect(address(row(),column())) ,{0})=0)", multipleValue));
                    HSSFDataValidation orderMultipleValidation = new HSSFDataValidation(new CellRangeAddressList(cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex), dvConstraint);
                    orderMultipleValidation.CreateErrorBox("Multiple Value Cell", string.Format("You must enter a multiple of {0} in this cell.", multipleValue).Replace(".00 ", " ").Replace(".0 ", " "));
                    prdSheet.AddValidationData(orderMultipleValidation);
                }
            }
        }

        public void InsertPageBreaks(HSSFWorkbook workbook, string shtName)
        {
            //Find instances of the word "pagebreak" and replace with a page break
            HSSFSheet sheet = (HSSFSheet)workbook.GetSheet(shtName);
            int cellctr = 0;
            HSSFCell range = (HSSFCell)sheet.GetRow(cellctr).GetCell(0);
            while (range == null || range.ToString() != "EOF")
            {
                if (range != null && range.ToString() == "pagebreak")
                {
                    sheet.SetRowBreak(cellctr);
                    range.SetCellValue(string.Empty);
                }
                cellctr++;
                range = (HSSFCell)sheet.GetRow(cellctr).GetCell(0);
            }
            range.SetCellValue(string.Empty);
            range = (HSSFCell)sheet.GetRow(0).GetCell(0);
        }

        private void SetCellValue(HSSFSheet sheet, int rowIndex, int cellIndex, string value)
        {
            var row = sheet.GetRow(rowIndex) == null ? (HSSFRow)sheet.CreateRow(rowIndex) : (HSSFRow)sheet.GetRow(rowIndex);
            var cell = row.GetCell(cellIndex) == null ? (HSSFCell)row.CreateCell(cellIndex) : (HSSFCell)row.GetCell(cellIndex);
            cell.SetCellValue(value);
        }

        public void SetUploadShipDate(HSSFWorkbook workbook, string uploadshipdate, int maxcols, SqlDataReader reader)
        {
            bool multicolumndeptlabel = (ConfigurationManager.AppSettings["multicolumndeptlabel"].ToString().ToUpper() == "TRUE" ? true : false);

            //Set global variables in the spreadsheet for use by macros
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            HSSFSheet sheet = (HSSFSheet)workbook.GetSheet("GlobalVariables");

            SetCellValue(sheet, 10, 1, uploadshipdate.ToUpper());
            //HSSFCell range = (HSSFCell)sheet.GetRow(10).GetCell(1);//("B11", Type.Missing);

            string column1heading = System.Configuration.ConfigurationManager.AppSettings["column1heading"].ToString();

            //range = (HSSFCell)sheet.GetRow(5).GetCell(1);//.get_Range("B6", Type.Missing);
            SetCellValue(sheet, 5, 1, column1heading);

            //Totals column
            //range = (HSSFCell)sheet.GetRow(4).GetCell(1);//.get_Range("B5", Type.Missing);
            if (multicolumndeptlabel)
                SetCellValue(sheet, 4, 1, GetColumnLetter(maxcols + 5));
            else
                SetCellValue(sheet, 4, 1, "F");

            //Description column
            //range = (HSSFCell)sheet.GetRow(7).GetCell(1);//.get_Range("B8", Type.Missing);
            if (multicolumndeptlabel)
                SetCellValue(sheet, 7, 1, GetColumnLetter(maxcols + 2));
            else
                SetCellValue(sheet, 7, 1, "C");

            //Dimension (product) column
            //range = (HSSFCell)sheet.GetRow(6).GetCell(1);//.get_Range("B7", Type.Missing);
            if (multicolumndeptlabel)
                SetCellValue(sheet, 6, 1, GetColumnLetter(maxcols + 7));
            else
                SetCellValue(sheet, 6, 1, "H");

            //Are all dept descriptions in first header cell
            //range = (HSSFCell)sheet.GetRow(12).GetCell(1);//.get_Range("B13", Type.Missing);
            if (multicolumndeptlabel)
                SetCellValue(sheet, 12, 1, "N");
            else
                SetCellValue(sheet, 12, 1, "Y");

            //range = (HSSFCell)sheet.GetRow(99).GetCell(1);//.get_Range("B100", Type.Missing);
            SetCellValue(sheet, 99, 1, reader["catalog"].ToString());

            //range = (HSSFCell)sheet.GetRow(100).GetCell(1);//.get_Range("B101", Type.Missing);
            SetCellValue(sheet, 100, 1, reader["catalogname"].ToString());

            reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'CancelDefaultDays'";
            if (reader.GetSchemaTable().DefaultView.Count > 0 && reader["CancelDefaultDays"] != null)
            {
                SetCellValue(sheet, 101, 1, reader["CancelDefaultDays"].ToString());
            }

            reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'SafetyStockQty'";
            if (reader.GetSchemaTable().DefaultView.Count > 0)
            {
                var SafetyStockQty = reader["SafetyStockQty"] == null ? string.Empty : reader["SafetyStockQty"].ToString();
                SetCellValue(sheet, 98, 1, SafetyStockQty);
            }

            SetCellValue(sheet, 18, 1, OrderShipDate ? "1" : "0");

            App app = new App();
            string isEnabled = app.GetB2BSetting("ShippingPage", "DisplayShipVia");
            if (isEnabled != "0")
            {
                DataSet shipmethods = app.GetAllShipMethod();
                if (shipmethods.Tables.Count > 0 && shipmethods.Tables[0].Rows.Count > 0)
                {
                    //range = (HSSFCell)sheet.GetRow(198).GetCell(1);//.get_Range("B199", Type.Missing);
                    SetCellValue(sheet, 198, 1, shipmethods.Tables[0].Rows.Count.ToString());
                    int i = 0;
                    foreach (DataRow row in shipmethods.Tables[0].Rows)
                    {
                        string shipviades = row["description"].ToString();
                        string shipviacode = row["code"].ToString();

                        //range = (HSSFCell)sheet.GetRow(199 + i).GetCell(1);//.get_Range("B" + (200 + i).ToString(), Type.Missing);
                        SetCellValue(sheet, 199 + i, 1, shipviades);

                        //range = (HSSFCell)sheet.GetRow(199 + i).GetCell(2);//.get_Range("C" + (200 + i).ToString(), Type.Missing);
                        SetCellValue(sheet, 199 + i, 2, shipviacode);

                        i++;
                    }
                }
            }
            else
            {
                HSSFSheet sheet1 = (HSSFSheet)workbook.GetSheet("Order Creation");
                for (int i = 6; i < 20; i++)
                {
                    HSSFCell rg = (HSSFCell)sheet1.GetRow(i - 1).GetCell(0);//.get_Range(string.Format("A{0}", i), Type.Missing);
                    if (rg != null && rg.StringCellValue == "ShipMethod")
                    {
                        //rg.EntireRow.Hidden = true;
                        var rowStyle = workbook.CreateCellStyle();
                        rowStyle.IsHidden = true;
                        rg.Row.RowStyle = rowStyle;
                        rg.Row.ZeroHeight = true;
                        rg.Row.HeightInPoints = 0;
                        break;
                    }
                }
            }
        }

        public void SetCancelAfterDate(HSSFWorkbook workbook)
        {
            if (ConfigurationManager.AppSettings["ShipCancelAfterDate"] != null && ConfigurationManager.AppSettings["ShipCancelAfterDate"].ToString() == "0")
            {
                HSSFSheet sheet1 = (HSSFSheet)workbook.GetSheet("Order Creation");
                for (int i = 6; i < 20; i++)
                {
                    HSSFCell rg = (HSSFCell)sheet1.GetRow(i - 1).GetCell(0);//.get_Range(string.Format("A{0}", i), Type.Missing);
                    if (rg != null && rg.StringCellValue == "Cancel After *")
                    {
                        //rg.EntireRow.Hidden = true;
                        var rowStyle = workbook.CreateCellStyle();
                        rowStyle.IsHidden = true;
                        rg.Row.RowStyle = rowStyle;
                        rg.Row.ZeroHeight = true;
                        rg.Row.HeightInPoints = 0;
                        break;
                    }
                }
            }
        }
    }

    public static class NPOIUntilityExt  // for NPOI Version 1.2.4
    {
        public static ISheet CopyTo(this ISheet sourceSheet, HSSFWorkbook targetBook, string targetSheetName, bool includeStyle, bool includeFormula)
        {
            ISheet sht = targetBook.CreateSheet(targetSheetName);
            targetBook.CopySheets((HSSFSheet)sourceSheet, (HSSFSheet)sht, includeStyle);

            return sht;
        }
    }
}
