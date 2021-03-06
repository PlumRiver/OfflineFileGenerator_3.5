﻿using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Linq;
using System.Linq.Expressions;
using System.Xml.Serialization;

using System.Drawing;
using System.Drawing.Imaging;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;

namespace OfflineFileGenerator
{
    public partial class App
    {
        int? CancelAfterDays = null;
        public string Currency = string.Empty;
        public bool SupportATP = true;
        OLOFSetting OLOFSettings = null;

        public string PriceType { get; set; }

        string[] ATPLevelBackColors = (ConfigurationManager.AppSettings["ATPLevelBackColors"] ?? "0|Gray").Split(new char[] { ',', '|' }, StringSplitOptions.RemoveEmptyEntries);
        Dictionary<int, Color> colors = new Dictionary<int, Color>();

        public int GenerateStandardXLSMData(string filepath, string savedirectory, SqlDataReader reader)
        {
            string soldto = reader["SoldTo"].ToString();
            string catalog = reader["Catalog"].ToString();
            string pricecode = reader["PriceCode"].ToString();
            this.CatalogName = reader["CatalogName"].ToString();
            reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'ATPFlag'");
            if (reader.GetSchemaTable().DefaultView.Count > 0)
            {
                SupportATP = reader["ATPFlag"] == null ? true : (reader["ATPFlag"].ToString() == "0" ? false : true);
            }

            if (colors.Count == 0)
            {
                for (int j = 0; j < ATPLevelBackColors.Length / 2; j++)
                {
                    colors.Add(int.Parse(ATPLevelBackColors[j * 2]), Color.FromName(ATPLevelBackColors[j * 2 + 1]));
                }
            }

            //Make sure data doesn't carry across multiple catalogs
            departments.Clear();
            cols.Clear();
            multiples.Clear();
            locked.Clear();
            unlocked.Clear();

            //now rename and make sure the file name with suffix ".xlsm".
            FileInfo newFile = new FileInfo(filepath);
            if (newFile.Extension != ".xlsm")
                newFile = new FileInfo(filepath.Replace(newFile.Extension, ".xlsm"));

            string directroy = newFile.DirectoryName;
            string filename = newFile.Name;
            string tmpFile = Path.Combine(directroy, string.Format("tmp.{0:HHmmss}.xlsm",DateTime.Now));
            if (File.Exists(tmpFile))
            {
                File.Delete(tmpFile);
            }

            string tmpsheet = ConfigurationManager.AppSettings["templatefile"] ?? string.Empty;
            FileInfo tmpXlsm = new FileInfo(tmpsheet);

            newFile = new FileInfo(tmpFile);

            try
            {
                File.Copy(tmpsheet, tmpFile, true);
                using (ExcelPackage pck = new ExcelPackage(new FileInfo(tmpFile)))
                //using (ExcelPackage pck = new ExcelPackage(newFile, tmpXlsm))
                {
                    int cntSheet = pck.Workbook.Worksheets.Count;

                    // -----------------------------------------------
                    //  Generate Order Template Worksheet
                    // -----------------------------------------------
                    WriteToLog("Generate Order Template Sheet begin:" + DateTime.Now.TimeOfDay);
                    this.GenerateXLSMOrderTemplate(pck, soldto, catalog, pricecode, savedirectory, reader);
                    WriteToLog("Generate Order Template Sheet end:" + DateTime.Now.TimeOfDay);

                    Regex reg = new Regex(@"Sheet[1-3]");
                    for (int i = 1; i <= cntSheet; i++)
                    {
                        if (reg.IsMatch(pck.Workbook.Worksheets[i].Name))
                            pck.Workbook.Worksheets[i].Hidden = eWorkSheetHidden.Hidden;
                        pck.Workbook.Worksheets.MoveToEnd(i);
                    }

                    //This saves the XML version of the generated spreadsheet for use by the next
                    //step in the process
                    //pck.Workbook.CreateVBAProject();
                    //Password protect your code
                    //pck.Workbook.VbaProject.Protection.SetPassword("Plumriver");

                    //Now add some code to update the text of the shape...
                    if (System.IO.File.Exists(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\Module.vba")))
                    {
                        var module = pck.Workbook.VbaProject.Modules.AddModule("PRUtilities");
                        module.Code = File.ReadAllText(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\Module.vba"));
                    }

                    var sb = new StringBuilder();
                    if (System.IO.File.Exists(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\VBA.vba")))
                    {
                        sb.Append(File.ReadAllText(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\VBA.vba")));
                    }
                    else
                    {
                        sb.AppendLine("Private Sub Workbook_Open()");
                        sb.AppendLine("    'Solo Created");
                        sb.AppendLine("End Sub");
                    }
                    pck.Workbook.CodeModule.Code = sb.ToString();

                    WriteToLog("Template data sheet is gernerated");
                    pck.Save();
                }

                if (File.Exists(tmpFile))
                {
                    /*
                     0: only non-image
                     1: both , [default]
                     2: only image
                     */
                    int FormStyles = int.Parse(ConfigurationManager.AppSettings["FormStyles"] ?? "1");  
                    string targetName = Path.Combine(directroy, filename);

                    if (FormStyles <= 1)
                    {
                        try
                        {
                            File.Copy(tmpFile, targetName);
                            using (ExcelPackage targepck = new ExcelPackage(new FileInfo(targetName)))
                            {
                                GenerateDateSheet(targepck, reader, false);
                                targepck.Save();
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteToLog("Fail to Generate DateSheet:" + ex.Message);

                            if (File.Exists(targetName))
                            {
                                File.Delete(targetName);
                            }
                        }
                    }

                    //for TOMS, it's configured as 2.
                    WriteToLog("Generate Order Date Sheets begin:" + DateTime.Now.TimeOfDay);
                    if (FormStyles >= 1)
                    {
                        reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'ImageFileName'";
                        if (reader.GetSchemaTable().DefaultView.Count > 0 && reader["ImageFileName"] != null)
                        {
                            string targetImgName = Path.Combine(directroy, CleanFileName(reader["ImageFileName"].ToString()));
                            try
                            {
                                FileInfo imgFile = new FileInfo(targetImgName);
                                if (imgFile.Extension != ".xlsm")
                                    targetImgName = targetImgName.Replace(imgFile.Extension, ".xlsm");

                                File.Copy(tmpFile, targetImgName);
                                using (ExcelPackage targeImgpck = new ExcelPackage(new FileInfo(targetImgName)))
                                {
                                    GenerateDateSheet(targeImgpck, reader, true);
                                    targeImgpck.Save();
                                }
                            }
                            catch (Exception ex)
                            {
                                WriteToLog("Error:" + ex.Message);
                                WriteToLog(ex.StackTrace);

                                if (File.Exists(targetImgName))
                                {
                                    File.Delete(targetImgName);
                                }
                            }
                        }
                    }
                    WriteToLog("Generate Order Date Sheets end:" + DateTime.Now.TimeOfDay);
                }
            }
            catch(Exception ex) {
                WriteToLog("Error:" + ex.Message);
                WriteToLog("Error:" + ex.StackTrace);
            }
            finally
            {
                if (File.Exists(tmpFile))
                {
                    File.Delete(tmpFile);
                }
            }

            System.IO.Directory.GetFiles(System.Windows.Forms.Application.StartupPath, "*.xml").ToList().ForEach(x => File.Delete(x));
            return -1;

        }

        private int GenerateXLSMOrderTemplate(ExcelPackage XLSMPck, string soldto, string catalog, string pricecode, string savedirectory, SqlDataReader xReader)
        {
            var colMultiple = 5; //50;
            //Pull values out of the config file
            int column1width = Convert.ToInt32(ConfigurationManager.AppSettings["column1width"].ToString()) / colMultiple;
            int column2width = Convert.ToInt32(ConfigurationManager.AppSettings["column2width"].ToString()) / colMultiple;
            int column3width = Convert.ToInt32(ConfigurationManager.AppSettings["column3width"].ToString()) / colMultiple;
            int column4width = Convert.ToInt32(ConfigurationManager.AppSettings["column4width"].ToString()) / colMultiple;
            int column5width = Convert.ToInt32(ConfigurationManager.AppSettings["column5width"].ToString()) / colMultiple;
            int column6width = Convert.ToInt32(ConfigurationManager.AppSettings["column6width"].ToString()) / colMultiple;
            int column7width = Convert.ToInt32(ConfigurationManager.AppSettings["column7width"].ToString()) / colMultiple;
            int column8width = Convert.ToInt32(ConfigurationManager.AppSettings["column8width"].ToString()) / colMultiple;
            int column9width = Convert.ToInt32(ConfigurationManager.AppSettings["column9width"].ToString()) / colMultiple;
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            int gridatlevel = Convert.ToInt32(ConfigurationManager.AppSettings["gridatlevel"].ToString());

            //Get catalog data from the database
            //int iret = PrepareData(soldto, catalog, pricecode, savedirectory);
            //if (iret == -1)
            //    return iret;
            WriteToLog("PrepareAllProductData Get Product XML from DB begin:" + soldto + "|" + catalog + "|" + pricecode + "|" + DateTime.Now.TimeOfDay);
            if (!PrepareAllProductData(soldto, catalog, pricecode, savedirectory)) return -1;
            WriteToLog("PrepareAllProductData Get Product XML from DB end:" + DateTime.Now.TimeOfDay);

            XmlReader reader = GetDepartments2(catalog, pricecode);
            reader.Read();
            reader.MoveToNextAttribute();
            maxcols = Convert.ToInt32(reader.Value);
            
            #region "Main tabs"
            var sheet = XLSMPck.Workbook.Worksheets.Add("Order Template");
            sheet.Cells.Style.Font.SetFromFont("".ToFont());
            //SKU tab - hidden - used to create sku column in upload tab
            var skusheet = XLSMPck.Workbook.Worksheets.Add("WebSKU"); skusheet.Hidden = eWorkSheetHidden.Hidden;
            //Wholesale price tab - hidden - used by price calculation formulas
            var pricesheet = XLSMPck.Workbook.Worksheets.Add("WholesalePrice"); pricesheet.Hidden = eWorkSheetHidden.Hidden;
            //Lists of cells to be locked and have order multiple restrictions placed
            var formatcellssheet = XLSMPck.Workbook.Worksheets.Add("CellFormats"); formatcellssheet.Hidden = eWorkSheetHidden.Hidden;
            //ATPDATE tab
            var atpsheet = XLSMPck.Workbook.Worksheets.Add("ATPDate"); atpsheet.Hidden = eWorkSheetHidden.Hidden;
            var upcSheetName = ConfigurationManager.AppSettings["CatalogUPCSheetName"];
            var upcsheet = string.IsNullOrEmpty(upcSheetName) ? null : XLSMPck.Workbook.Worksheets.Add(upcSheetName);
            //Lists of cells to be locked and have order multiple restrictions placed
            var OrderMultiplesheet = XLSMPck.Workbook.Worksheets.Add("OrderMultiple"); OrderMultiplesheet.Hidden = eWorkSheetHidden.Hidden;
            var soldtoshiptosheet = XLSMPck.Workbook.Worksheets.Add(ConfigurationManager.AppSettings["CustomerSheetName"] ?? "SoldToShipTo"); soldtoshiptosheet.Hidden = eWorkSheetHidden.Hidden;
            var globalsheet = XLSMPck.Workbook.Worksheets.Add("GlobalVariables");
            #endregion
            
            AddFormatCellsHeader(formatcellssheet);

            #region "Init Width"
            int colctr = 0;
            if (multicolumndeptlabel)
            {
                for (colctr = 0; colctr < maxcols; colctr++)
                {
                    SetColumnWidth(sheet, column1width, skusheet, pricesheet, colctr);
                }
            }
            else
            {
                SetColumnWidth(sheet, column1width, skusheet, pricesheet, colctr);
            }
            colctr++;

            SetColumnWidth(sheet, column2width, skusheet, pricesheet, colctr++);

            SetColumnWidth(sheet, column3width, skusheet, pricesheet, colctr++);

            SetColumnWidth(sheet, column4width, skusheet, pricesheet, colctr++);

            SetColumnWidth(sheet, column5width, skusheet, pricesheet, colctr++);

            SetColumnWidth(sheet, column6width, skusheet, pricesheet, colctr++);

            SetColumnWidth(sheet, column7width, skusheet, pricesheet, colctr++);

            #endregion

            #region "Init ShipInfo"
            int POLength = 50;
            string POLengthStr = GetB2BSetting("ShippingPage", "POLength");
            int.TryParse(POLengthStr, out POLength);

            string configPath = Path.Combine(System.Environment.CurrentDirectory, "StandardXML.config");
            if (File.Exists(configPath))
            {
                WriteToLog("Load Summary from config file " + DateTime.Now.TimeOfDay);
                using (StreamReader str = new StreamReader(configPath))
                {
                    XmlSerializer xSerializer = new XmlSerializer(typeof(OLOFSetting));
                    OLOFSettings = (OLOFSetting)xSerializer.Deserialize(str);
                }
            }
            else
            {   // Default Settings for Toms
                string[][] rows = new string[][]{ // ColIndex, text, Font, Locked
                    new string[]{"2", "26", "3;"+CatalogName+";18pt.bold;1"},
                    new string[]{"5", "20", "2;Sold-To *;Bold;1|3;;Underline;0"},
                    new string[]{"6", "20", "2;Ship-To *;Bold;1|3;;Underline;0"},
                    new string[]{"7", "20", string.Format(@"2;PO #;Bold;1|3;;Underline;0|5;=IF(LEN(C7)>{0},""The max length of PO Number allowed is {0}"","""");red;1", POLength)},  // C7 is hardcoded here
                    new string[]{"8", "20", @"2;Requested Del Date *;Bold;1|3;;Underline.ddmmmyyyy;0|5;=IF(ISNUMBER(C8),"""",IF(C8="""","""","" Invalid date (yyyy-mm-dd)""));#990000;1"},
                    new string[]{"9", "20", @"2;Cancel Date *;Bold;1|3;;Underline.ddmmmyyyy;0|5;=IF(ISNUMBER(C9),IF(C9<C8,""Cancel Date must be greater than Requested Delivery Date"",""""),IF(C9="""","""","" Invalid date (yyyy-mm-dd)""));#990000;1"}
                };

                int i = 0;
                OLOFSettings = new OLOFSetting()
                {
                    Summary = new OLOFSettingLineItems[rows.Length]
                };
                foreach (string[] arr in rows)
                {
                    int j = 0;
                    OLOFSettingLineItems line = new OLOFSettingLineItems();
                    line.RowNumber = int.Parse(arr[0]);
                    line.RowHeight = int.Parse(arr[1]);

                    string[] cellArr = arr[2].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    line.LineItem = new OLOFSettingLineItemsLineItem[cellArr.Length];
                    foreach (string cl in cellArr)
                    {
                        string[] attrs = cl.Split(new char[] { ';' });
                        line.LineItem[j] = new OLOFSettingLineItemsLineItem();
                        line.LineItem[j].columnNumber = int.Parse(attrs[0]);
                        line.LineItem[j].value = attrs[1];
                        line.LineItem[j].columnSpan = 1;
                        line.LineItem[j].style = attrs[2];
                        line.LineItem[j].locked = attrs[3];
                        j++;
                    }
                    OLOFSettings.Summary[i] = line;
                    i++;
                }
            }

            bool? lockedDate = null;
            try
            {
                xReader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'lockedDate'";
                if (xReader.GetSchemaTable().DefaultView.Count > 0 && xReader["lockedDate"] != null)
                {
                    lockedDate = (xReader["lockedDate"].ToString() == "1");
                    WriteToLog("Locked the req ship date and cancel date");
                }
            }
            catch
            { }

            foreach (var line in OLOFSettings.Summary)
            {
                sheet.SetRowHeight(line.RowNumber, line.RowHeight);
                foreach (var item in line.LineItem)
                {
                    item.value = item.value.Replace("{CatalogName}", CatalogName);
                    item.value = item.value.Replace("{POLength}", POLength.ToString());
                    using (ExcelRange cell = sheet.Cells[line.RowNumber, item.columnNumber, line.RowNumber, item.columnNumber + item.columnSpan - 1])
                    {
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.General;
                        cell.Style.Font.SetFromFont(item.style.ToFont("."));

                        foreach (string style in item.style.ToLower().Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            string numberformat = "@";
                            switch (style)
                            {
                                case "underline":
                                    cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    break;
                                case "ddmmmyyyy":
                                    numberformat = "ddmmmyyyy";
                                    break;
                                default:
                                    {
                                        if (style.StartsWith("#"))
                                        {
                                            cell.Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml(style));
                                        }
                                    }
                                    break;
                            }
                            cell.Style.Numberformat.Format = numberformat;
                        }
                        if (lockedDate.HasValue)
                        {
                            if (item.columnName == "ReqDate" || item.columnName == "CancelDate")
                            {
                                if (!lockedDate.Value)
                                    cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                cell.Style.Locked = lockedDate.Value;
                            }
                            else
                                cell.Style.Locked = (item.locked == "1");
                        }
                        else if (item.locked == "0")
                            cell.Style.Locked = false;

                        if (item.value.Length > 0)
                            if (item.value.StartsWith("="))
                                cell.Formula = item.value.Substring(1, item.value.Length - 1);
                            else
                                cell.Value = item.value;

                        if (item.columnSpan > 1)
                            cell.Merge = true;

                    }
                }
            }
            if (OLOFSettings.Settings == null)
                OLOFSettings.Settings = new OLOFSettingAdd[] { }; 

            if (OLOFSettings.Settings.ExistsKey("DisplaySoldtoShipList"))
            {
                DataTable dt = GetCustomerData(soldto, catalog);
                SetCellValue(soldtoshiptosheet, string.Format("A{0}", 1), "SoldTo", false);
                SetCellValue(soldtoshiptosheet, string.Format("B{0}", 1), "ShipTo", false);
                SetCellValue(soldtoshiptosheet, string.Format("C{0}", 1), "SoldtoName", false);
                var SoldToShipToWithShipAddress = string.IsNullOrEmpty(ConfigurationManager.AppSettings["SoldToShipToWithShipAddress"]) ? "0" : ConfigurationManager.AppSettings["SoldToShipToWithShipAddress"];
                if(SoldToShipToWithShipAddress == "1")
                    SetCellValue(soldtoshiptosheet, string.Format("D{0}", 1), "ShipToAddress", false);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];
                    SetCellValue(soldtoshiptosheet, string.Format("A{0}", i + 2), dr["SoldTo"].ToString(), false);
                    SetCellValue(soldtoshiptosheet, string.Format("B{0}", i + 2), dr["ShipTo"].ToString(), false);
                    SetCellValue(soldtoshiptosheet, string.Format("C{0}", i + 2), dr["SoldtoName"].ToString(), false);
                    if(SoldToShipToWithShipAddress == "1")
                        SetCellValue(soldtoshiptosheet, string.Format("D{0}", i + 2), dr["ShipToAddress"].ToString(), false);
                }
            }
            #endregion

            #region Init PONumber Checking
            /*var StandardPONumberValidation = ConfigurationManager.AppSettings["StandardPONumberValidation"];
            if (!string.IsNullOrEmpty(StandardPONumberValidation))
            {
                var StandardPONumberValidationErrorMessage = string.IsNullOrEmpty(ConfigurationManager.AppSettings["StandardPONumberValidationErrorMessage"]) ? "Alphanumeric characters only" : ConfigurationManager.AppSettings["StandardPONumberValidationErrorMessage"];
                var StandardPONumberCell = string.IsNullOrEmpty(ConfigurationManager.AppSettings["StandardPONumberCell"]) ? "D7" : ConfigurationManager.AppSettings["StandardPONumberCell"];
                var validation = sheet.DataValidations.AddCustomValidation(StandardPONumberCell);
                validation.AllowBlank = true;
                validation.ShowInputMessage = true;
                validation.ShowErrorMessage = true;
                validation.Error = StandardPONumberValidationErrorMessage;
                validation.Formula.ExcelFormula = string.Format(StandardPONumberValidation, StandardPONumberCell);
            }*/

            var StandardPONumberValidationMessage = ConfigurationManager.AppSettings["StandardPONumberValidationMessage"];
            if(!string.IsNullOrEmpty(StandardPONumberValidationMessage))
            {
                var StandardPONumberCell = string.IsNullOrEmpty(ConfigurationManager.AppSettings["StandardPONumberCell"]) ? "D7" : ConfigurationManager.AppSettings["StandardPONumberCell"];
                StandardPONumberValidationMessage = string.Format(StandardPONumberValidationMessage, StandardPONumberCell, POLength.ToString());
                var StandardPONumberValidationMessageCell = string.IsNullOrEmpty(ConfigurationManager.AppSettings["StandardPONumberValidationMessageCell"]) ? "H7" : ConfigurationManager.AppSettings["StandardPONumberValidationMessageCell"];
                
                if (StandardPONumberValidationMessage.StartsWith("="))
                {
                    var cell = sheet.Cells[StandardPONumberValidationMessageCell];
                    cell.Formula = StandardPONumberValidationMessage.Substring(1, StandardPONumberValidationMessage.Length - 1);
                }
                else
                    SetCellValue(sheet, StandardPONumberValidationMessageCell, StandardPONumberValidationMessage, false);
            }
            #endregion

            #region "Order Template"
            ListDictionary colpositions = null;
            WriteToLog("CreateGridRows2 begin:" + DateTime.Now.TimeOfDay);
            XElement xroot = XElement.Load(TemplateXMLFilePath);

            string lastRootDeptID = "";
            bool? groupdTitle = null;
            TotalQtyfn = new StringBuilder();
            TotalAmtfn = new StringBuilder();
            StringBuilder SubTotalQtyfn = new StringBuilder();
            StringBuilder SubTotalAmtfn = new StringBuilder();

            var PRDheaders = OLOFSettings.Settings.GetValueByKey("ProductDetailHeaders", "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            var lastDeptColumn = (PRDheaders.Length == 0 ? 3 : (PRDheaders.Length + 1));
            
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.EndElement && reader.Name == "ProductLevel")
                {
                    Dictionary<string, string[]> rootDept = new Dictionary<string, string[]>();
                    //Build a grid
                    XmlReader currentNode = reader.ReadSubtree();
                    int deptid = 0;
                    string deptname = string.Empty, PrepackList = string.Empty;
                    OrderedDictionary depts = GetDeptHierarchy(currentNode, ref deptid, ref deptname, ref rootDept);

                    var sheetRow = sheet.Dimension.End.Row + 1;
                    if (rootDept.Count > 0)
                    {
                        if (rootDept != null && lastRootDeptID != rootDept.FirstOrDefault().Key)
                        {
                            if (SubTotalQtyfn.Length > 0)
                            {
                                var rowIndex = sheet.LastRowNum() + 1;
                                AddCategorySubtotal(sheet, SubTotalQtyfn, SubTotalAmtfn, rowIndex);

                                sheetRow++;
                                rowIndex++;
                                SubTotalQtyfn = new StringBuilder();
                                SubTotalAmtfn = new StringBuilder();
                            }

                            sheetRow++;
                            upcsheet = XLSMPck.Workbook.Worksheets.Add(rootDept.FirstOrDefault().Value[0]);
                            AddCatalogUPCSheetHeader(upcsheet);

                            string[] arr = rootDept.FirstOrDefault().Value;
                            AddBreakRows(sheet, 1);
                            AddTextRow(sheet, sheetRow, 20, new string[][] { new string[] { arr[0], "18pt", "0", "2" } });
                            sheetRow++;
                            AddTextRow(sheet, sheetRow, 15, new string[][] { new string[] { string.Format("{0}: {1}", arr[1], arr[2]),"12pt", "0", lastDeptColumn.ToString(), "0" },
                                                                                               new string[] { "BLACKED OUT = NOT AVAILABLE","Bold", (lastDeptColumn+1).ToString(), "8" } });
                            sheetRow++;
                            PrepackList = arr[3];
                            groupdTitle = true;
                            lastRootDeptID = rootDept.FirstOrDefault().Key;
                        }
                        else
                        {
                            groupdTitle = false;
                        }
                    }

                    CreateGrid2(xroot, depts, deptid, sheet, sheetRow, ref colpositions, soldto, catalog, groupdTitle, PrepackList);

                    int[] borderRange = new int[4];
                    borderRange[0] = sheet.Dimension.End.Row; borderRange[1] = 1;

                    int gId = sheet.Dimension.End.Row;
                    int maxColIdx = 0;
                    try
                    {
                        maxColIdx = CreateGridRows2(depts, sheet, skusheet, pricesheet, atpsheet, upcsheet, OrderMultiplesheet, deptid, deptname, ref colpositions);
                    }
                    catch 
                    { 
                    
                    }
                    TotalQtyfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalQtyCol.IntToMoreChar(), gId, sheet.Dimension.End.Row);
                    TotalAmtfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalAmtCol.IntToMoreChar(), gId, sheet.Dimension.End.Row);
                    SubTotalQtyfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalQtyCol.IntToMoreChar(), gId, sheet.Dimension.End.Row);
                    SubTotalAmtfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalAmtCol.IntToMoreChar(), gId, sheet.Dimension.End.Row);

                    if (upcsheet != null)
                    {
                        upcsheet.Cells.AutoFitColumns();
                    }

                    borderRange[2] = sheet.Dimension.End.Row; borderRange[3] = maxColIdx;
                    for (int row = borderRange[0]; row <= borderRange[2]; row++)
                    {
                        for (int col = borderRange[1]; col < borderRange[3]; col++)
                        {
                            var cl = sheet.Cells[col.IntToMoreChar() + row.ToString()];
                            if (col >= (PRDheaders.Length == 0 ? 8 : (PRDheaders.Length + 6)))
                            {
                                if (!unlocked.Contains(cl.Address))
                                {
                                    cl.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    cl.Style.Fill.BackgroundColor.SetColor(colors[0]);
                                }
                            }
                            cl.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        }
                    }
                }
            }

            if (SubTotalQtyfn.Length > 0)
            {
                var rowIndex = sheet.LastRowNum() + 1;
                AddCategorySubtotal(sheet, SubTotalQtyfn, SubTotalAmtfn, rowIndex);
            }

            if (TotalSection)
            {
                var rowIndex = sheet.LastRowNum() + 2;
                sheet.Row(rowIndex).Height = RowHeight * 1.5;
                sheet.Row(rowIndex + 1).Height = RowHeight * 1.5;
                using (ExcelRange r = sheet.Cells[rowIndex, TotalAmtCol - 1, rowIndex, TotalAmtCol])
                {
                    r.Style.Font.SetFromFont("bold".ToFont());
                    r.Value = ConfigurationManager.AppSettings["TotalQtyLabel"] ?? "TOTAL PAIRS";
                    r.Merge = true;
                    r.AddBorder();
                }

                using (ExcelRange r = sheet.Cells[rowIndex, TotalAmtCol + 1, rowIndex, TotalAmtCol + 2])
                {
                    r.Style.Numberformat.Format = "#,##0.00";
                    r.Formula = "0" + TotalQtyfn.ToString();
                    r.Merge = true;
                    r.AddBorder();
                }

                rowIndex++;

                using (ExcelRange r = sheet.Cells[rowIndex, TotalAmtCol - 1, rowIndex, TotalAmtCol])
                {
                    r.Style.Font.SetFromFont("bold".ToFont());
                    r.Value = ConfigurationManager.AppSettings["TotalAmountLabel"] ?? "TOTAL COST";
                    r.Merge = true;
                    r.AddBorder();
                }

                using (ExcelRange r = sheet.Cells[rowIndex, TotalAmtCol + 1, rowIndex, TotalAmtCol + 2])
                {
                    r.Style.Numberformat.Format = "#,##0.00";
                    r.Formula = "0" + TotalAmtfn.ToString();
                    r.Merge = true;
                    r.AddBorder();
                }
            }

            WriteToLog("CreateGridRows2 end:" + DateTime.Now.TimeOfDay);
            WriteXLSMCellFormatValues(formatcellssheet);

            if (PRDheaders.Length == 0)
            {
                sheet.Cells.AutoFitColumns(8);
                sheet.SetColumnWidth(3, column4width);
                // Unit Cost
                sheet.SetColumnWidth(4, column7width);
                sheet.SetColumnWidth(5, column8width);
                sheet.SetColumnWidth(6, column9width);
            }
            else
            {
                sheet.Cells.AutoFitColumns(lastDeptColumn + 4);
                // Unit Cost
                sheet.SetColumnWidth(lastDeptColumn + 1, column7width);
                sheet.SetColumnWidth(lastDeptColumn + 2, column8width);
                sheet.SetColumnWidth(lastDeptColumn + 3, column9width);
            }

            #region "Order Template"
            globalsheet.Hidden = eWorkSheetHidden.Hidden;

            SetCellValue(globalsheet, 18, 2, "OrderShipDate");
            SetCellValue(globalsheet, 18, 1, OrderShipDate ? "1" : "0");

            SetCellValue(globalsheet, 97, 0, "Create Date");
            SetCellValue(globalsheet, 97, 1, String.Format("[{0:s}]", DateTime.Now));

            xReader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'SafetyStockQty'";
            if (xReader.GetSchemaTable().DefaultView.Count > 0)
            {
                var SafetyStockQty = xReader["SafetyStockQty"] == null ? string.Empty : xReader["SafetyStockQty"].ToString();
                SetCellValue(globalsheet, 98, 1, SafetyStockQty);
            }

            SetCellValue(globalsheet, 99, 0, "catalog");
            SetCellValue(globalsheet, 99, 1, xReader["catalog"].ToString());

            SetCellValue(globalsheet, 100, 0, "catalogname");
            SetCellValue(globalsheet, 100, 1, xReader["catalogname"].ToString());

            SetCellValue(globalsheet, 101, 0, "CancelDefaultDays");
            xReader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'CancelDefaultDays'";
            if (xReader.GetSchemaTable().DefaultView.Count > 0 && xReader["CancelDefaultDays"] != null)
            {
                CancelAfterDays = int.Parse(xReader["CancelDefaultDays"].ToString());
                SetCellValue(globalsheet, 101, 1, xReader["CancelDefaultDays"].ToString());
            }

            SetCellValue(globalsheet, 102, 0, "RestrictedReqDateRange");
            xReader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'RestrictedReqDateRange'";
            if (xReader.GetSchemaTable().DefaultView.Count > 0 && xReader["RestrictedReqDateRange"] != null)
            {
                SetCellValue(globalsheet, 102, 1, xReader["RestrictedReqDateRange"].ToString());
            }

            SetCellValue(globalsheet, 199, 1, ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            SetCellValue(globalsheet, 199, 2, "SheetPassword");

            SetCellValue(globalsheet, 200, 2, "DisplaySoldtoShipList");
            if (OLOFSettings.Settings.ExistsKey("DisplaySoldtoShipList"))
            {
                SetCellValue(globalsheet, 200, 1, "1");
            }
            #endregion

            globalsheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            skusheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            pricesheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            formatcellssheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            atpsheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            soldtoshiptosheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");

            var StandardProtectStructure = string.IsNullOrEmpty(ConfigurationManager.AppSettings["StandardProtectStructure"]) ?
                "1" : ConfigurationManager.AppSettings["StandardProtectStructure"];
            if(StandardProtectStructure == "1")
                XLSMPck.Workbook.Protection.LockStructure = true;
            #endregion

            return maxcols;
        }

        private void AddCategorySubtotal(ExcelWorksheet sheet, StringBuilder SubTotalQtyfn, StringBuilder SubTotalAmtfn, int rowIndex)
        {
            sheet.Row(rowIndex).Height = RowHeight;
            using (ExcelRange r = sheet.Cells[rowIndex, TotalAmtCol - 2, rowIndex, TotalAmtCol - 2])
            {
                r.Style.Font.SetFromFont("bold".ToFont());
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                r.Value = "SUBTOTAL";
                r.AddBorder();
            }
            using (ExcelRange r = sheet.Cells[rowIndex, TotalAmtCol - 1, rowIndex, TotalAmtCol - 1])
            {
                r.Style.Numberformat.Format = "#,##0";
                r.Formula = "0" + SubTotalQtyfn.ToString();
                r.AddBorder();
            }
            using (ExcelRange r = sheet.Cells[rowIndex, TotalAmtCol, rowIndex, TotalAmtCol])
            {
                r.Style.Numberformat.Format = "#,##0.00";
                r.Formula = "0" + SubTotalAmtfn.ToString();
                r.AddBorder();
            }
        }

        private static void SetColumnWidth(ExcelWorksheet sheet, int columnwidth, ExcelWorksheet skusheet, ExcelWorksheet pricesheet, int colctr)
        {
            sheet.SetColumnWidth(colctr, columnwidth);
            if (skusheet != null)
                skusheet.SetColumnWidth(colctr, columnwidth);
            if (pricesheet != null)
                pricesheet.SetColumnWidth(colctr, columnwidth);
            if (columnwidth == 0)
                sheet.Column(colctr).Hidden = true;
        }

        private void SetCellValue(ExcelWorksheet sheet, int row, int col, object value)
        {
            sheet.Cells[row + 1, col + 1].Value = value;
        }

        private void GenerateDateSheet(ExcelPackage XLSMPck, SqlDataReader reader, bool withImage)
        {
            List<DateTime> DateSheetCollection = new List<DateTime>();
            bool TopDeptGroupStyle = (ConfigurationManager.AppSettings["TopDeptGroupStyle"] ?? "0") == "1";
            bool AddThresholdQtyComments = false;

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

                try
                {
                    reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'AddThresholdQtyComments'";
                    if (reader.GetSchemaTable().DefaultView.Count > 0 && reader["AddThresholdQtyComments"] != null)
                    {
                        AddThresholdQtyComments = (reader["AddThresholdQtyComments"].ToString() == "1");
                        WriteToLog("Add comments is enabled");
                    }
                }
                catch
                { }
            }

            ExcelWorksheet skuSheet = XLSMPck.Workbook.Worksheets["CellFormats"];
            ExcelWorksheet atpSheet = XLSMPck.Workbook.Worksheets["ATPDate"];
            ExcelWorksheet tmpSheet = XLSMPck.Workbook.Worksheets["Order Template"];
            ExcelWorksheet globalsheet = XLSMPck.Workbook.Worksheets["GlobalVariables"];
            string lastSheetName = "Order Template";

            bool RestrictedReqDateRange = false; // if true then DateShipEnd is required
            try
            {
                reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'RestrictedReqDateRange'";
                if (reader.GetSchemaTable().DefaultView.Count > 0 && reader["RestrictedReqDateRange"] != null)
                {
                    RestrictedReqDateRange = (reader["RestrictedReqDateRange"].ToString() == "1");
                }
            }
            catch
            { }

            DataTable imageTB = null;
            if (withImage)
            {
                var skus = GetFirstSKUList(XLSMPck);
                //generate images in offline
                imageTB = GetSKUImageList(skus);
                if (imageTB != null)
                {
                    if (ConfigurationManager.AppSettings["column1Imagewidth"] != null)
                        tmpSheet.Column(1).Width = Convert.ToInt32(ConfigurationManager.AppSettings["column1Imagewidth"].ToString()) / 5;

                    SaveImagesToOffline(imageTB, tmpSheet);
                }
            }
            else
            {
                AddLogoToDateSheet(tmpSheet);
            }

            ExcelWorksheet dftSheet = null;
            var DateSheets = new List<ExcelWorksheet>();
            int datediff = 0;
            for (int i = 0; i < DateSheetCollection.Count; i++)
            {
                var dtSheet = XLSMPck.Workbook.Worksheets.Copy("Order Template", string.Format("{0:ddMMMyyyy}", DateSheetCollection[i]));
                DateSheets.Add(dtSheet);
                if (i == 0) dftSheet = dtSheet;

                if (reader.GetSchemaTable().Columns.Contains("DateShipEnd"))
                {
                    SetCellValue(globalsheet, 0, i, string.Format("{0:ddMMMyyyy}", DateSheetCollection[i]));

                    SetCellValue(globalsheet, 1, i, DateSheetCollection[i]);

                    if (DateSheetCollection.Count > i + 1)
                    {
                        datediff = (int)(DateSheetCollection[i + 1] - DateSheetCollection[i]).TotalDays - 1;
                    }
                    else
                    {
                        datediff = (int)((DateTime)reader["DateShipEnd"] - DateSheetCollection[i]).TotalDays;
                    }
                    SetCellValue(globalsheet, 2, i, DateSheetCollection[i].AddDays(datediff));
                }

                dtSheet.Cells["A:XFD"].Style.Font.Name = ConfigurationManager.AppSettings["FontName"] ?? "Arial";

                string address = OLOFSettings.Settings.GetValueByKey("DefaultRequestDelDate", "C8");
                dtSheet.Cells[address].Value = DateSheetCollection[i];

                if (CancelAfterDays.HasValue)
                {
                    address = OLOFSettings.Settings.GetValueByKey("DefaultCancelDate", "C9");
                    dtSheet.Cells[address].Value = DateSheetCollection[i].AddDays(CancelAfterDays.Value);
                }

                XLSMPck.Workbook.Worksheets.MoveAfter(dtSheet.Name, lastSheetName);
                lastSheetName = dtSheet.Name;

                WriteToLog("Date Sheet " + dtSheet.Name + " ATP Check, Add Comments etc. begin:" + DateTime.Now.TimeOfDay);
                PresetForProductSheet(dtSheet, skuSheet, atpSheet, DateSheetCollection[i], colors, AddThresholdQtyComments);
                WriteToLog("Date Sheet " + dtSheet.Name + " ATP Check, Add Comments etc. end:" + DateTime.Now.TimeOfDay);

                //if (imageTB != null)
                //{
                //    WriteToLog("Date Sheet " + dtSheet.Name + " Save Image to Sheet begin:" + DateTime.Now.TimeOfDay);
                //    SaveImagesToOffline(imageTB, dtSheet);
                //    WriteToLog("Date Sheet " + dtSheet.Name + " Save Image to Sheet end:" + DateTime.Now.TimeOfDay);
                //}

                if (System.IO.File.Exists(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\DateSheet.vba")))
                    dtSheet.CodeModule.Code = File.ReadAllText(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\DateSheet.vba"));

                dtSheet.Protection.IsProtected = true;
                dtSheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            }
            //if (imageTB != null)
            //{
            //    WriteToLog("Date Sheets Save Image to Sheet begin:" + DateTime.Now.TimeOfDay);
            //    SaveImagesToOfflineDateSheets(imageTB, DateSheets);
            //    WriteToLog("Date Sheets Save Image to Sheet end:" + DateTime.Now.TimeOfDay);
            //}
            //foreach (var dateSheet in DateSheets)
            //{
            //    if (System.IO.File.Exists(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\DateSheet.vba")))
            //        dateSheet.CodeModule.Code = File.ReadAllText(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\DateSheet.vba"));

            //    dateSheet.Protection.IsProtected = true;
            //    dateSheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            //}

            dftSheet.View.TabSelected = true;
            tmpSheet.Hidden = eWorkSheetHidden.Hidden;
        }

        private void AddLogoToDateSheet(ExcelWorksheet tmpSheet)
        {
            var logoPath = ConfigurationManager.AppSettings["LogoPath"];
            Image img = Image.FromFile(logoPath);
            if (img != null)
            {
                ExcelPicture pic = tmpSheet.Drawings.AddPicture("logo", img);
                pic.From.Column = 0;
                pic.From.Row = 0;
                pic.To.Column = 0;
                pic.To.Row = 1;
                pic.From.ColumnOff = Pixel2MTU(10);
                pic.From.RowOff = Pixel2MTU(10);
                pic.SetSize(100);
            }
        }

        private List<SKUImageInfo> GetFirstSKUList(ExcelPackage XLSMPck)
        {
            var column4heading = System.Configuration.ConfigurationManager.AppSettings["column4heading"];
            List<SKUImageInfo> skus = new List<SKUImageInfo>();
            var sheet = XLSMPck.Workbook.Worksheets["WebSKU"];

            var tsheet = XLSMPck.Workbook.Worksheets["Order Template"];

            var lastRow = sheet.LastRowNum();
            var maxColNumber = 20;
            var firstColumn = 8; //H
            var firstRow = int.Parse(System.Configuration.ConfigurationManager.AppSettings["SheetPrdFirstRow"] ?? "6");
            for (int i = firstRow; i <= lastRow; i++)
            {
                var skuvalue = string.Empty;
                for (int j = firstColumn; j < maxColNumber + firstColumn; j++)
                {
                    skuvalue = sheet.Cells[i, j].Text;
                    if (!string.IsNullOrEmpty(skuvalue))
                    {
                        break;
                    }
                }
                if (string.IsNullOrEmpty(skuvalue))
                    if (tsheet.Cells[i, 1].Value == null && tsheet.Cells[i, 2].Value != null )
                    {
                        skuvalue = tsheet.Cells[i, 2].Value.ToString();
                    }
                if (!string.IsNullOrEmpty(skuvalue) && skuvalue != column4heading)
                {
                    var sku = new SKUImageInfo();
                    sku.SKU = skuvalue;
                    sku.RowIndex = i;
                    skus.Add(sku);
                }
            }
            return skus;
        }

        private void SaveImagesToOffline(DataTable dt, ExcelWorksheet tmpSheet)
        {
            var ImageSizeReset = string.IsNullOrEmpty(ConfigurationManager.AppSettings["ImageSizeReset"]) ? "1" : ConfigurationManager.AppSettings["ImageSizeReset"];
            var ImageResetPercent = string.IsNullOrEmpty(ConfigurationManager.AppSettings["ImageResetPercent"]) ? 80 : int.Parse(ConfigurationManager.AppSettings["ImageResetPercent"]);
            var imageFolder = ConfigurationManager.AppSettings["ImageFolder"];
            string[] ignoreStyles = (ConfigurationManager.AppSettings["IgnoreStyles"] ?? "").Split(new char[] { ',' });
            foreach (DataRow row in dt.Rows)
            {
                var rowIndex = (int)row["ROW"];
                row["ImageName"] = row["ImageName"].ToString().Replace(".jpg", ".png");
                var imageName = row["ImageName"].ToString();
                var imagePath = Path.Combine(imageFolder, imageName);
                if (ignoreStyles.Contains(imageName.Replace(".png", "")))
                {
                    tmpSheet.Row(rowIndex).Height = 30;
                }
                else if (File.Exists(imagePath))
                {
                    Image prdimg = Image.FromFile(imagePath);
                    tmpSheet.Row(rowIndex).Height = prdimg.Height * 0.9;

                    /*WriteToLog(imagePath);
                    if (ImageFormat.Jpeg.Equals(prdimg.RawFormat))
                        WriteToLog("JPEG");
                    if (ImageFormat.Png.Equals(prdimg.RawFormat))
                        WriteToLog("PNG");
                    if (ImageFormat.Gif.Equals(prdimg.RawFormat))
                        WriteToLog("GIF");
                    if (ImageFormat.Bmp.Equals(prdimg.RawFormat))
                        WriteToLog("BMP");
                    WriteToLog("Height: " + prdimg.Height.ToString());
                    WriteToLog("Width: " + prdimg.Width.ToString());*/
                }
            }

            AddLogoToDateSheet(tmpSheet);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                var sku = row["SKU"].ToString();
                var rowIndex = (int)row["ROW"];
                var imageName = row["ImageName"].ToString();
                var imagePath = Path.Combine(imageFolder, imageName);
                if (ignoreStyles.Contains(imageName.Replace(".png", "")))
                    continue;
                if (File.Exists(imagePath))
                {
                    ExcelPicture pic = tmpSheet.Drawings.AddPicture(System.Guid.NewGuid().ToString(), new FileInfo(imagePath));
                    pic.From.Column = 0;
                    pic.From.Row = rowIndex - 1;
                    pic.From.ColumnOff = Pixel2MTU(10);
                    pic.From.RowOff = Pixel2MTU(5);

                    if (ImageSizeReset == "1")
                        pic.SetSize(ImageResetPercent);
                    else
                    {
                        try
                        {
                            var imgHeight = 0;
                            var imgWidth = 0;
                            using (Image prdimg = Image.FromFile(imagePath))
                            {
                                imgHeight = prdimg.Height;
                                imgWidth = prdimg.Width;
                            }
                            if (imgWidth > 0)
                                pic.SetSize((int)(imgWidth * ImageResetPercent / 100), (int)(imgHeight * ImageResetPercent / 100));
                            else
                                pic.SetSize(ImageResetPercent);
                        }
                        catch
                        {
                            pic.SetSize(ImageResetPercent);
                        }
                    }
                }
            }
        }

        private void SaveImagesToOfflineDateSheets(DataTable dt, List<ExcelWorksheet> DateSheets)
        {
            var ImageSizeReset = string.IsNullOrEmpty(ConfigurationManager.AppSettings["ImageSizeReset"]) ? "1" : ConfigurationManager.AppSettings["ImageSizeReset"];
            var ImageResetPercent = string.IsNullOrEmpty(ConfigurationManager.AppSettings["ImageResetPercent"]) ? 80 : int.Parse(ConfigurationManager.AppSettings["ImageResetPercent"]);
            var imageFolder = ConfigurationManager.AppSettings["ImageFolder"];
            string[] ignoreStyles = (ConfigurationManager.AppSettings["IgnoreStyles"] ?? "").Split(new char[] { ',' });

            foreach (DataRow row in dt.Rows)
            {
                var rowIndex = (int)row["ROW"];
                row["ImageName"] = row["ImageName"].ToString().Replace(".jpg", ".png");
                var imageName = row["ImageName"].ToString();
                var imagePath = Path.Combine(imageFolder, imageName);

                if (ignoreStyles.Contains(imageName.Replace(".png", "")))
                {
                    foreach (var tmpSheet in DateSheets)
                        tmpSheet.Row(rowIndex).Height = 30;
                }
                else if (File.Exists(imagePath))
                {
                    using (Image prdimg = Image.FromFile(imagePath))
                    {
                        foreach (var tmpSheet in DateSheets)
                            tmpSheet.Row(rowIndex).Height = prdimg.Height * 0.9;
                    }
                }
            }
            return;

            foreach (var tmpSheet in DateSheets)
                AddLogoToDateSheet(tmpSheet);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                var rowIndex = (int)row["ROW"];
                row["ImageName"] = row["ImageName"].ToString().Replace(".jpg", ".png");
                var imageName = row["ImageName"].ToString();
                var imagePath = Path.Combine(imageFolder, imageName);
                var sku = row["SKU"].ToString();

                if (ignoreStyles.Contains(imageName.Replace(".png", "")))
                    continue;

                if (File.Exists(imagePath))
                {
                    using (Image prdimg = Image.FromFile(imagePath))
                    {
                        foreach (var tmpSheet in DateSheets)
                        {
                            //ExcelPicture pic = tmpSheet.Drawings.AddPicture(System.Guid.NewGuid().ToString(), new FileInfo(imagePath));
                            ExcelPicture pic = tmpSheet.Drawings.AddPicture(System.Guid.NewGuid().ToString(), prdimg);
                            pic.From.Column = 0;
                            pic.From.Row = rowIndex - 1;
                            pic.From.ColumnOff = Pixel2MTU(10);
                            pic.From.RowOff = Pixel2MTU(5);

                            if (ImageSizeReset == "1")
                                pic.SetSize(ImageResetPercent);
                            else
                            {
                                try
                                {
                                    var imgHeight = prdimg.Height;
                                    var imgWidth = prdimg.Width;
                                    if (imgWidth > 0)
                                        pic.SetSize((int)(imgWidth * ImageResetPercent / 100), (int)(imgHeight * ImageResetPercent / 100));
                                    else
                                        pic.SetSize(ImageResetPercent);
                                }
                                catch (Exception ex)
                                {
                                    WriteToLog(ex.Message);
                                    WriteToLog(ex.StackTrace);
                                    pic.SetSize(ImageResetPercent);
                                }
                            }
                        }
                    }
                }
            }
        }

        public int Pixel2MTU(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
        }

        private void PresetForProductSheet(ExcelWorksheet prdSheet, ExcelWorksheet skuSheet, ExcelWorksheet atpSheet,
            DateTime? sheetDate, Dictionary<int, Color> colors, bool AddThresholdQtyComments)
        {
            int? ThresholdQty = null;

            if (AddThresholdQtyComments)
            {
                int qty;
                if (int.TryParse(ConfigurationManager.AppSettings["ThresholdQty"]??"99", out qty))
                {
                    ThresholdQty = qty;
                }
            }
            for (int i = 2; i <= skuSheet.LastRowNum(); i++)
            {
                string cellPos = skuSheet.Cells[i, 1].Text;
                string multipleValue = skuSheet.Cells[i, 4].Text;

                ExcelRange cell = prdSheet.Cells[cellPos];
                if (sheetDate.HasValue)
                {
                    try
                    {
                        string dateValues = atpSheet.Cells[cellPos].Text;
                        if (SupportATP)
                        {
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
                            if (availbleQty > 0)
                            {
                                //var list1 = cell.DataValidation.AddCustomDataValidation();
                                //list1.Formula.ExcelFormula = string.Format("(MOD(indirect(address(row(),column())) ,{0})=0)", multipleValue);
                                //list1.ShowErrorMessage = true;
                                //list1.Error = string.Format("You must enter a multiple of {0} in this cell.", multipleValue).Replace(".00 ", " ").Replace(".0 ", " ");

                                if (ThresholdQty.HasValue)
                                {
                                    string cmt = string.Empty;
                                    if (availbleQty > ThresholdQty)
                                    {
                                        cmt = string.Format("{0}+", ThresholdQty.Value);
                                    }
                                    else
                                    {
                                        cmt = string.Format("{0}", availbleQty);
                                    }
                                    var comment = cell.AddComment(cmt, "Plumriver");
                                    comment.Locked = true;
                                    comment.AutoFit = true;
                                }
                                cell.Style.Locked = false;
                                cell.Style.Numberformat.Format = "##0";
                            }
                            else
                            {
                                availbleQty = 0;
                            }
                            var list = colors.Where(kvp => kvp.Key <= availbleQty);
                            if (list.Count() > 0)
                            {
                                int colorIdx = list.LastOrDefault().Key;
                                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(colors[colorIdx]);
                            }
                        }
                        else
                        {
                            cell.Style.Locked = false;
                            cell.Style.Numberformat.Format = "##0";
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteToLog(ex.Message);
                    }

                }
                else
                {
                    cell.Style.Locked = false;
                }
            }
        }

        private void WriteXLSMCellFormatValues(ExcelWorksheet formatcellssheet)
        {
            int ctr = 2;
            var rowNum = formatcellssheet.Dimension.End.Row + 1;
            //Write previously stored order multiple restrictions in the cellformat sheet
            if (multiples.Count > unlocked.Count)
            {
                foreach (DictionaryEntry de in multiples)
                {
                    //HSSFRow row = (HSSFRow)formatcellssheet.CreateRow(rowNum);
                    //row.CreateCell(0);
                    //row.CreateCell(1);
                    //row.CreateCell(2).SetCellValue(Convert.ToString(de.Key));
                    //row.CreateCell(3).SetCellValue(Convert.ToString(de.Value));
                    formatcellssheet.Cells[3.IntToMoreChar() + rowNum.ToString()].Value = Convert.ToString(de.Key);
                    formatcellssheet.Cells[4.IntToMoreChar() + rowNum.ToString()].Value = Convert.ToString(de.Value);
                    rowNum++;
                }
                //Write previously stored unlocked cell inventory in the cellformat sheet
                foreach (DictionaryEntry de in unlocked)
                {
                    //formatcellssheet.GetRow(ctr).GetCell(0).SetCellValue(Convert.ToString(de.Key));
                    formatcellssheet.Cells[1.IntToMoreChar()+ (ctr + 1).ToString()].Value = Convert.ToString(de.Key);
                    ctr++;
                }
            }
            else
            {
                foreach (DictionaryEntry de in unlocked)
                {
                    //HSSFRow row = (HSSFRow)formatcellssheet.CreateRow(rowNum);
                    //row.CreateCell(0).SetCellValue(Convert.ToString(de.Key));
                    //row.CreateCell(1);
                    //row.CreateCell(2);
                    //row.CreateCell(3);
                    formatcellssheet.Cells[1.IntToMoreChar() + rowNum.ToString()].Value = Convert.ToString(de.Key);
                    rowNum++;
                }
                foreach (DictionaryEntry de in multiples)
                {
                    //formatcellssheet.GetRow(ctr).GetCell(2).SetCellValue(Convert.ToString(de.Key));
                    //formatcellssheet.GetRow(ctr).GetCell(3).SetCellValue(Convert.ToString(de.Value));
                    formatcellssheet.Cells[3.IntToMoreChar() + ctr.ToString()].Value = Convert.ToString(de.Key);
                    formatcellssheet.Cells[4.IntToMoreChar() + ctr.ToString()].Value = Convert.ToString(de.Value);
                    ctr++;
                }
            }
        }

        private void AddCatalogUPCSheetHeader(ExcelWorksheet sheet)
        {
            if (sheet == null)
                return;
            var header = OLOFSettings.Settings.GetValueByKey("CatalogUPCSheetHeader", "Style Name|Material #|Gender|Size|UPC Code|Purchase Price");
            var headers = header.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < headers.Count(); i++)
            {
                sheet.SetValue(1, i + 1, headers[i]);
            }
        }

        private int AddBreakRows(ExcelWorksheet sheet, int rownumber)
        {
           //Todo
            return 0;
        }

        private void AddFormatCellsHeader(ExcelWorksheet sheet)
        {
            sheet.Cells["A1"].Value = ("Unlocked Cells");
            sheet.Cells["C1"].Value = ("Multiple Cells");
            sheet.Cells["D1"].Value = ("Multiple Value");
        }

        private void AddTextRow(ExcelWorksheet sheet, int sheetRow, float rowHeight, string[][] texts)
        {
            sheet.Row(sheetRow).Height = (double)rowHeight;
            for (int i = 0; i < texts.Length; i++)
            {
                string[] arr = texts[i];
                int colStartNumber = int.Parse(arr[2]) + 1;
                int colEndNumber = int.Parse(arr[3]) + 1;
                sheet.SetValue(sheetRow, colStartNumber, arr[0]);

                using (ExcelRange r = sheet.Cells[string.Format("{0}{2}:{1}{2}", colStartNumber.IntToMoreChar(), colEndNumber.IntToMoreChar(), sheetRow)])
                {
                    if (colStartNumber != colEndNumber) 
                        r.Merge = true;

                    r.Style.Font.SetFromFont(arr[1].ToFont());
                }
            }
        }

        private void CreateGrid2(XElement xroot, OrderedDictionary depts, int deptid, ExcelWorksheet sheet, int sheetRow, 
            ref ListDictionary colpositions, string soldto, string catalog, bool? groupedTitle, string PrepackList)
        {
            //Create a new grid
            OrderedDictionary deptcols = null;
            bool bRectangular = ConfigurationManager.AppSettings["AttributeRectangular"] == null ? true : (ConfigurationManager.AppSettings["AttributeRectangular"] == "1" ? true : false);
            if (!bRectangular) deptcols = GetDeptCols(deptid, soldto, catalog);
            else deptcols = GetDeptColsFromXML(xroot, deptid);
            AddBlankHeader(ref deptcols, deptid);

            if (deptcols != null && deptcols.Count > 0)
            {
                if (groupedTitle.HasValue)
                {
                    if (groupedTitle.Value)
                    {
                        colpositions = AddHeader2(depts, deptcols, sheet, ref sheetRow, true, PrepackList);
                        sheetRow++;
                    }
                    
                    colpositions = AddHeader2(depts, deptcols, sheet,ref sheetRow, false, PrepackList);
                }
                else
                {
                    colpositions = AddHeader2(depts, deptcols, sheet,ref sheetRow, null, PrepackList);
                }
            }
        }

        private ListDictionary AddHeader2(OrderedDictionary depts, OrderedDictionary cols,
            ExcelWorksheet sheet,ref int sheetRow, bool? groupedTitle, string PrepackList)
        {
            int[] borderRange = new int[4];
            borderRange[0] = sheetRow; borderRange[1]=1;

            bool topTitle = false; int PrepackAmt = 0;
            string[] sizePrepack = PrepackList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (this.TopDeptGroupStyle && groupedTitle.HasValue && groupedTitle.Value)
            {
                topTitle = true;
                if (sizePrepack.Length > 0)
                {
                    string[] arr = sizePrepack[0].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    PrepackAmt = arr.Length;
                }
            }

            ListDictionary colpositions = new ListDictionary();
            int colcounter = 0;
            string column1heading = System.Configuration.ConfigurationManager.AppSettings["column1heading"].ToString();
            string column2heading = System.Configuration.ConfigurationManager.AppSettings["column2heading"].ToString();
            string column3heading = System.Configuration.ConfigurationManager.AppSettings["column3heading"].ToString();
            string column4heading = System.Configuration.ConfigurationManager.AppSettings["column4heading"].ToString();
            string column5heading = System.Configuration.ConfigurationManager.AppSettings["column5heading"].ToString();
            string column6heading = System.Configuration.ConfigurationManager.AppSettings["column6heading"].ToString();
            string column7heading = System.Configuration.ConfigurationManager.AppSettings["column7heading"].ToString();
            string column8heading = System.Configuration.ConfigurationManager.AppSettings["column8heading"].ToString();
            string column9heading = System.Configuration.ConfigurationManager.AppSettings["column9heading"].ToString();
            if (groupedTitle.HasValue && !groupedTitle.Value)
            {
                column1heading = column2heading = column3heading = column4heading = column5heading = column6heading = column7heading = column8heading = column9heading = string.Empty;
            }
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            bool firstrow = true;
            var headers = OLOFSettings.Settings.GetValueByKey("ProductDetailHeaders", "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

            var cellIndex = 0;
            cellIndex++;
            if (!multicolumndeptlabel)   //Put the whole dept hierarchy in the first header cell
            {
                string deptlabel = "> ";
                foreach (DictionaryEntry de in depts)
                    deptlabel += de.Value + " > ";
                deptlabel = deptlabel.Substring(0, deptlabel.Length - 3);
                ExcelRange er = sheet.Cells[cellIndex.IntToMoreChar() + sheetRow.ToString()];
                if (!groupedTitle.HasValue || groupedTitle == false)
                {
                    er.Value = deptlabel;
                    er.Style.Font.SetFromFont("bold".ToFont());
                }
            }
            else   //Lay the dept hierarchy out w/each level in its own column
            {
                sheet.SetValue(sheetRow, cellIndex, column1heading);
                foreach (DictionaryEntry de in depts)
                {
                    if (!firstrow)
                    {
                        cellIndex++;
                    }
                    firstrow = false;
                }
                //Fill the rest of the dept column labels horizontally with blanks
                if (depts.Count < maxcols)
                {
                    for (int i = 0; i < maxcols - depts.Count; i++)
                    {
                        cellIndex++;
                    }
                }
            }
            if (this.TopDeptGroupStyle && !topTitle)
            {
                using (ExcelRange r = sheet.Cells[string.Format("{0}{2}:{1}{2}", 1.IntToMoreChar(), (headers.Length == 0 ? 4 : (2 + headers.Length)).IntToMoreChar(), sheetRow)])
                {
                    r.Merge = true;
                }
            }
            if (topTitle)
            {
                using (ExcelRange r = sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow+1)])
                {
                    r.Merge = true;
                }
            }
            cellIndex++;
            // Material #
            var cell = sheet.Cells[string.Format("{0}{1}", cellIndex.IntToMoreChar(), sheetRow)];
            cell.Value = column4heading;
            if (topTitle)
            {
                sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow + 1)].Merge = true;
            }
            cellIndex++;

            if (headers.Length > 0)
            {
                for (int i = 0; i < headers.Count(); i++)
                {
                    cell = sheet.Cells[string.Format("{0}{1}", cellIndex.IntToMoreChar(), sheetRow)];
                    cell.Value = headers[i];
                    if (topTitle)
                    {
                        sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow + 1)].Merge = true;
                    }
                    cellIndex++;
                }
            }
            else
            {
                cell = sheet.Cells[string.Format("{0}{1}", cellIndex.IntToMoreChar(), sheetRow)];
                cell.Value = (column5heading);
                if (topTitle)
                {
                    sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow + 1)].Merge = true;
                }
                cellIndex++;
                cell = sheet.Cells[string.Format("{0}{1}", cellIndex.IntToMoreChar(), sheetRow)];
                cell.Value = (column6heading);
                if (topTitle)
                {
                    sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow + 1)].Merge = true;
                }
                cellIndex++;
            }

            // Unit cost
            cell = sheet.Cells[string.Format("{0}{1}", cellIndex.IntToMoreChar(), sheetRow)];
            cell.Value = (column7heading);
            if (topTitle)
            {
                sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow + 1)].Merge = true;
            }
            cellIndex++;
            cell = sheet.Cells[string.Format("{0}{1}", cellIndex.IntToMoreChar(), sheetRow)];
            cell.Value = (column8heading);
            if (topTitle)
            {
                sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow + 1)].Merge = true;
            }
            if (TotalSection)
            {
                TotalQtyCol = cellIndex;
            }
            cellIndex++;
            cell = sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow + 1)];
            cell.Value = (column9heading);
            if (topTitle)
            {
                sheet.Cells[string.Format("{0}{1}:{0}{2}", cellIndex.IntToMoreChar(), sheetRow, sheetRow + 1)].Merge = true;
            }
            if (TotalSection)
            {
                TotalAmtCol = cellIndex;
            }
            cellIndex++;
            if (topTitle)
            {
                int j = 0;
                foreach (DictionaryEntry de in cols)
                {
                    cell = sheet.Cells[string.Format("{0}{1}:{0}{2}", (cellIndex + (j++)).IntToMoreChar(), sheetRow, sheetRow + 1)];
                    if (j == 1)
                    {
                        cell.Value = (ConfigurationManager.AppSettings["Attr1Label"] ?? "Size");
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }
                sheet.Cells[string.Format("{0}{2}:{1}{2}", cellIndex.IntToMoreChar(), (cellIndex + cols.Count - 1).IntToMoreChar(), sheetRow )].Merge = true;
            }

            if (!groupedTitle.HasValue || groupedTitle == true)
                sheetRow++;

            int preparkIdx = 0; int[] preparkRows = new int[PrepackAmt];
            colcounter = cellIndex ;
            for (int i = 0; i < PrepackAmt; i++)
            {
                preparkRows[i] = sheetRow + i ;
            }
            int maxColIdx = 0;
            string displaySubTitle = ConfigurationManager.AppSettings["DisplaySubTitle"] ?? "0";
            foreach (DictionaryEntry de in cols)
            {
                if (!groupedTitle.HasValue || groupedTitle == true || displaySubTitle == "1")
                {
                    cell = sheet.Cells[(colcounter).IntToMoreChar() + sheetRow.ToString()];
                    cell.Value = (string.IsNullOrEmpty(Convert.ToString(de.Value)) ? " " : StripHTML(Convert.ToString(de.Value)).Replace("\r", System.Environment.NewLine));
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.WrapText = true;
                    if (displaySubTitle == "1" && groupedTitle == false)
                        cell.Style.Font.Color.SetColor(Color.Gray);
                    if (sizePrepack.Length > preparkIdx)
                    {
                        string[] preparkItem = sizePrepack[preparkIdx++].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < preparkRows.Length; i++)
                        {
                            int preparkQty = 0;
                            int.TryParse(preparkItem[i], out preparkQty);
                            sheet.SetValue(preparkRows[i] + 1, colcounter, preparkQty);
                        }
                    }
                }
                colpositions.Add(StripHTML(Convert.ToString(de.Value)), colcounter);
                colcounter++;
                if (colcounter > maxColIdx) maxColIdx = colcounter;
            }
            if (topTitle)
            {
                for (int i = 0; i < PrepackAmt; i++)
                {

                    sheetRow++;
                    sheet.Row(sheetRow).Height = (double)RowHeight;

                    for (int j = 0; j < cellIndex; j++)
                    {
                        if (j == 0)
                        {
                            ExcelRange er = sheet.Cells[1.IntToMoreChar() + sheetRow.ToString()];
                            er.Value = string.Format("Prepack {0}  ", ((char)(65 + i)).ToString());
                            er.Style.Font.SetFromFont("bold".ToFont());
                            er.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        }
                        else if (j == cellIndex - 2)
                        {
                            ExcelRange er = sheet.Cells[(j ).IntToMoreChar() + sheetRow.ToString()];
                            er.Formula = string.Format("SUM({0}{2}:{1}{2})", (cellIndex-1).ToName(), (cellIndex + sizePrepack.Length-2).ToName(), sheetRow);
                        }
                    }
                    sheet.Cells[string.Format("{0}{2}:{1}{2}", 1.IntToMoreChar(), (cellIndex - 3).IntToMoreChar(), sheetRow)].Merge = true;
                }
            }
            borderRange[2] = sheetRow; borderRange[3] = maxColIdx;
            for (int row = borderRange[0]; row <= borderRange[2]; row++)
            {
                for (int col = borderRange[1]; col < borderRange[3]; col++)
                    sheet.Cells[col.IntToMoreChar()+row.ToString()].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }
            return colpositions;
        }

        private int CreateGridRows2(OrderedDictionary depts, ExcelWorksheet sheet, ExcelWorksheet skusheet, ExcelWorksheet pricesheet,
            ExcelWorksheet atpsheet, ExcelWorksheet upcsheet, ExcelWorksheet OrderMultiplesheet, 
            int deptid, string deptname, ref ListDictionary colpositions)
        {
            bool firstrow = true, lastrow = false;
            int i = 0, maxColIdx = 0;
            OrderedDictionary rows = null;
            if (departments.Contains(deptid))
                rows = (OrderedDictionary)departments[deptid];
            else
                WriteToLog("CreateGridRows2 cannot find department id from department list: " + deptid.ToString());

            if (rows != null)
            {
                foreach (DictionaryEntry de in rows)
                {
                    i++;
                    if (i == rows.Count)
                    {
                        lastrow = true;
                    }
                    StringDictionary rowattribs = (StringDictionary)de.Key;
                    OrderedDictionary skus = (OrderedDictionary)de.Value;
                    try
                    {
                        int col = AddRow2(depts, rowattribs, skus, colpositions, sheet, skusheet, pricesheet, atpsheet, upcsheet, OrderMultiplesheet, firstrow, lastrow);

                        maxColIdx = col > maxColIdx ? col : maxColIdx;
                        firstrow = false;
                    }
                    catch (Exception ex)
                    {
                        WriteToLog(deptid.ToString());
                        WriteToLog(rowattribs["Style"].ToString());
                        throw ex;
                    }
                }
            }
            return maxColIdx;
        }

        private int AddRow2(OrderedDictionary DeptLevels, StringDictionary rowattribs, OrderedDictionary skus, ListDictionary colpositions,
            ExcelWorksheet sheet, ExcelWorksheet skusheet, ExcelWorksheet pricesheet, ExcelWorksheet atpsheet, ExcelWorksheet upcsheet, ExcelWorksheet OrderMultiplesheet,
            bool firstrow, bool lastrow)
        {
            string LineFontStyle = ConfigurationManager.AppSettings["LineFontStyle"] ?? (ConfigurationManager.AppSettings["FontSize"] ?? "8");

            string errMsg = "";
            var rowIndex = sheet.Dimension.End.Row + (firstrow ? 0 : 1);
            var upcRowIndex = 0;
            if (upcsheet != null) upcRowIndex = upcsheet.LastRowNum() + 1;
            //Add the row to each sheet to keep them in synch
            int offset = 6;
            if (DeptLevels != null && DeptLevels.Count > 0 && multicolumndeptlabel)
                offset = DeptLevels.Count + 5;
            ExcelRange cell = null;
            string valueformula = "";
            string ttlformula = "";
            sheet.Row(rowIndex).Height = RowHeight;
            //Add spacer cells to the secondary sheets
            string wholesaleprice = "";
            try
            {
                if (this.IncludePricing)
                {
                    //Strip a star from the end of wholesale price
                    if (rowattribs["RowPriceWholesale"].ToString().Length == 6 && rowattribs["RowPriceWholesale"].ToString().Substring(5, 1) == "*")
                        wholesaleprice = rowattribs["RowPriceWholesale"].ToString().Substring(0, 5);
                    else
                        wholesaleprice = rowattribs["RowPriceWholesale"].ToString();
                }
                else
                {
                    wholesaleprice = ExcludePriceValue; // "0";
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
            }
            //If it's not a new top level, merge vertically - same idea for subsequent levels
            int ictr = 0;
            int cellIndex = 1;
            if (!multicolumndeptlabel)
            {
                using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                {

                }
                cellIndex++;
            }
            else
            {
                foreach (DictionaryEntry de in DeptLevels)
                {
                    if (firstrow)
                    {
                        using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                        {
                            r.Value = de.Value.ToString();
                        }

                        cellIndex++;
                    }
                    else
                    {
                        using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                        {
                            //r.Value = de.Value.ToString();
                        }
                        cellIndex++;
                    }
                    ictr++;
                }
                //Fill the rest of the dept column labels horizontally with blanks
                if (ictr < maxcols)
                {
                    for (int i = 0; i < maxcols - ictr; i++)
                    {
                        if (firstrow)
                        {
                            using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                            {
                                //r.Value = de.Value.ToString();
                            }
                            cellIndex++;
                        }
                        else
                        {
                            using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                            {
                                //r.Value = de.Value.ToString();
                            }
                            cellIndex++;
                        }
                    }
                }
            }
            //Mat.
            using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
            {
                r.Style.Font.SetFromFont(LineFontStyle.ToFont());
                r.Value = rowattribs["Style"].ToString();
            }
            cellIndex++;

            var firstitem = skus.Cast<DictionaryEntry>().ElementAt(0);
            var attr2 = ((StringDictionary)(firstitem.Value)).ContainsKey("AttributeValue2") ? ((StringDictionary)(firstitem.Value))["AttributeValue2"].ToString() : string.Empty;
            var attr3 = ((StringDictionary)(firstitem.Value)).ContainsKey("AttributeValue3") ? ((StringDictionary)(firstitem.Value))["AttributeValue3"].ToString() : string.Empty;
            var attr4 = ((StringDictionary)(firstitem.Value)).ContainsKey("AttributeValue4") ? ((StringDictionary)(firstitem.Value))["AttributeValue4"].ToString() : string.Empty;
            var styleName = rowattribs["ProductName"] == null ? rowattribs["Style"].ToString() : rowattribs["ProductName"].ToString();
            var style = rowattribs["Style"].ToString();
            var upc = ((StringDictionary)(firstitem.Value))["UPC"].ToString();
            var unitprice = ((StringDictionary)(firstitem.Value))["PriceWholesale"].ToString();


            var lines = OLOFSettings.Settings.GetValueByKey("ProductDetailLines", "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length > 0)
            {
                for (int i = 0; i < lines.Count(); i++)
                {
                    var cellvalue = string.Empty;
                    switch (lines[i])
                    {

                        case "attr2":
                            cellvalue = attr2;
                            break;
                        case "attr3":
                            cellvalue = attr3;
                            break;
                        case "attr4":
                            cellvalue = attr4;
                            break;
                        case "styleName":
                            cellvalue = styleName;
                            break;
                    }

                    using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                    {
                        r.Style.Font.SetFromFont(LineFontStyle.ToFont());
                        r.Value = cellvalue;
                    }
                    cellIndex++;
                }
            }
            else
            {
                //Mat Desc
                using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                {
                    var descString = string.Empty;
                    r.Style.Font.SetFromFont(LineFontStyle.ToFont());
                    if (rowattribs["ProductName"] == null)
                    {
                        r.Value = (rowattribs["Style"].ToString());
                    }
                    else
                    {
                        r.Value = (rowattribs["ProductName"].ToString());
                    }
                }
                cellIndex++;

                //Dim1
                using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                {
                    r.Style.Font.SetFromFont(LineFontStyle.ToFont());
                    r.Value = rowattribs["GridAttributeValues"].ToString();
                }
                cellIndex++;
            }

            //WHS $
            var price = this.IncludePricing ? double.Parse(rowattribs["RowPriceWholesale"].ToString()) : double.Parse(ExcludePriceValue); // 0;
            SetCellValue(sheet, rowIndex, cellIndex, price, true, true).Style.Font.SetFromFont(LineFontStyle.ToFont()); ;
            cellIndex++;

            //TTL
            var TTLIndex = cellIndex;
            ExcelRange ttlcell = sheet.Cells[rowIndex, cellIndex];
            ttlcell.Style.Font.SetFromFont(LineFontStyle.ToFont());
            cellIndex++;

            //TTL Value
            ExcelRange valuecell = sheet.Cells[rowIndex, cellIndex];
            valuecell.Style.Font.SetFromFont(LineFontStyle.ToFont());
            valuecell.Style.Numberformat.Format = "#,##0.00";
            cellIndex++;

            //Lay out the row with empty, greyed-out cells
            for (int i = 0; i < colpositions.Count; i++)
            {
                //cell = (HSSFCell)newrow.CreateCell(cellIndex);
                //cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s81");
                //if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
                cellIndex++;
            }
            //AddSecondaryCells(newrowa, newrowb, newrowc, offset + colpositions.Count);

            //Now enable valid cells
            try
            {
                foreach (DictionaryEntry de in skus)
                {
                    //SKU must be enabled
                    if (((StringDictionary)(de.Value))["Enabled"].ToString() == "1")
                    {
                        //This is the column header value
                        string attr1 = ((StringDictionary)(de.Value))["AttributeValue1"].ToString();
                        string cellposition = "";
                        //If the row contains a cell with the column header value, enable it
                        attr1 = StripHTML(Convert.ToString(attr1));
                        if (colpositions.Contains(attr1))
                        {
                            int colposition = Convert.ToInt32(colpositions[attr1].ToString());
                            if (multicolumndeptlabel)
                                cell = sheet.Cells[rowIndex, Convert.ToInt32(colpositions[attr1].ToString()) + 1];
                            else
                                cell = sheet.Cells[rowIndex, Convert.ToInt32(colpositions[attr1].ToString())];
                            cellposition = cell.Address;
                            cell.Style.Font.SetFromFont(LineFontStyle.ToFont());

                            unlocked.Add(cellposition, cellposition);
                            if (((StringDictionary)(de.Value)).ContainsKey("OrderMultiple"))
                            {
                                string multipleValue = ((StringDictionary)(de.Value))["OrderMultiple"].ToString();
                                multiples.Add(cellposition, multipleValue);
                            }
                            cell = SetCellValue(OrderMultiplesheet, cellposition, ((StringDictionary)(de.Value))["OrderMultiple"].ToString(), true);
                            //Plant the QuickWebSKU on its sheet
                            cell = SetCellValue(skusheet, cellposition, ((StringDictionary)(de.Value))["QuickWebSKUValue"].ToString(), false);
                            //Plant the ATPDate on its sheet
                            if (OrderShipDate)
                                cell = SetCellValue(atpsheet, cellposition, ((StringDictionary)(de.Value))["ATPDate"].ToString(), false);
                            else
                                cell = SaveATPValueDataToSheet(atpsheet, cellposition, ((StringDictionary)(de.Value))["QuickWebSKUValue"].ToString());
                            //Plant the wholesale price on its sheet
                            cell = SetCellValue(pricesheet, cellposition, ((StringDictionary)(de.Value))["PriceWholesale"].ToString(), true);
                            //Build the ttl cell formula 
                            ttlformula += cellposition + "+";
                            //Build the value cell formula
                            valueformula += "(" + GetExcelColumnName(colposition) + (rowIndex).ToString() +
                                "*WholesalePrice!" + GetExcelColumnName(colposition) + (rowIndex).ToString() + ")" + "+";

                            if (upcsheet != null)
                            {
                                var lineItems = OLOFSettings.Settings.GetValueByKey("CatalogUPCSheetLine", "prdName|style|attr3|attr1|upc|unitprice").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                if(ConfigurationManager.AppSettings["ShowUPCFromSKUValue"] == "1")
                                    upc = ((StringDictionary)(de.Value))["UPC"].ToString();
                                for (int i = 0; i < lineItems.Length; i++)
                                {
                                    //prdName|style|attr3|attr4|attr2|attr1|upc|unitprice
                                    switch (lineItems[i])
                                    {
                                        case "styleName":
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, styleName, false);
                                            break;
                                        case "style":
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, style, false);
                                            break;
                                        case "attr3":
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, attr3, false);
                                            break;
                                        case "attr2":
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, attr2, false);
                                            break;
                                        case "attr1":
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, attr1, false);
                                            break;
                                        case "attr4":
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, attr4, false);
                                            break;
                                        case "upc":
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, upc, false);
                                            break;
                                        case "unitprice":
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, unitprice, false);
                                            break;
                                    }
                                }
                                upcRowIndex++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            { WriteToLog("foreach (DictionaryEntry de in skus)" + rowattribs["PfID"]); throw; }
            //Assign the formulas
            if (ttlformula.Length > 0)
                ttlcell.Formula = ttlformula.Substring(0, ttlformula.Length - 1);
            if (valueformula.Length > 0)
                valuecell.Formula = valueformula.Substring(0, valueformula.Length - 1);

            return cellIndex;
        }

        private ExcelRange SetCellValue(ExcelWorksheet sheet, int rowIndex, int cellIndex, object value, bool bdouble)
        { 
            return SetCellValue(sheet, rowIndex, cellIndex, value, bdouble, false);
        }

        private ExcelRange SetCellValue(ExcelWorksheet sheet, int rowIndex, int cellIndex, object value, bool bdouble, bool withCurrency)
        {
            ExcelRange cell = sheet.Cells[rowIndex, cellIndex];
            if (bdouble)
            {
                cell.Style.Numberformat.Format = withCurrency ? string.Format(@"#,##0.00 ""{0}""", this.Currency) : "#,##0.00";
                cell.Value = double.Parse(value.ToString());
            }
            else
                cell.Value = (value);
            return cell;
        }

        private ExcelRange SetCellValue(ExcelWorksheet sheet, string address, object value, bool bdouble)
        { 
            return SetCellValue(sheet, address, value, bdouble, false);
        }
        
        private ExcelRange SetCellValue(ExcelWorksheet sheet, string address, object value, bool bdouble, bool withCurrency)
        {
            ExcelRange cell = sheet.Cells[address];
            if (bdouble)
            {
                cell.Style.Numberformat.Format = withCurrency ? string.Format("#,##0.00 {0}", this.Currency) : "#,##0.00";
                cell.Value = double.Parse(value.ToString());
            }
            else
                cell.Value = (value);
            return cell;
        }

        private ExcelRange SaveATPValueDataToSheet(ExcelWorksheet sheet, string address, string SKU)
        {
            ExcelRange cell = null;
            if (ATPValueData == null)
                return null;
            var cellValue = string.Empty;
            var atpQty = 0;
            foreach (DataRow drow in ATPValueData.Rows)
            {
                if (drow["SKU"].ToString() == SKU)
                {
                    var dQty = (int)drow["ATPQty"];
                    atpQty += dQty;
                    cellValue += drow["ATPDate"].ToString() + "|" + atpQty.ToString() + ",";
                }
            }
            if (!string.IsNullOrEmpty(cellValue))
            {
                cellValue = cellValue.Substring(0, cellValue.Length - 1);
                cell = SetCellValue(sheet, address, cellValue, false);
            }
            return cell;
        }

        private string CleanFileName(string fileName)
        {
            return Path.GetInvalidFileNameChars().Aggregate(fileName, (current, c) => current.Replace(c.ToString(), ""));
        }
    }

    public static class ExtEPPlus
    {
        public static bool ExistsKey(this OLOFSettingAdd[] list, string keyname)
        {
            return (list.Any(e => e.key == keyname));
        }

        public static string GetValueByKey(this OLOFSettingAdd[] list, string keyname)
        {
            return list.GetValueByKey(keyname, string.Empty);
        }

        public static string GetValueByKey(this OLOFSettingAdd[] list, string keyname, string defaultValue)
        {
            var item = list.Where(e => e.key == keyname).FirstOrDefault<OLOFSettingAdd>();
            if (item != null)
                return item.value;
            else
                return defaultValue;
        }

        public static void AddBorder(this ExcelRange rg)
        {
            rg.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
        }

        public static void SetColumnWidth(this ExcelWorksheet sheet, int column, int width)
        {
            sheet.Column(column + 1).Width = (double)width;
        }

        public static int LastRowNum(this ExcelWorksheet sheet)
        {
            return sheet.Dimension==null?0:sheet.Dimension.End.Row;
        }

        public static void SetRowHeight(this ExcelWorksheet sheet, int row, int height)
        {
            sheet.Row(row + 1).Height = (double)height;
        }

        public static Font ToFont(this string value)
        {
            int fontsize = 0; 
            FontStyle fontstyle = FontStyle.Regular;

            if (!int.TryParse(Regex.Replace(value, "[a-zA-Z]+", ""), out fontsize))
            {
                fontsize = short.Parse(ConfigurationManager.AppSettings["FontSize"] ?? "8");
            }
            if (value.ToLower().Contains("bold"))
                fontstyle = FontStyle.Bold;
            else
                fontstyle = FontStyle.Regular;

            return new Font(ConfigurationManager.AppSettings["FontName"] ?? "Arial", fontsize, fontstyle);
        }

        public static Font ToFont(this string value, string splitchar)
        {
            string[] arr = value.ToLower().Split(splitchar.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            int fontsize = 0;
            FontStyle fontstyle = FontStyle.Regular;
            fontsize = short.Parse(ConfigurationManager.AppSettings["FontSize"] ?? "8");
            fontstyle = FontStyle.Regular;

            foreach (string str in arr)
            {
                if (str.EndsWith("pt"))
                    int.TryParse(Regex.Replace(str, "[a-zA-Z]+", ""), out fontsize);

                if (str.ToLower().Contains("bold"))
                    fontstyle = FontStyle.Bold;
            }
            return new Font(ConfigurationManager.AppSettings["FontName"] ?? "Arial", fontsize, fontstyle);
        }

        public static string IntToMoreChar(this int value)
        {
            string rtn = string.Empty;
            List<int> iList = new List<int>();

            //To single Int
            while (value / 26 != 0 || value % 26 != 0)
            {
                iList.Add(value % 26);
                value /= 26;
            }

            //Change 0 To 26
            for (int j = 0; j < iList.Count - 1; j++)
            {
                if (iList[j] == 0)
                {
                    iList[j + 1] -= 1;
                    iList[j] = 26;
                }
            }

            //Remove 0 at last
            if (iList[iList.Count - 1] == 0)
            {
                iList.Remove(iList[iList.Count - 1]);
            }

            //To String
            for (int j = iList.Count - 1; j >= 0; j--)
            {
                char c = (char)(iList[j] + 64);
                rtn += c.ToString();
            }

            return rtn;
        }
    }
}
