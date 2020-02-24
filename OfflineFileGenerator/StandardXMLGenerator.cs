using System;
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
        public bool OfflineGenerateSuccessful { get; set; }

        int? CancelAfterDays = null;
        public string Currency = string.Empty;
        public bool SupportATP = true;
        OLOFSetting OLOFSettings = null;

        public string PriceType { get; set; }
        public string OfflineDeptID { get; set; }

        public bool CatalogSupportSalesProgram { get; set; }
        public List<SalesProgram> CatalogSalesPrograms { get; set; }
        public string CatalogSalesProgramAddress { get; set; }
        public string BookingOrder { get; set; }
        public List<string> UPCSheets { get; set; }

        public int SKUTotalUnitColumn { get; set; }
        public string AllowBackOrder { get; set; }

        string[] ATPLevelBackColors = (ConfigurationManager.AppSettings["ATPLevelBackColors"] ?? "0|Gray").Split(new char[] { ',', '|' }, StringSplitOptions.RemoveEmptyEntries);
        Dictionary<int, Color> colors = new Dictionary<int, Color>();

        public int GenerateStandardXLSMData(string filepath, string savedirectory, SqlDataReader reader)
        {
            OfflineGenerateSuccessful = true;

            //LoadSalesProgramList("VANS_US_F16_FOOTWEAR_DC20", "10052735");
            string soldto = reader["SoldTo"].ToString();
            string catalog = reader["Catalog"].ToString();
            string pricecode = reader["PriceCode"].ToString();
            this.CatalogName = reader["CatalogName"].ToString();
            reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'ATPFlag'");
            if (reader.GetSchemaTable().DefaultView.Count > 0)
            {
                SupportATP = reader["ATPFlag"] == null ? true : (reader["ATPFlag"].ToString() == "0" ? false : true);
            }
            reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'BookingCatalog'");
            if (reader.GetSchemaTable().DefaultView.Count > 0)
                BookingOrder = reader["BookingCatalog"].ToString();
            reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'AllowBackOrder'");
            if (reader.GetSchemaTable().DefaultView.Count > 0)
                AllowBackOrder = reader["AllowBackOrder"].ToString();
            WriteToLog("AllowBackOrder: " + AllowBackOrder);

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
                    var HideConditionSheetIndex = ConfigurationManager.AppSettings["HideConditionSheetIndex"]; //GENERAL TERMS & CONDITIONS
                    for (int i = 1; i <= cntSheet; i++)
                    {
                        if (reg.IsMatch(pck.Workbook.Worksheets[1].Name) || HideConditionSheetIndex == i.ToString())
                            pck.Workbook.Worksheets[1].Hidden = eWorkSheetHidden.Hidden;
                        pck.Workbook.Worksheets.MoveToEnd(1);
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
                                GenerateDateSheet(targepck, reader, false, catalog, soldto);
                                targepck.Save();
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteToLog("Fail to Generate DateSheet:" + ex.Message);
                            OfflineGenerateSuccessful = false;

                            if (File.Exists(targetName))
                            {
                                File.Delete(targetName);
                            }
                        }
                    }

                    //for TOMS, it's configured as 2.
                    WriteToLog("Generate Order Date Sheets begin:" + DateTime.Now.TimeOfDay);
                    if (FormStyles > 1)
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
                                    GenerateDateSheet(targeImgpck, reader, true, catalog, soldto);
                                    targeImgpck.Save();
                                }
                            }
                            catch (Exception ex)
                            {
                                WriteToLog("Error:" + ex.Message);
                                WriteToLog(ex.StackTrace);
                                if (ex.InnerException != null)
                                {
                                    try
                                    {
                                        WriteToLog(ex.InnerException.Message);
                                        WriteToLog(ex.InnerException.StackTrace);
                                    }
                                    catch { }
                                }
                                OfflineGenerateSuccessful = false;

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
                OfflineGenerateSuccessful = false;
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
            var DisplayPriceAmount = OLOFSettings.Settings.GetValueByKey("DisplayPriceAmount", "1");

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

                        try
                        {
                            if (!string.IsNullOrEmpty(item.color))
                            {
                                var color = System.Drawing.ColorTranslator.FromHtml(item.color);
                                cell.Style.Font.Color.SetColor(color);
                            }
                        }
                        catch { }

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

                        if (item.columnName == "CancelDate")
                        {
                            var DisplayCancelDate = OLOFSettings.Settings.GetValueByKey("DisplayCancelDate", "1");
                            if (DisplayCancelDate == "0")
                                sheet.Row(line.RowNumber).Hidden = true;
                        }
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

                    soldtoshiptosheet.Row(i + 2).Hidden = true; //Crocs EU - OLOF, they can unprotect the sheet #226 
                }
                soldtoshiptosheet.Row(1).Hidden = true; //Crocs EU - OLOF, they can unprotect the sheet #226 
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
            var TotalFormulaWithDetailCell = string.IsNullOrEmpty(ConfigurationManager.AppSettings["TotalFormulaWithDetailCell"]) ?
                "1" : ConfigurationManager.AppSettings["TotalFormulaWithDetailCell"];

            var PRDheaders = OLOFSettings.Settings.GetValueByKey("ProductDetailHeaders", "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            var lastDeptColumn = (PRDheaders.Length == 0 ? 3 : (PRDheaders.Length + 1));
            UPCSheets = new List<string>();
            
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.EndElement && reader.Name == "ProductLevel")
                {
                    Dictionary<string, string[]> rootDept = new Dictionary<string, string[]>();
                    //Build a grid
                    XmlReader currentNode = reader.ReadSubtree();
                    int deptid = 0;
                    string deptname = string.Empty, PrepackList = string.Empty;
                    OrderedDictionary depts = GetDeptHierarchy(currentNode, ref deptid, ref deptname, ref rootDept, catalog);

                    var sheetRow = sheet.Dimension.End.Row + 1;
                    if (rootDept.Count > 0)
                    {
                        if (rootDept != null && lastRootDeptID != rootDept.FirstOrDefault().Key)
                        {
                            if (SubTotalQtyfn.Length > 0)
                            {
                                var rowIndex = sheet.LastRowNum() + 1;
                                AddCategorySubtotal(sheet, SubTotalQtyfn, SubTotalAmtfn, rowIndex);
                                if (TotalFormulaWithDetailCell != "1")
                                {
                                    TotalQtyfn.AppendFormat("+SUM({0}{1}:{0}{2})", (TotalAmtCol - 1).IntToMoreChar(), rowIndex, rowIndex);
                                    TotalAmtfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalAmtCol.IntToMoreChar(), rowIndex, rowIndex);
                                }

                                sheetRow++;
                                rowIndex++;
                                SubTotalQtyfn = new StringBuilder();
                                SubTotalAmtfn = new StringBuilder();
                            }

                            sheetRow++;

                            var StandardUPCSheetExtension = ConfigurationManager.AppSettings["StandardUPCSheetExtension"];
                            var deptUPCSheetName = rootDept.FirstOrDefault().Value[0] + (string.IsNullOrEmpty(StandardUPCSheetExtension) ? string.Empty : " " + StandardUPCSheetExtension);
                            deptUPCSheetName = deptUPCSheetName.Length > 22 ?
                                (rootDept.FirstOrDefault().Value[0].Length > 22 ? rootDept.FirstOrDefault().Value[0].Substring(0, 22) : rootDept.FirstOrDefault().Value[0])
                                : deptUPCSheetName;
                            upcsheet = XLSMPck.Workbook.Worksheets.Add(deptUPCSheetName);
                            var StandardUPCSheetTabColor = ConfigurationManager.AppSettings["StandardUPCSheetTabColor"];//A9A9A9 FOR EC
                            if(!string.IsNullOrEmpty(StandardUPCSheetTabColor))
                                upcsheet.TabColor = System.Drawing.ColorTranslator.FromHtml(StandardUPCSheetTabColor);
                            AddCatalogUPCSheetHeader(upcsheet);
                            UPCSheets.Add(deptUPCSheetName);

                            string[] arr = rootDept.FirstOrDefault().Value;
                            AddBreakRows(sheet, 1);
                            AddTextRow(sheet, sheetRow, 20, new string[][] { new string[] { arr[0], "18pt", "0", "2" } });
                            sheetRow++;
                            var blackedOutMergeCells = DisplayPriceAmount == "1" ? 3 : 6;
                            AddTextRow(sheet, sheetRow, 15, new string[][] { 
                                new string[] { (string.Format("{0}: {1}", arr[1], arr[2]) == ": ") ? string.Empty : string.Format("{0}: {1}", arr[1], arr[2]),"12pt", "0", lastDeptColumn.ToString(), "0" },
                                new string[] { "BLACKED OUT = NOT AVAILABLE","Bold", (lastDeptColumn+1).ToString(), (lastDeptColumn+blackedOutMergeCells).ToString(), "8" } 
                            });
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
                    if (TotalFormulaWithDetailCell == "1")
                    {
                        TotalQtyfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalQtyCol.IntToMoreChar(), gId, sheet.Dimension.End.Row);
                        TotalAmtfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalAmtCol.IntToMoreChar(), gId, sheet.Dimension.End.Row);
                    }
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

                if (TotalFormulaWithDetailCell != "1")
                {
                    TotalQtyfn.AppendFormat("+SUM({0}{1}:{0}{2})", (TotalAmtCol - 1).IntToMoreChar(), rowIndex, rowIndex);
                    TotalAmtfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalAmtCol.IntToMoreChar(), rowIndex, rowIndex);
                }
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

                if (DisplayPriceAmount == "1")
                {
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
                var ProductDetailWidthList = OLOFSettings.Settings.GetValueByKey("ProductDetailWidth", "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                if (ProductDetailWidthList.Length > 0)
                {
                    var columnBegin = lastDeptColumn + 1 - ProductDetailWidthList.Length;
                    for (int pw = 0; pw < ProductDetailWidthList.Length; pw++)
                    {
                        var columnWidth = int.Parse(ProductDetailWidthList[pw]);
                        sheet.SetColumnWidth(columnBegin + pw, columnWidth);
                    }
                }
                else
                    sheet.Cells.AutoFitColumns(lastDeptColumn + 6);
                // Unit Cost
                sheet.SetColumnWidth(lastDeptColumn + 1, column7width);
                sheet.SetColumnWidth(lastDeptColumn + 2, column8width);
                sheet.SetColumnWidth(lastDeptColumn + 3, column9width);
            }

            if (DisplayPriceAmount == "0")
            {
                //unit price column
                sheet.Column(lastDeptColumn + 1 + 1).Hidden = true;
                //total amount column
                sheet.Column(lastDeptColumn + 3 + 1).Hidden = true;
            }

            //SET WIDTH OF ORDER DATA COLUMNS
            var StandardResetOrderDataColumnWidth = ConfigurationManager.AppSettings["StandardResetOrderDataColumnWidth"];
            if (StandardResetOrderDataColumnWidth == "1")
            {
                var totalAmountColumn = lastDeptColumn + 3 + 1;
                int datacolumnwidth = Convert.ToInt32(ConfigurationManager.AppSettings["datacolumnwidth"] ?? "12");
                int dataColumnNumber = ConfigurationManager.AppSettings.Get("datacolumnnumber") == null ? 30 : Convert.ToInt32(ConfigurationManager.AppSettings["datacolumnnumber"].ToString());
                for (int i = 0; i < dataColumnNumber; i++)
                {
                    sheet.SetColumnWidth(totalAmountColumn + i, datacolumnwidth);
                }
            }

            #region "Order Template" SET GlobalVariables
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

            //set for summary tab
            SetCellValue(globalsheet, 103, 1, "");
            globalsheet.Names.Add("SummarizeAll", globalsheet.Cells[104, 2]);
            SetCellValue(globalsheet, 104, 1, "N");
            globalsheet.Names.Add("bSelectedHere", globalsheet.Cells[105, 2]);
            var totalColumn = 'A';
            if (SKUTotalUnitColumn > 0)
                totalColumn = (char)(totalColumn + SKUTotalUnitColumn - 1);
            SetCellValue(globalsheet, 105, 1, totalColumn.ToString());
            globalsheet.Names.Add("TotalsColumn", globalsheet.Cells[106, 2]);

            SetCellValue(globalsheet, 199, 1, ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
            SetCellValue(globalsheet, 199, 2, "SheetPassword");
            globalsheet.Row(200).Hidden = true; //hide the pwd //Crocs EU - OLOF, they can unprotect the sheet #226

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
            for (int i = 0; i < UPCSheets.Count; i++)
            {
                var upcSheet = XLSMPck.Workbook.Worksheets[UPCSheets[i]];
                if (upcSheet != null )
                {
                    upcSheet.Protection.IsProtected = true;
                    upcSheet.Protection.AllowAutoFilter = true;
                    upcSheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
                }
            }

            //Crocs EU - OLOF, they can unprotect the sheet #226
            var CategoryProtectVBAProject = ConfigurationManager.AppSettings["CategoryProtectVBAProject"] ?? "0";
            if (CategoryProtectVBAProject == "1")
                XLSMPck.Workbook.VbaProject.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");

            var StandardProtectStructure = string.IsNullOrEmpty(ConfigurationManager.AppSettings["StandardProtectStructure"]) ?
                "1" : ConfigurationManager.AppSettings["StandardProtectStructure"];
            if(StandardProtectStructure == "1")
                XLSMPck.Workbook.Protection.LockStructure = true;
            #endregion

            #region and totals on header for EC & Vans; //EC QAS/PRD: Add Default Zoom to OLOF and Display Total Units near top of OLOF #232
            var StandardShowTotalOnTop = ConfigurationManager.AppSettings["StandardShowTotalOnTop"] ?? "0";
            if (StandardShowTotalOnTop == "1")
            {
                var StandardProductHeaderBGColor = string.Empty;//ConfigurationManager.AppSettings["StandardProductHeaderBGColor"] ?? (GenerateStandardXLSMDataWithTemplateTabs == "2" ? "#CC0000" : string.Empty);
                var headerColor = !string.IsNullOrEmpty(StandardProductHeaderBGColor) ? System.Drawing.ColorTranslator.FromHtml(StandardProductHeaderBGColor) : Color.Transparent;
                //product totals
                var TotalPairsColumn = string.Empty;
                if (TotalAmtCol <= 0)
                    TotalAmtCol = int.Parse(ConfigurationManager.AppSettings["StandardTotalAmountColumn"] ?? "8");
                if (TotalAmtCol > 0)
                    TotalPairsColumn = (TotalAmtCol - 1).IntToMoreChar();
                //else
                //    TotalPairsColumn = ConfigurationManager.AppSettings["StandardTotalPairsColumn"] ?? "G";
                var headerTotalTitleColumn = (TotalAmtCol - 2).IntToMoreChar();
                var headerCell = SetCellValue(sheet, headerTotalTitleColumn + "11", "TOTALS:");
                if (!string.IsNullOrEmpty(StandardProductHeaderBGColor))
                {
                    headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerCell.Style.Fill.BackgroundColor.SetColor(headerColor);
                }
                else
                    SetCellStyle(headerCell);
                headerCell = GetCell(sheet, TotalPairsColumn + "11");
                headerCell.Formula = "VLOOKUP(\"total pairs\",$" + TotalPairsColumn + "$15:$J$15000,3,0)";
                if (!string.IsNullOrEmpty(StandardProductHeaderBGColor))
                {
                    headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerCell.Style.Fill.BackgroundColor.SetColor(headerColor);
                }
                else
                    SetCellStyle(headerCell);
                var headerTotalAmtColumn = TotalAmtCol.IntToMoreChar();
                headerCell = GetCell(sheet, headerTotalAmtColumn + "11");
                headerCell.Formula = "VLOOKUP(\"total cost\",$" + TotalPairsColumn + "$15:$J$15000,3,0)";
                if (!string.IsNullOrEmpty(StandardProductHeaderBGColor))
                {
                    headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerCell.Style.Fill.BackgroundColor.SetColor(headerColor);
                }
                else
                    SetCellStyle(headerCell);
            }
            #endregion

            return maxcols;
        }

        private void AddCategorySubtotal(ExcelWorksheet sheet, StringBuilder SubTotalQtyfn, StringBuilder SubTotalAmtfn, int rowIndex)
        {
            var SubTotalQtyfnStr = SubTotalQtyfn.ToString();
            var SubTotalAmtfnStr = SubTotalAmtfn.ToString();
            var TotalFormulaWithDetailCell = string.IsNullOrEmpty(ConfigurationManager.AppSettings["TotalFormulaWithDetailCell"]) ? 
                "1" : ConfigurationManager.AppSettings["TotalFormulaWithDetailCell"];
            if (TotalFormulaWithDetailCell == "0")
            {
                if (SubTotalQtyfnStr.IndexOf(":") > 0)
                {
                    var fQIndex = SubTotalQtyfnStr.IndexOf(":");
                    var lQIndex = SubTotalQtyfnStr.LastIndexOf(":");
                    SubTotalQtyfnStr = SubTotalQtyfnStr.Substring(0, fQIndex) + SubTotalQtyfnStr.Substring(lQIndex);
                }

                if (SubTotalAmtfnStr.IndexOf(":") > 0)
                {
                    var fAIndex = SubTotalAmtfnStr.IndexOf(":");
                    var lAIndex = SubTotalAmtfnStr.LastIndexOf(":");
                    SubTotalAmtfnStr = SubTotalAmtfnStr.Substring(0, fAIndex) + SubTotalAmtfnStr.Substring(lAIndex);
                }
            }

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
                r.Formula = "0" + SubTotalQtyfnStr;
                r.AddBorder();
            }
            using (ExcelRange r = sheet.Cells[rowIndex, TotalAmtCol, rowIndex, TotalAmtCol])
            {
                r.Style.Numberformat.Format = "#,##0.00";
                r.Formula = "0" + SubTotalAmtfnStr;
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

        private void GenerateDateSheet(ExcelPackage XLSMPck, SqlDataReader reader, bool withImage, string catalog, string soldto)
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

            try
            {
                ExcelWorksheet skuSheet = XLSMPck.Workbook.Worksheets["CellFormats"];
                ExcelWorksheet atpSheet = XLSMPck.Workbook.Worksheets["ATPDate"];
                ExcelWorksheet tmpSheet = XLSMPck.Workbook.Worksheets["Order Template"];
                ExcelWorksheet globalsheet = XLSMPck.Workbook.Worksheets["GlobalVariables"];
                string lastSheetName = "Order Template";

                #region CatalogSalesProgram
                //WriteToLog("sales program begin");
                CatalogSalesPrograms = new List<SalesProgram>();
                CatalogSalesProgramAddress = string.Empty;
                if (string.IsNullOrEmpty(BookingOrder) || BookingOrder == "0")
                    CheckLoadSalesProgram(catalog, soldto, DateSheetCollection);
                /* Will load sales program according to the tab ship-date
                if (CatalogSalesPrograms.Count > 0)
                {
                    WriteToLog(CatalogSalesPrograms.Count.ToString());
                    var SalesProgramIndex = 1;
                    foreach (SalesProgram row in CatalogSalesPrograms)
                    {
                        CatalogSalesProgramAddress = "P" + SalesProgramIndex.ToString();
                        globalsheet.Cells[CatalogSalesProgramAddress].Value = row.SalesProgramCode;
                        SalesProgramIndex++;
                    }
                    CatalogSalesProgramAddress = "P1:" + CatalogSalesProgramAddress;
                    WriteToLog(CatalogSalesProgramAddress);
                }*/
                //WriteToLog("sales program end");
                #endregion

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

                        //moved image-add from date sheet to template so that spend less time
                        //IMAGE MIXED WHEN COPY FROM TEMP TO DATE SHEET, HAVE TO DISABLE IT AND SET StandardAddImageInTemp = 0
                        if (ConfigurationManager.AppSettings["StandardAddImageInTemp"] == "1")
                        {
                            WriteToLog("Date Sheets Save Image to Temp begin:" + DateTime.Now.TimeOfDay);
                            SaveImagesToOffline(imageTB, tmpSheet);
                            WriteToLog("Date Sheets Save Image to Temp begin:" + DateTime.Now.TimeOfDay);
                        }
                    }
                }
                else
                {
                    AddLogoToDateSheet(tmpSheet);
                }
                WriteToLog("template image");

                ExcelWorksheet dftSheet = null;
                int datediff = 0;
                var SalesProgramIndex = 1;

                if (DateSheetCollection == null || DateSheetCollection.Count <= 0)
                    WriteToLog("No valid DateSheetCollection," + DateTime.Now.TimeOfDay);
                else
                {
                    for (int i = 0; i < DateSheetCollection.Count; i++)
                    {
                        try
                        {
                            var dtSheet = XLSMPck.Workbook.Worksheets.Copy("Order Template", string.Format("{0:ddMMMyyyy}", DateSheetCollection[i]));
                            var StandardDateSheetTabColor = ConfigurationManager.AppSettings["StandardDateSheetTabColor"];//FFFF00 FOR EC
                            if (!string.IsNullOrEmpty(StandardDateSheetTabColor))
                                dtSheet.TabColor = System.Drawing.ColorTranslator.FromHtml(StandardDateSheetTabColor);
                            if (i == 0) dftSheet = dtSheet;

                            if (reader.GetSchemaTable().Columns.Contains("DateShipEnd"))
                            {
                                WriteToLog("Date Sheet " + dtSheet.Name + " DateShipEnd," + DateTime.Now.TimeOfDay);
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

                            WriteToLog("Date Sheet " + dtSheet.Name + " DefaultRequestDelDate," + DateTime.Now.TimeOfDay);
                            var DefaultDateFormat = OLOFSettings.Settings.GetValueByKey("DefaultDateFormat", "");
                            string address = OLOFSettings.Settings.GetValueByKey("DefaultRequestDelDate", "C8");
                            if (string.IsNullOrEmpty(DefaultDateFormat))
                                dtSheet.Cells[address].Value = DateSheetCollection[i];
                            else
                                dtSheet.Cells[address].Value = DateSheetCollection[i].ToString(DefaultDateFormat);

                            if (CancelAfterDays.HasValue && CancelAfterDays.Value > 0)
                            {
                                WriteToLog("Date Sheet " + dtSheet.Name + " DefaultCancelDate," + DateTime.Now.TimeOfDay);
                                address = OLOFSettings.Settings.GetValueByKey("DefaultCancelDate", "C9");
                                if (string.IsNullOrEmpty(DefaultDateFormat))
                                    dtSheet.Cells[address].Value = DateSheetCollection[i].AddDays(CancelAfterDays.Value);
                                else
                                    dtSheet.Cells[address].Value = DateSheetCollection[i].AddDays(CancelAfterDays.Value).ToString(DefaultDateFormat);
                            }

                            //SET SALES PROGRAM IN EACH DATE SHEET
                            string snaddress = OLOFSettings.Settings.GetValueByKey("DefaultSalesProgramName", "");
                            string svaddress = OLOFSettings.Settings.GetValueByKey("DefaultSalesProgramValue", "");
                            if (CatalogSupportSalesProgram)
                            {
                                WriteToLog("sales program begin:" + DateSheetCollection[i].ToShortDateString());
                                if (!string.IsNullOrEmpty(svaddress) && CatalogSalesPrograms.Count > 0)
                                {
                                    var validCount = 0;
                                    var salesPorgramAddress = "P" + SalesProgramIndex.ToString();
                                    foreach (SalesProgram row in CatalogSalesPrograms)
                                    {
                                        if (row.ShipStartDate <= DateSheetCollection[i] && row.ShipEndDate >= DateSheetCollection[i])
                                        {
                                            globalsheet.Cells["P" + SalesProgramIndex.ToString()].Value = row.SalesProgramCode;
                                            validCount++;
                                            SalesProgramIndex++;
                                        }
                                    }
                                    if (validCount > 0)
                                    {
                                        salesPorgramAddress = "=INDIRECT(\"GlobalVariables!" + salesPorgramAddress + ":P" + (SalesProgramIndex - 1).ToString() + "\")";
                                        var validationCell = dtSheet.DataValidations.AddListValidation(svaddress);
                                        validationCell.AllowBlank = true;
                                        validationCell.Formula.ExcelFormula = salesPorgramAddress;
                                        validationCell.ShowErrorMessage = true;
                                        validationCell.Error = "Please select from dropdown list";
                                        WriteToLog(DateSheetCollection[i].ToShortDateString() + "   " + salesPorgramAddress);
                                    }
                                }
                                WriteToLog("sales program end");

                                /*if (!string.IsNullOrEmpty(CatalogSalesProgramAddress))
                                {
                                    var salesPorgramAddress = "=INDIRECT(\"GlobalVariables!" + CatalogSalesProgramAddress + "\")";
                                    var validationCell = dtSheet.DataValidations.AddListValidation(svaddress);
                                    validationCell.AllowBlank = true;
                                    validationCell.Formula.ExcelFormula = salesPorgramAddress;
                                    validationCell.ShowErrorMessage = true;
                                    validationCell.Error = "Please select from dropdown list";
                                    WriteToLog(DateSheetCollection[i].ToShortDateString() + "   " + salesPorgramAddress);
                                }*/
                            }
                            else
                            {
                                //remove setting in B7 AND C7
                                if (!string.IsNullOrEmpty(snaddress))
                                    dtSheet.Cells[snaddress].Value = string.Empty;
                                if (!string.IsNullOrEmpty(svaddress))
                                {
                                    dtSheet.Cells[svaddress].Style.Locked = true;
                                    dtSheet.Cells[svaddress].Style.Border.Bottom.Style = ExcelBorderStyle.None;
                                    dtSheet.Cells[svaddress].Merge = false;
                                }
                            }

                            XLSMPck.Workbook.Worksheets.MoveAfter(dtSheet.Name, lastSheetName);
                            lastSheetName = dtSheet.Name;

                            WriteToLog("Date Sheet " + dtSheet.Name + " ATP Check, Add Comments etc. begin:" + DateTime.Now.TimeOfDay);
                            PresetForProductSheet(dtSheet, skuSheet, atpSheet, DateSheetCollection[i], colors, AddThresholdQtyComments);
                            WriteToLog("Date Sheet " + dtSheet.Name + " ATP Check, Add Comments etc. end:" + DateTime.Now.TimeOfDay);

                            if (imageTB != null && (string.IsNullOrEmpty(ConfigurationManager.AppSettings["StandardAddImageInTemp"]) || ConfigurationManager.AppSettings["StandardAddImageInTemp"] != "1"))
                            {
                                WriteToLog("Date Sheet " + dtSheet.Name + " Save Image to Sheet begin:" + DateTime.Now.TimeOfDay);
                                SaveImagesToOffline(imageTB, dtSheet);
                                WriteToLog("Date Sheet " + dtSheet.Name + " Save Image to Sheet end:" + DateTime.Now.TimeOfDay);
                            }

                            var StandardFreezeColumn = ConfigurationManager.AppSettings["StandardFreezeColumn"];
                            if (!string.IsNullOrEmpty(StandardFreezeColumn))
                                dtSheet.View.FreezePanes(1, int.Parse(StandardFreezeColumn));
                            //var StandardOrderAutoFitColumn = ConfigurationManager.AppSettings["StandardOrderAutoFitColumn"];
                            //if (!string.IsNullOrEmpty(StandardOrderAutoFitColumn))
                            //    dtSheet.Cells[1, int.Parse(StandardOrderAutoFitColumn), 1, int.Parse(StandardOrderAutoFitColumn) + 50].AutoFitColumns();

                            var StandardFilterAddress = ConfigurationManager.AppSettings["StandardFilterAddress"];
                            if (!string.IsNullOrEmpty(StandardFilterAddress))
                            {
                                var lastrowNumber = dtSheet.LastRowNum();
                                if (lastrowNumber > 0)
                                {
                                    StandardFilterAddress = string.Format(StandardFilterAddress, lastrowNumber.ToString());
                                    dtSheet.Cells[StandardFilterAddress].AutoFilter = true;
                                }
                            }

                            var ProductDetailFreezeCell = OLOFSettings.Settings.GetValueByKey("ProductDetailFreezeCell", "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                            if (ProductDetailFreezeCell.Length > 0)
                            {
                                var freezeRow = int.Parse(ProductDetailFreezeCell[0]);
                                var freezeCol = int.Parse(ProductDetailFreezeCell[1]);
                                dtSheet.View.FreezePanes(freezeRow, freezeCol);
                            }

                            if (System.IO.File.Exists(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\DateSheet.vba")))
                                dtSheet.CodeModule.Code = File.ReadAllText(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\DateSheet.vba"));

                            //EC QAS/PRD: Add Default Zoom to OLOF and Display Total Units near top of OLOF #232
                            var CategorySKUSheetZoom = ConfigurationManager.AppSettings["CategorySKUSheetZoom"];
                            var zoomScale = 0;
                            if (int.TryParse(CategorySKUSheetZoom, out zoomScale))
                            {
                                if (zoomScale > 0)
                                    dtSheet.View.ZoomScale = zoomScale;
                            }

                            dtSheet.Protection.IsProtected = true;
                            dtSheet.Protection.AllowAutoFilter = true;
                            dtSheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
                        }
                        catch (Exception ex)
                        {
                            WriteToLog(ex.Message);
                            WriteToLog(ex.StackTrace);
                        }
                    }
                }

                //change for crocs
                /*for (int i = 0; i < UPCSheets.Count; i++)
                {
                    var upcSheet = XLSMPck.Workbook.Worksheets[UPCSheets[i]];
                    if (upcSheet != null && System.IO.File.Exists(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\UPCSheet.vba")))
                    {
                        upcSheet.CodeModule.Code = File.ReadAllText(Path.Combine(System.Windows.Forms.Application.StartupPath, "VBA\\UPCSheet.vba"));

                        //upcSheet.Protection.IsProtected = true;
                        //upcSheet.Protection.AllowAutoFilter = true;
                        //upcSheet.Protection.SetPassword(ConfigurationManager.AppSettings["SheetPassword"] ?? "Plumriver");
                    }
                }*/

                if (ConfigurationManager.AppSettings["MACO365"] == "1")
                {
                    //SetCellValue(globalsheet, 201, 2, dftSheet.Name);
                    var cell = globalsheet.Cells[202, 2];
                    cell.Style.Numberformat.Format = "@";
                    cell.Value = dftSheet.Name;
                    var sheetTemp = ConfigurationManager.AppSettings["SheetTemplate365"];
                    if (!string.IsNullOrEmpty(sheetTemp))
                    {
                        var dsheet = XLSMPck.Workbook.Worksheets.Add(sheetTemp);
                        dsheet.View.TabSelected = true;
                    }
                    else
                    {
                        var dsheet = XLSMPck.Workbook.Worksheets[UPCSheets[0]];
                        dsheet.View.TabSelected = true;
                    }
                }
                else
                    dftSheet.View.TabSelected = true;
                tmpSheet.Hidden = eWorkSheetHidden.Hidden;
            }
            catch (Exception ex)
            {
                WriteToLog(ex.Message);
                WriteToLog(ex.StackTrace);
                if (ex.InnerException != null)
                {
                    WriteToLog(ex.InnerException.Message);
                    WriteToLog(ex.InnerException.StackTrace);
                }
                throw ex;
            }
        }

        private void AddLogoToDateSheet(ExcelWorksheet tmpSheet)
        {
            var logoPath = ConfigurationManager.AppSettings["LogoPath"];
            Image img = Image.FromFile(logoPath);
            if (img != null)
            {
                //ExcelPicture pic = tmpSheet.Drawings.AddPicture("logo", img);
                ExcelPicture pic = tmpSheet.Drawings.AddPicture("logo" + System.Guid.NewGuid().ToString(), new FileInfo(logoPath));
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
            var ImageFormat = string.IsNullOrEmpty(ConfigurationManager.AppSettings["ImageFormat"]) ? ".png" : ConfigurationManager.AppSettings["ImageFormat"];
            string[] ignoreStyles = (ConfigurationManager.AppSettings["IgnoreStyles"] ?? "").Split(new char[] { ',' });
            foreach (DataRow row in dt.Rows)
            {
                var rowIndex = (int)row["ROW"];
                row["ImageName"] = row["ImageName"].ToString().Replace(".jpg", ImageFormat);
                var imageName = row["ImageName"].ToString();
                Path.GetInvalidFileNameChars().ToList().ForEach(c => { imageName = imageName.Replace(c.ToString(), ""); });
                imageName = string.IsNullOrEmpty(new FileInfo(imageName).Extension) ? imageName + ImageFormat : imageName;

                var imagePath = Path.Combine(imageFolder, imageName);
                if (ignoreStyles.Contains(imageName.Replace(ImageFormat, "")))
                {
                    if(tmpSheet.Row(rowIndex).Height != 45)
                        tmpSheet.Row(rowIndex).Height = 30;
                }
                else if (File.Exists(imagePath))
                {
                    try
                    {
                        Image prdimg = Image.FromFile(imagePath);
                        tmpSheet.Row(rowIndex).Height = prdimg.Height * 0.9;
                    }
                    catch (Exception ex)
                    {
                        WriteToLog(ex.Message);
                        WriteToLog(ex.StackTrace);
                        WriteToLog(imagePath);
                        continue;
                    }

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
                Path.GetInvalidFileNameChars().ToList().ForEach(c => { imageName = imageName.Replace(c.ToString(), ""); });
                imageName = string.IsNullOrEmpty(new FileInfo(imageName).Extension) ? imageName + ImageFormat : imageName;

                var imagePath = Path.Combine(imageFolder, imageName);
                if (ignoreStyles.Contains(imageName.Replace(ImageFormat, "")))
                    continue;
                if (File.Exists(imagePath))
                {
                    if (imageName.Contains(" "))
                    {
                        var newimagePath = Path.Combine(imageFolder, imageName.Replace(" ", "_"));
                        System.IO.File.Copy(imagePath, newimagePath, true);
                        imagePath = newimagePath;
                    }
                    ExcelPicture pic = null;
                    try
                    {
                        pic = tmpSheet.Drawings.AddPicture(System.Guid.NewGuid().ToString(), new FileInfo(imagePath));
                    }
                    catch (Exception ex)
                    {
                        WriteToLog(ex.Message);
                        WriteToLog(ex.StackTrace);
                        WriteToLog(imagePath);
                        continue;
                    }
                    pic.From.Column = 0;
                    pic.From.Row = rowIndex - 1;
                    pic.From.ColumnOff = Pixel2MTU(10);
                    pic.From.RowOff = Pixel2MTU(5);
                    pic.EditAs = eEditAs.TwoCell;

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
                var ATPQtyThreshold = GetB2BSetting("ATP", "ATPQtyThreshold");
                if (string.IsNullOrEmpty(ATPQtyThreshold))
                    ATPQtyThreshold = ConfigurationManager.AppSettings["ThresholdQty"] ?? "99";
                int qty;
                if (int.TryParse(ATPQtyThreshold, out qty))
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
                            if (availbleQty > 0 || AllowBackOrder == "1")
                            {
                                //var list1 = cell.DataValidation.AddCustomDataValidation();
                                //list1.Formula.ExcelFormula = string.Format("(MOD(indirect(address(row(),column())) ,{0})=0)", multipleValue);
                                //list1.ShowErrorMessage = true;
                                //list1.Error = string.Format("You must enter a multiple of {0} in this cell.", multipleValue).Replace(".00 ", " ").Replace(".0 ", " ");

                                if (ThresholdQty.HasValue && availbleQty > 0)
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
                                if(availbleQty == 0 && AllowBackOrder == "1")
                                    availbleQty = 999999;
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
            var StandardHideNavigationWithFilter = ConfigurationManager.AppSettings["StandardHideNavigationWithFilter"];
            if (string.IsNullOrEmpty(StandardHideNavigationWithFilter) || StandardHideNavigationWithFilter != "1")
            {
                if (!multicolumndeptlabel)   //Put the whole dept hierarchy in the first header cell //heading brakes ">"
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
            }
            else
            {
                //DO NOT SHOW THE NAVIGATION, WILL SAVE BRAND, GENDER, COLOR, DESCRIPTION ETC IN THESE CELLS IN ADD ROWS METHOD -- FOR CROCUS
                var FilterColumns = (headers.Length == 0 ? 4 : (2 + headers.Length));
                for (int i = 1; i < FilterColumns; i++)
                {
                    ExcelRange er = sheet.Cells[i.IntToMoreChar() + sheetRow.ToString()];
                    if (!groupedTitle.HasValue || groupedTitle == false)
                    {
                        er.Value = string.Empty;
                        er.Style.Font.Color.SetColor(Color.White);
                    }
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
                    cell = sheet.Cells[string.Format("{0}{1}:{0}{2}", (cellIndex + (j++)).IntToMoreChar(), sheetRow, sheetRow)]; // + 1)];
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
            string displayGroupTitle = ConfigurationManager.AppSettings["DisplayGroupTitle"] ?? "1";
            //groupedTitle = true THE TOP LEVEL DEPARTMENT
            //DisplaySubTitle = 1 THE LOWEST LEVEL DEPARTMENT WILL SHOW THE SIZE
            //ADD SIZE LIST IN DEPARTMENT LEVEL, NAVIGATION, ATTRIBUTEVALUE1
            foreach (DictionaryEntry de in cols)
            {
                if (!groupedTitle.HasValue || (groupedTitle == true && displayGroupTitle == "1") || (displaySubTitle == "1" && groupedTitle == false))
                {
                    cell = sheet.Cells[(colcounter).IntToMoreChar() + sheetRow.ToString()];
                    cell.Value = (string.IsNullOrEmpty(Convert.ToString(de.Value)) ? " " : StripHTML(Convert.ToString(de.Value)).Replace("\r", System.Environment.NewLine));
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.WrapText = true;
                    if (displaySubTitle == "1" && groupedTitle == false)
                        cell.Style.Font.Color.SetColor(Color.Gray);
                    //var StandardOrderCellWidth = ConfigurationManager.AppSettings["StandardOrderCellWidth"];
                    //if (!string.IsNullOrEmpty(StandardOrderCellWidth))
                    //    sheet.Column(colcounter).Width = int.Parse(StandardOrderCellWidth);

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
                        var StandardAddZeroSKURow = string.IsNullOrEmpty(ConfigurationManager.AppSettings["StandardAddZeroSKURow"]) ? "0" : ConfigurationManager.AppSettings["StandardAddZeroSKURow"];
                        if (skus.Count > 0 || StandardAddZeroSKURow == "1")
                        {
                            int col = AddRow2(depts, rowattribs, skus, colpositions, sheet, skusheet, pricesheet, atpsheet, upcsheet, OrderMultiplesheet, firstrow, lastrow);

                            maxColIdx = col > maxColIdx ? col : maxColIdx;
                            firstrow = false;
                        }
                        if (skus.Count <= 0)
                        {
                            WriteToLog(deptid.ToString());
                            WriteToLog(rowattribs["Style"].ToString());
                            var groupKey = rowattribs.ContainsKey("GroupKey") ? rowattribs["GroupKey"].ToString() : string.Empty;
                            WriteToLog("There's no enabled SKUs (Enabled = 1) in Department " + deptid.ToString() + " with Style " + rowattribs["Style"].ToString() + " " + groupKey);
                        }
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
            try
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

                var attr2 = string.Empty;
                var attr3 = string.Empty;
                var attr4 = string.Empty;
                var attr5 = string.Empty;
                var unitprice = string.Empty;
                var upc = string.Empty;
                var styleName = rowattribs["ProductName"] == null ? rowattribs["Style"].ToString() : rowattribs["ProductName"].ToString();
                var style = rowattribs["Style"].ToString();

                if (skus.Count > 0)
                {
                    var firstitem = skus.Cast<DictionaryEntry>().ElementAt(0);
                    attr2 = ((StringDictionary)(firstitem.Value)).ContainsKey("AttributeValue2") ? ((StringDictionary)(firstitem.Value))["AttributeValue2"].ToString() : string.Empty;
                    attr3 = ((StringDictionary)(firstitem.Value)).ContainsKey("AttributeValue3") ? ((StringDictionary)(firstitem.Value))["AttributeValue3"].ToString() : string.Empty;
                    attr4 = ((StringDictionary)(firstitem.Value)).ContainsKey("AttributeValue4") ? ((StringDictionary)(firstitem.Value))["AttributeValue4"].ToString() : string.Empty;
                    attr5 = ((StringDictionary)(firstitem.Value)).ContainsKey("AttributeValue5") ? ((StringDictionary)(firstitem.Value))["AttributeValue5"].ToString() : string.Empty;
                    upc = ((StringDictionary)(firstitem.Value))["UPC"].ToString();
                    unitprice = ((StringDictionary)(firstitem.Value))["PriceWholesale"].ToString();
                }
                else
                {
                    attr2 = rowattribs.ContainsKey("AttributeValue2") ? rowattribs["AttributeValue2"].ToString() : string.Empty;
                    attr3 = rowattribs.ContainsKey("AttributeValue3") ? rowattribs["AttributeValue3"].ToString() : string.Empty;
                    attr4 = rowattribs.ContainsKey("AttributeValue4") ? rowattribs["AttributeValue4"].ToString() : string.Empty;
                    attr5 = rowattribs.ContainsKey("AttributeValue5") ? rowattribs["AttributeValue5"].ToString() : string.Empty;
                    unitprice = rowattribs.ContainsKey("PriceWholesale") ? rowattribs["PriceWholesale"].ToString() : string.Empty;
                }


                var lines = OLOFSettings.Settings.GetValueByKey("ProductDetailLines", "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                var StandardHideNavigationWithFilter = ConfigurationManager.AppSettings["StandardHideNavigationWithFilter"];
                var ProductFilterLines = OLOFSettings.Settings.GetValueByKey("ProductFilterLines", "").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
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
                            case "attr5":
                                cellvalue = attr5;
                                break;
                            case "styleName":
                                cellvalue = styleName;
                                break;
                            default:
                                if (rowattribs.ContainsKey(lines[i]))
                                    cellvalue = rowattribs[lines[i]];
                                break;
                        }

                        using (ExcelRange r = sheet.Cells[rowIndex, cellIndex])
                        {
                            var ProductDetailWidth = OLOFSettings.Settings.GetValueByKey("ProductDetailWidth", "");
                            if (!string.IsNullOrEmpty(ProductDetailWidth))
                                r.Style.WrapText = true;
                            r.Style.Font.SetFromFont(LineFontStyle.ToFont());
                            r.Value = cellvalue;

                            var launchDate = ConfigurationManager.AppSettings["DisplayLaunchDateText"];
                            if (lines[i] == "styleName" && launchDate == "1" && skus.Count > 0)
                            {
                                var firstitem = skus.Cast<DictionaryEntry>().ElementAt(0);
                                var attr5LaunchDateText = ((StringDictionary)(firstitem.Value)).ContainsKey("AttributeValue5") ? ((StringDictionary)(firstitem.Value))["AttributeValue5"].ToString() : string.Empty;
                                if (!string.IsNullOrEmpty(attr5LaunchDateText.Trim()))
                                {
                                    r.Style.WrapText = true;
                                    r.Value = cellvalue + "\n";
                                    var rtDir2 = r.RichText.Add(attr5LaunchDateText);
                                    rtDir2.Color = Color.Red;
                                    rtDir2.Bold = true;
                                    rtDir2.Size = 9;

                                    sheet.Row(rowIndex).Height = 45;
                                }
                            }
                        }

                        //DO NOT SHOW THE NAVIGATION IN DEPARTMENT ROW, SAVE BRAND, GENDER, COLOR, DESCRIPTION ETC IN THESE CELLS IN ADD ROWS METHOD -- FOR CROCUS
                        if (firstrow && StandardHideNavigationWithFilter == "1")
                        {

                            using (ExcelRange r = sheet.Cells[rowIndex - 1, cellIndex])
                            {
                                if (ProductFilterLines.Length <= 0 || ProductFilterLines.Contains(lines[i]))
                                    r.Value = cellvalue;
                                else
                                    r.Value = string.Empty;
                                r.Style.Font.Color.SetColor(Color.White);
                            }
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
                var msrp = !rowattribs.ContainsKey("MSRP") ? string.Empty : rowattribs["MSRP"].ToString();
                var ProductShowMSRPInPrice = OLOFSettings.Settings.GetValueByKey("ProductShowMSRPInPrice", "");
                if (string.IsNullOrEmpty(ProductShowMSRPInPrice) || ProductShowMSRPInPrice == "0" || string.IsNullOrEmpty(msrp))
                    SetCellValue(sheet, rowIndex, cellIndex, price, true, true).Style.Font.SetFromFont(LineFontStyle.ToFont());
                else
                {
                    var rowPrice = price.ToString("N2") + this.Currency + "/" + msrp + this.Currency;
                    SetCellValue(sheet, rowIndex, cellIndex, rowPrice, false).Style.Font.SetFromFont(LineFontStyle.ToFont());
                }
                cellIndex++;

                //TTL
                var TTLIndex = SKUTotalUnitColumn = cellIndex;
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

                        //This is the column header value
                        string attr1 = ((StringDictionary)(de.Value))["AttributeValue1"].ToString();
                        string cellposition = "";
                        //If the row contains a cell with the column header value, enable it
                        attr1 = StripHTML(Convert.ToString(attr1));
                        //SKU must be enabled
                        if (((StringDictionary)(de.Value))["Enabled"].ToString() == "1")
                        {
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
                            }
                        }

                        if (upcsheet != null)
                        {
                            var StandardUPCSheetWithCurrency = ConfigurationManager.AppSettings["StandardUPCSheetWithCurrency"];
                            var lineItems = OLOFSettings.Settings.GetValueByKey("CatalogUPCSheetLine", "prdName|style|attr3|attr1|upc|unitprice").Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                            var ShowUPCFromSKUValue = string.IsNullOrEmpty(ConfigurationManager.AppSettings["ShowUPCFromSKUValue"]) ? "1" : ConfigurationManager.AppSettings["ShowUPCFromSKUValue"];
                            if (ShowUPCFromSKUValue == "1")
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
                                        {
                                            if (StandardUPCSheetWithCurrency == "1")
                                                SetCellValue(upcsheet, upcRowIndex, i + 1, unitprice + " " + Currency, false);
                                            else
                                                cell = SetCellValue(upcsheet, upcRowIndex, i + 1, unitprice, false);
                                            break;
                                        }
                                    case "MSRP":
                                        {
                                            var cellvalue = string.Empty;
                                            if (((StringDictionary)(de.Value)).ContainsKey("MSRP"))
                                                cellvalue = ((StringDictionary)(de.Value))["MSRP"].ToString();
                                            if (!string.IsNullOrEmpty(cellvalue) && !string.IsNullOrEmpty(StandardUPCSheetWithCurrency) && StandardUPCSheetWithCurrency == "1")
                                                cell = SetCellValue(upcsheet, upcRowIndex, i + 1, cellvalue + " " + Currency, false);
                                            else
                                                cell = SetCellValue(upcsheet, upcRowIndex, i + 1, cellvalue, false);
                                            break;
                                        }
                                    case "SKU":
                                        {
                                            var upcSKU = ((StringDictionary)(de.Value))["QuickWebSKUValue"].ToString();
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, upcSKU, false);
                                            break;
                                        }
                                    default:
                                        {
                                            var cellvalue = string.Empty;
                                            if (rowattribs.ContainsKey(lineItems[i]))
                                                cellvalue = rowattribs[lineItems[i]];
                                            cell = SetCellValue(upcsheet, upcRowIndex, i + 1, cellvalue, false);
                                            break;
                                        }
                                }
                            }
                            upcRowIndex++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    WriteToLog("foreach (DictionaryEntry de in skus)" + rowattribs["PfID"]);
                    throw;
                }
                //Assign the formulas
                if (ttlformula.Length > 0)
                    ttlcell.Formula = ttlformula.Substring(0, ttlformula.Length - 1);
                if (valueformula.Length > 0)
                    valuecell.Formula = valueformula.Substring(0, valueformula.Length - 1);

                return cellIndex;
            }
            catch (Exception ex)
            {
                WriteToLog("AddRow2");
                WriteToLog(ex.Message);
                WriteToLog(ex.StackTrace);
                throw ex;
            }
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
                if (drow["SKU"].ToString().Trim() == SKU.Trim())
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
            //WriteToLog("ATPValueData IS " + cellValue + " " + SKU);
            return cell;
        }

        private string CleanFileName(string fileName)
        {
            return Path.GetInvalidFileNameChars().Aggregate(fileName, (current, c) => current.Replace(c.ToString(), ""));
        }

        private void CheckLoadSalesProgram(string catalog, string soldto, List<DateTime> DateSheetCollection)
        {
            try
            {
                CatalogSupportSalesProgram = false;
                CatalogSalesPrograms = new List<SalesProgram>();
                CatalogSalesProgramAddress = string.Empty;
                if (ConfigurationManager.AppSettings["CatalogNeedSalesProgram"] == "1")
                {
                    var support = GetSalesProgramLink(catalog, soldto);
                    CatalogSupportSalesProgram = support == 1 ? true : false;
                    if (CatalogSupportSalesProgram)
                    {
                        LoadSalesProgramList(catalog, soldto, DateSheetCollection);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToLog(ex.Message);
                WriteToLog(ex.StackTrace);
                throw ex;
            }
        }

        private int GetSalesProgramLink(string catalogCode, string CustomerId)
        {

            //SqlParameter paramShopperId = new SqlParameter("@ShopperID", SqlDbType.UniqueIdentifier);
            SqlParameter paramCatalog = new SqlParameter("@Catalog", SqlDbType.VarChar, 80);
            SqlParameter paramSoldTo = new SqlParameter("@SoldTo", SqlDbType.NVarChar, 80);
            SqlParameter paramDisplayLink = new SqlParameter("@DisplaySalesProgramLink", SqlDbType.SmallInt);//@DisplaySalesProgramLink

            //paramShopperId.Value = new Guid(shopperId);
            paramCatalog.Value = catalogCode;
            paramSoldTo.Value = CustomerId;
            paramDisplayLink.Direction = ParameterDirection.Output;

            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_DisplaySelectSalesProgramLink", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddRange(new SqlParameter[] { paramCatalog, paramSoldTo, paramDisplayLink });//paramShopperId
                cmd.ExecuteNonQuery();
                return Convert.ToInt32(paramDisplayLink.Value);
            }
        }

        private void LoadSalesProgramList(string catalogCode, string CustomerId, List<DateTime> DateSheetCollection)
        {
            var xmlvalue = string.Empty;

            SqlParameter paramCatalog = new SqlParameter("@Catalog", SqlDbType.VarChar, 80);
            SqlParameter paramSoldTo = new SqlParameter("@SoldTo", SqlDbType.NVarChar, 80);
            paramCatalog.Value = catalogCode;
            paramSoldTo.Value = CustomerId;

            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_GetCatalogSalesPrograms", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddRange(new SqlParameter[] { paramCatalog, paramSoldTo });

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        xmlvalue += reader.GetValue(0).ToString();
                    }
                }
            }

            if (!string.IsNullOrEmpty(xmlvalue))
            {
                /*var minReqShipDate = DateTime.MaxValue;
                var maxReqShipDate = DateTime.MinValue;
                foreach (var date in DateSheetCollection)
                {
                    if (minReqShipDate > date)
                        minReqShipDate = date;
                    if (maxReqShipDate < date)
                        maxReqShipDate = date;
                }*/

                StringReader reader = new StringReader(xmlvalue);
                var ds = new DataSet();
                ds.ReadXml(reader);

                var stable = ds.Tables["SalesProgram"];
                foreach (DataRow row in stable.Rows)
                {
                    var salesprogram = new SalesProgram
                    {
                        SalesProgramCode = row["code"].ToString(),
                        ShipStartDate = DateTime.Parse(row["DateShipStart"].ToString()),
                        ShipEndDate = DateTime.Parse(row["DateShipEnd"].ToString())
                    };
                    //if (salesprogram.ShipStartDate <= minReqShipDate && salesprogram.ShipEndDate >= maxReqShipDate)
                    CatalogSalesPrograms.Add(salesprogram);
                }
            }
        }



        protected ExcelRange SetCellValue(ExcelWorksheet sheet, string address, string value)
        {
            return SetCellValue(sheet, address, value, string.Empty);
        }

        protected ExcelRange SetCellValue(ExcelWorksheet sheet, string address, string value, string dataType)
        {
            ExcelRange cell = sheet.Cells[address];
            switch (dataType)
            {
                case "1": //int
                    cell.Style.Numberformat.Format = "#,##0";
                    if (!string.IsNullOrEmpty(value))
                        cell.Value = int.Parse(value);
                    break;
                case "2": //double
                    cell.Style.Numberformat.Format = "#,##0.00";
                    if (!string.IsNullOrEmpty(value))
                        cell.Value = double.Parse(value);
                    break;
                case "3":
                    cell.Style.Numberformat.Format = "mm-dd-yy";
                    if (!string.IsNullOrEmpty(value))
                        cell.Value = value;
                    break;
                default:
                    cell.Value = value;
                    break;
            }

            return cell;
        }

        protected ExcelRange GetCell(ExcelWorksheet sheet, string address)
        {
            var cell = sheet.Cells[address];
            return cell;
        }

        protected void SetCellStyle(ExcelRange cell)
        {
            SetCellStyle(cell, null);
        }

        protected void SetCellStyle(ExcelRange cell, System.Drawing.Color? bgColor)
        {
            SetCellStyle(cell, bgColor, true, true, true, true);
        }

        protected void SetCellStyle(ExcelRange cell, System.Drawing.Color? bgColor, bool borderLeft, bool borderRight, bool borderTop, bool borderBottom)
        {
            if (bgColor != null)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor((System.Drawing.Color)bgColor);
            }
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            if (!borderLeft)
                cell.Style.Border.Left.Style = ExcelBorderStyle.None;
            if (!borderRight)
                cell.Style.Border.Right.Style = ExcelBorderStyle.None;
            if (!borderTop)
                cell.Style.Border.Top.Style = ExcelBorderStyle.None;
            if (!borderBottom)
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.None;
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

    public class SalesProgram
    {
        public string SalesProgramCode { get; set; }
        public DateTime ShipStartDate { get; set; }
        public DateTime ShipEndDate { get; set; }
    }
}
