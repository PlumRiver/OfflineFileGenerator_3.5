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
using System.Xml;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;

namespace OfflineFileGenerator
{
    public class CategoryTabularFormGenerator
    {
        public string PriceType { get; set; }
        public bool BookingOrder { get; set; }
        public string ProgramCode { get; set; }

        public string OfflineTemplateStyle
        {
            get
            {
                return ConfigurationManager.AppSettings["OfflineTemplateStyle"] == null ?
                    "Normal" : ConfigurationManager.AppSettings["OfflineTemplateStyle"].ToString();
            }
        }

        //Configurations in Summary Sheet
        string SummarySheetName = ConfigurationManager.AppSettings["SummarySheetName"] == null ? 
            "Summary" : ConfigurationManager.AppSettings["SummarySheetName"];
        string CatalogCodeLocation = ConfigurationManager.AppSettings["SummaryCatalogCodeLocation"] == null ?
            "A|6" : ConfigurationManager.AppSettings["SummaryCatalogCodeLocation"];
        string CatalogNameLocation = ConfigurationManager.AppSettings["SummaryCatalogNameLocation"] == null ?
            "C|6" : ConfigurationManager.AppSettings["SummaryCatalogNameLocation"];
        uint SummaryCategoryStartRow = ConfigurationManager.AppSettings["SummaryCategoryStartRow"] == null ?
            22 : uint.Parse(ConfigurationManager.AppSettings["SummaryCategoryStartRow"]);
        string SummaryCategoryColumns = ConfigurationManager.AppSettings["SummaryCategoryColumns"] == null ?
            "C|D|E|F|G|H|I|J" : ConfigurationManager.AppSettings["SummaryCategoryColumns"];
        string SummaryCurrencyColumn = ConfigurationManager.AppSettings["SummaryCurrencyColumn"] == null ?
            "K" : ConfigurationManager.AppSettings["SummaryCurrencyColumn"];
        uint SummaryOrderTotalUnitsRow = ConfigurationManager.AppSettings["SummaryOrderTotalUnitsRow"] == null ?
            20 : uint.Parse(ConfigurationManager.AppSettings["SummaryOrderTotalUnitsRow"]);
        uint SummaryOrderTotalValueRow = ConfigurationManager.AppSettings["SummaryOrderTotalValueRow"] == null ?
            21 : uint.Parse(ConfigurationManager.AppSettings["SummaryOrderTotalValueRow"]);
        uint bgfillid1 = ConfigurationManager.AppSettings["SummaryBGFillid1"] == null ?
            6 : uint.Parse(ConfigurationManager.AppSettings["SummaryBGFillid1"]);
        uint bgfillid2 = ConfigurationManager.AppSettings["SummaryBGFillid2"] == null ?
            2 : uint.Parse(ConfigurationManager.AppSettings["SummaryBGFillid2"]);
        uint bolderfont = ConfigurationManager.AppSettings["SummaryBolderFont"] == null ?
            7 : uint.Parse(ConfigurationManager.AppSettings["SummaryBolderFont"]);
        uint numformatid = ConfigurationManager.AppSettings["SummaryNumFormatID"] == null ?
            178 : uint.Parse(ConfigurationManager.AppSettings["SummaryNumFormatID"]);
        uint intformatid = ConfigurationManager.AppSettings["SummaryIntFormatID"] == null ?
            179 : uint.Parse(ConfigurationManager.AppSettings["SummaryIntFormatID"]);
        uint pinkfileid = ConfigurationManager.AppSettings["SummaryPinkFileID"] == null ?
            5 : uint.Parse(ConfigurationManager.AppSettings["SummaryPinkFileID"]);
        string SummaryPaymentTermsLocation = ConfigurationManager.AppSettings["SummaryPaymentTermsLocation"] == null ?
            "D|11" : ConfigurationManager.AppSettings["SummaryPaymentTermsLocation"];
        string SummaryProgramCodeLocation = ConfigurationManager.AppSettings["SummaryProgramCodeLocation"] == null ?
            "D|12" : ConfigurationManager.AppSettings["SummaryProgramCodeLocation"];

        //Sheet Name
        string CategoryTemplateSheetName = ConfigurationManager.AppSettings["CategoryTemplateSheetName"] == null ?
            "CategoryTemplate" : ConfigurationManager.AppSettings["CategoryTemplateSheetName"];
        string CategoryListSheetName = ConfigurationManager.AppSettings["CategoryListSheetName"] == null ?
            "CategoryList" : ConfigurationManager.AppSettings["CategoryListSheetName"];

        //Configuration in Category Sheets
        uint startRow = ConfigurationManager.AppSettings["CategorySKUStartRow"] == null ? 
            11 : uint.Parse(ConfigurationManager.AppSettings["CategorySKUStartRow"]);
        string TabularNumericValidationCell = ConfigurationManager.AppSettings["TabularNumericValidationCell"] == null ? 
            "D1" : ConfigurationManager.AppSettings["TabularNumericValidationCell"];
        string OrderMultipleCheck = ConfigurationManager.AppSettings["OrderMultipleCheck"] == null ? 
            "0" : ConfigurationManager.AppSettings["OrderMultipleCheck"];
        string CategorySKUComColumnList = ConfigurationManager.AppSettings["CategorySKUComColumnList"] == null ?
            "A|Catalog,B|Level2DeptID,C|SKU,D|UPC,E|Department,F|Style,G|ProductName,H|AttributeValue2,I|AttributeValue1" : ConfigurationManager.AppSettings["CategorySKUComColumnList"];

        string CategorySKUTotalUnitsLocation = ConfigurationManager.AppSettings["CategorySKUTotalUnitsLocation"] == null ?
            "G|8" : ConfigurationManager.AppSettings["CategorySKUTotalUnitsLocation"];
        uint CategorySKUSubTotalRow = ConfigurationManager.AppSettings["CategorySKUSubTotalRow"] == null ?
            9 : uint.Parse(ConfigurationManager.AppSettings["CategorySKUSubTotalRow"]);
        string CategorySKUPriceColumn = ConfigurationManager.AppSettings["CategorySKUPriceColumn"] == null ?
            "J" : ConfigurationManager.AppSettings["CategorySKUPriceColumn"];
        string CategorySKUOrderColumnList = ConfigurationManager.AppSettings["CategorySKUOrderColumnList"] == null ?
            "K|L|M|N|O|P" : ConfigurationManager.AppSettings["CategorySKUOrderColumnList"];
        string CategorySKUUnitTotalColumn = ConfigurationManager.AppSettings["CategorySKUUnitTotalColumn"] == null ?
            "Q" : ConfigurationManager.AppSettings["CategorySKUUnitTotalColumn"];
        string CategorySKUValueTotalColumn = ConfigurationManager.AppSettings["CategorySKUValueTotalColumn"] == null ?
            "R" : ConfigurationManager.AppSettings["CategorySKUValueTotalColumn"];
        string CategotySKUCurrencyLocations = ConfigurationManager.AppSettings["CategorySKUCurrencyLocations"] == null ?
            "H|9,Q|9" : ConfigurationManager.AppSettings["CategorySKUCurrencyLocations"];
        string CategorySKUSummaryColumnList = ConfigurationManager.AppSettings["CategorySKUSummaryColumnList"] == null ?
            "A|B|C|D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R" : ConfigurationManager.AppSettings["CategorySKUSummaryColumnList"];
        string CategorySKUSummaryCategoryColumn = ConfigurationManager.AppSettings["CategorySKUSummaryCategoryColumn"] == null ?
            "E" : ConfigurationManager.AppSettings["CategorySKUSummaryCategoryColumn"];
        string CategorySKUSummaryCategoryIDColumn = ConfigurationManager.AppSettings["CategorySKUSummaryCategoryIDColumn"] == null ?
            "B" : ConfigurationManager.AppSettings["CategorySKUSummaryCategoryIDColumn"];
        uint blackFillID = ConfigurationManager.AppSettings["CategoryBlackFillID"] == null ?
            3 : uint.Parse(ConfigurationManager.AppSettings["CategoryBlackFillID"]);

        string CategoryTotalUnitsTitleCell = ConfigurationManager.AppSettings["CategoryTotalUnitsTitleCell"] == null ?
            "F|8" : ConfigurationManager.AppSettings["CategoryTotalUnitsTitleCell"];
        string CategoryTotalValueTitleCell = ConfigurationManager.AppSettings["CategoryTotalValueTitleCell"] == null ?
            "F|9" : ConfigurationManager.AppSettings["CategoryTotalValueTitleCell"];
        uint CategoryTabTotalFillId = ConfigurationManager.AppSettings["CategoryTabTotalFillId"] == null ?
            4 : uint.Parse(ConfigurationManager.AppSettings["CategoryTabTotalFillId"]);//YELLOW
        uint? CategoryTabTotalFontId = ConfigurationManager.AppSettings["CategoryTabTotalFontId"] == null ?
            (uint?)null : uint.Parse(ConfigurationManager.AppSettings["CategoryTabTotalFontId"]);
        string CategorySKUCombineProductName = ConfigurationManager.AppSettings["CategorySKUCombineProductName"] == null ?
            "1" : ConfigurationManager.AppSettings["CategorySKUCombineProductName"];
        string CategorySKUCombineDepartment = ConfigurationManager.AppSettings["CategorySKUCombineDepartment"] == null ?
            "1" : ConfigurationManager.AppSettings["CategorySKUCombineDepartment"];
        string CategorySKUCombineStyle = ConfigurationManager.AppSettings["CategorySKUCombineStyle"] == null ?
            "1" : ConfigurationManager.AppSettings["CategorySKUCombineStyle"];
        string CategorySKUCombineColor = ConfigurationManager.AppSettings["CategorySKUCombineColor"] == null ?
            "1" : ConfigurationManager.AppSettings["CategorySKUCombineColor"];

        string CategorySubTotalForStyle = ConfigurationManager.AppSettings["CategorySubTotalForStyle"] == null ?
            "0" : ConfigurationManager.AppSettings["CategorySubTotalForStyle"];
        string ShipDateRequiredForStart = ConfigurationManager.AppSettings["ShipDateRequiredForStart"] == null ?
            "0" : ConfigurationManager.AppSettings["ShipDateRequiredForStart"];

        string TabularCatalogUPCSheetName = ConfigurationManager.AppSettings["TabularCatalogUPCSheetName"] ?? "";
        bool CategoryWSPriceEditable = (ConfigurationManager.AppSettings["CategoryWSPriceEditable"] ?? "0") == "1";

        string CategorySKUAutoFilterRange = ConfigurationManager.AppSettings["CategorySKUAutoFilterRange"];

        string CategorySKUATPSheet = ConfigurationManager.AppSettings["CategorySKUATPSheet"] == null ?
            "0" : ConfigurationManager.AppSettings["CategorySKUATPSheet"];

        string CategoryEnableShipWindow = ConfigurationManager.AppSettings["CategoryEnableShipWindow"] == null ?
            "0" : ConfigurationManager.AppSettings["CategoryEnableShipWindow"];

        string SummaryShipDateRow = ConfigurationManager.AppSettings["SummaryShipDateRow"] == null ?
            "14" : ConfigurationManager.AppSettings["SummaryShipDateRow"];

        string SummaryEnableShipTo = string.IsNullOrEmpty(ConfigurationManager.AppSettings["SummaryEnableShipTo"]) ? 
            "0" : ConfigurationManager.AppSettings["SummaryEnableShipTo"];
        string SummaryEnableComments = string.IsNullOrEmpty(ConfigurationManager.AppSettings["SummaryEnableComments"]) ?
            "0" : ConfigurationManager.AppSettings["SummaryEnableComments"];
        string CustomerDataEnabled = string.IsNullOrEmpty(ConfigurationManager.AppSettings["CustomerDataEnabled"]) ?
            "0" : ConfigurationManager.AppSettings["CustomerDataEnabled"];

        private List<CategoryLevelInfo> CategoryLevelList { get; set; }
        private string CurrencyCode { get; set; }
        static uint CommonDateFormat = ConfigurationManager.AppSettings["CategoryCommonDateFormat"] == null ? 59 : uint.Parse(ConfigurationManager.AppSettings["CategoryCommonDateFormat"]);
        private uint CommonTextFormat { get; set; }
        private uint CommonBorderAllId { get; set; }
        private uint CommonFontId { get; set; }
        private uint CommonBoldTextId { get; set; }
        private List<ShipWindowInfo> OrderColumnList{ get; set; }

        public void GenerateCategoryTabularOrderForm(string templateFile, string filename, string soldto, string catalog, string pricecode, string savedirectory)
        {
            WriteToLog(DateTime.Now.ToString() + " Begin " + filename);
            if (File.Exists(templateFile))
            {
                try
                {
                    //save new file
                    var filePath = Path.Combine(savedirectory, filename);
                    File.Copy(templateFile, filePath, true);
                    //set catalog name
                    var catalogRow = GetCatalogInfo(soldto, catalog, savedirectory);
                    CurrencyCode = GetColumnValue(catalogRow, "CurrencyCode");
                    //get ordered columns and ship window
                    GetOrderColumns(catalog, savedirectory, catalogRow);
                    //get level1 and level2 category list
                    GetTabularCategoryData(soldto, catalog, savedirectory);
                    //save category list to Summary sheet, //----and also currency
                    SaveCategoryDataToSummarySheet(catalogRow, filename, savedirectory);
                    //generate top category sheets
                    SaveCategoryProductsToSheet(soldto, catalog, filePath, savedirectory);
                    //save category list
                    SaveCategoryListToSheet(filePath);
                    //validation
                    SaveCategoryDateValidationToSheet(catalogRow, filePath);
                    //save customer data to sheet
                    if (CustomerDataEnabled == "1")
                    {
                        var customerData = GetTabularCustomerData(soldto, catalog, savedirectory);
                        if (customerData != null && customerData.Rows.Count > 0)
                        {
                            SaveTabularCustomerDataToSheet(customerData, Path.Combine(savedirectory, filename));
                        }
                    }
                    //protect
                    ProtectWorkbook(filename, savedirectory);
                }
                catch (Exception ex)
                {
                    WriteToLog(ex.Message);
                    WriteToLog(ex.StackTrace);
                }
            }
            else
                WriteToLog("Cannot find template file " + templateFile);
            WriteToLog(DateTime.Now.ToString() + " End " + filename);
        }

        #region GetCatalogInfo
        private DataRow GetCatalogInfo(string soldto, string catalog, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_TabularCatalogInfo", conn);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
                cmd.Parameters["@catcd"].Value = catalog;

                cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
                cmd.Parameters["@SoldTo"].Value = soldto;

                try
                {
                    DataSet outds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(outds);
                    conn.Close();
                    if (outds.Tables.Count > 0 && outds.Tables[0].Rows.Count > 0)
                    {
                        return outds.Tables[0].Rows[0];
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    //throw;
                    WriteToLog("p_Offline_TabularCatalogInfo |" + catalog, ex, savedirectory);
                    return null;
                }
            }
        }
        #endregion

        #region GetTabularCategoryData [p_Offline_CategoryTabularCategoryList]
        private void GetTabularCategoryData(string soldto, string catalog, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_CategoryTabularCategoryList", conn);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
                cmd.Parameters["@catcd"].Value = catalog;

                cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
                cmd.Parameters["@SoldTo"].Value = soldto;

                try
                {
                    DataSet outds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(outds);
                    conn.Close();
                    CategoryLevelList = new List<CategoryLevelInfo>();
                    if (outds.Tables.Count > 0 && outds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow row in outds.Tables[0].Rows)
                        {
                            CategoryLevelInfo category = new CategoryLevelInfo
                            {
                                CategoryID = row["CategoryID"].ToString(),
                                Category = row["Category"].ToString(),
                                ParentID = row.IsNull("ParentID") ? string.Empty : row["ParentID"].ToString(),
                                ParentCategory = row.IsNull("ParentCategory") ? string.Empty : row["ParentCategory"].ToString(),
                                IsTopLevel = row["TopLevel"].ToString() == "1" ? true : false
                            };
                            category.Category = category.Category.Replace("\\", "").Replace("/", "").Replace("?", "").Replace("*", "").Replace("[", "").Replace("]", "");
                            category.Category = category.Category.Length > 22 ? category.Category.Substring(0, 22) : category.Category;
                            category.ParentCategory = category.ParentCategory.Replace("\\", "").Replace("/", "").Replace("?", "").Replace("*", "").Replace("[", "").Replace("]", "");
                            category.ParentCategory = category.ParentCategory.Length > 22 ? category.ParentCategory.Substring(0, 22) : category.ParentCategory;
                            CategoryLevelList.Add(category);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //throw;
                    WriteToLog("p_Offline_CategoryTabularCategoryList | " + catalog + " | " + soldto, ex, savedirectory);
                }
            }
        }
        #endregion

        private void SaveCategoryDataToSummarySheet(DataRow catalogRow, string filename, string savedirectory)
        {
            var filePath = Path.Combine(savedirectory, filename);
            uint index = 0;
            List<string> summaryColList = SummaryCategoryColumns.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            ExcelWorksheet sheet = null;
            using (ExcelPackage book = new ExcelPackage(new FileInfo(filePath)))
            {
                sheet = book.Workbook.Worksheets[SummarySheetName];
                if (sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + SummarySheetName + ", file path:" + filePath);

                if (sheet != null && CategoryLevelList.Count > 0)
                {
                    //var styleSheet = book.WorkbookPart.WorkbookStylesPart.Stylesheet;
                    //CommonFontId = CreateFontFormat(styleSheet);
                    //CommonBorderAllId = CreateBorderFormat(styleSheet, true, true, true, true);
                    //CommonTextFormat = CreateCellFormat(styleSheet, CommonFontId, null, null, UInt32Value.FromUInt32(49));
                    //CommonBoldTextId = CreateCellFormat(styleSheet, bolderfont, null, CommonBorderAllId, UInt32Value.FromUInt32(49));
                    //var textNullBoldStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid1, null, UInt32Value.FromUInt32(49));
                    //var textNullRightBoldStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid2, null, UInt32Value.FromUInt32(49), true, false, HorizontalAlignmentValues.Right);
                    //var doubleAllMainStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid1, CommonBorderAllId, numformatid);
                    //var doubleAllSubStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid2, CommonBorderAllId, numformatid);
                    //var intAllMainStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid1, CommonBorderAllId, intformatid);
                    //var intAllSubStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid2, CommonBorderAllId, intformatid);
                    //var doublePinkAllSubStyleId = CreateCellFormat(styleSheet, CommonFontId, pinkfileid, CommonBorderAllId, numformatid);
                    //var intPinkAllMainStyleId = CreateCellFormat(styleSheet, CommonFontId, pinkfileid, CommonBorderAllId, intformatid);

                    //save catalog code to A6, save catalog name to B6
                    var Locations = CatalogCodeLocation.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    SetCellValue(sheet, Locations[0] + Locations[1], GetColumnValue(catalogRow, "CatalogCode"));
                    Locations = CatalogNameLocation.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    SetCellValue(sheet, Locations[0] + Locations[1], GetColumnValue(catalogRow, "CatalogName"));

                    BookingOrder = GetColumnValue(catalogRow, "BookingCatalog") == "2" ? true : false;
                    if (BookingOrder)
                    {
                        //program code
                        Locations = SummaryProgramCodeLocation.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        SetCellValue(sheet, Locations[0] + Locations[1], ProgramCode);

                        //Payment Terms
                        var terms = CodeMethods("BookingPaymentTerms", savedirectory);
                        if (terms != null && terms.Rows.Count > 0)
                        {
                            var defaultPaymentTerms = terms.Rows[0]["code"].ToString();
                            Locations = SummaryPaymentTermsLocation.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                            var cell = SetCellValue(sheet, Locations[0] + Locations[1], defaultPaymentTerms);
                            cell.Style.Locked = false;
                            
                            var validationCell = sheet.DataValidations.AddListValidation(Locations[0] + Locations[1]);
                            foreach (DataRow row in terms.Rows)
                            {
                                validationCell.Formula.Values.Add(row["code"].ToString());
                            }
                            validationCell.ShowErrorMessage = true;
                            validationCell.Error = "Please select from dropdown list";
                        }
                    }

                    //save category level list
                    var SUMUnitsFormat = string.Empty;
                    var SUMValueFormat = string.Empty;
                    var categoryNameColumn = summaryColList[0];
                    var lastOrderColumn = OrderColumnList[OrderColumnList.Count - 1].SummaryColumn;
                    var topCategoryTotalColumn = ((char)(char.Parse(lastOrderColumn) + 1)).ToString();
                    var currencyColumn = ((char)(char.Parse(topCategoryTotalColumn) + 1)).ToString();

                    //total currency
                    if (CategorySubTotalForStyle == "1")
                        SetCellValue(sheet, SummaryCurrencyColumn + (SummaryCategoryStartRow - 1).ToString(), CurrencyCode);
                    else
                        SetCellValue(sheet, currencyColumn + (SummaryCategoryStartRow - 1).ToString(), CurrencyCode);

                    var tops = from c in CategoryLevelList where string.IsNullOrEmpty(c.ParentID) select c;
                    foreach (CategoryLevelInfo category in tops)
                    {
                        if (CategorySubTotalForStyle == "1")
                        {
                            #region FOR BAUER
                            //top category units
                            SetCellValue(sheet, "A" + (index + SummaryCategoryStartRow).ToString(), "U0");
                            SetCellValue(sheet, "B" + (index + SummaryCategoryStartRow).ToString(), category.CategoryID);
                            var cell = SetCellValue(sheet, summaryColList[0] + (index + SummaryCategoryStartRow).ToString(), category.Category + " Total");//textNullBoldStyleId
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                            SUMUnitsFormat = string.IsNullOrEmpty(SUMUnitsFormat) ? "COLUMN" + (index + SummaryCategoryStartRow).ToString() : SUMUnitsFormat + ", COLUMN" + (index + SummaryCategoryStartRow).ToString();
                            var level2s = from c in CategoryLevelList where c.ParentID == category.CategoryID select c;
                            StringBuilder unitSum = new StringBuilder(), subtotalSum = new StringBuilder();
                            for (int i = 1; i < summaryColList.Count; i++)
                            {
                                var tformula = string.Empty;
                                if (i == summaryColList.Count - 1)
                                {
                                    //SUM the units of this row
                                    tformula = "SUM(" + subtotalSum + ")";
                                }
                                else if (i == summaryColList.Count - 2)
                                {
                                    //SUM the units of this row
                                    tformula = "SUM(" + unitSum + ")";
                                }
                                else
                                {
                                    if (i % 2 == 1)
                                        unitSum.AppendFormat("{0}{1}{2}", (unitSum.Length > 0 ? "," : string.Empty), summaryColList[i], index + SummaryCategoryStartRow);
                                    else
                                        subtotalSum.AppendFormat("{0}{1}{2}", (subtotalSum.Length > 0 ? "," : string.Empty), summaryColList[i], index + SummaryCategoryStartRow);
                                    tformula = "SUM(" + summaryColList[i] + (index + 1 + SummaryCategoryStartRow).ToString() + ":" + summaryColList[i] + (index + level2s.Count() + SummaryCategoryStartRow).ToString() + ")";
                                }
                                var dataType = (i % 2 == 1 ? "1" : "2");
                                cell = SetCellValue(sheet, summaryColList[i] + (index + SummaryCategoryStartRow).ToString(), string.Empty, dataType);
                                cell.Formula = tformula;
                                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                //SetCellValue(cell, string.Empty, 3, (i % 2 == 1 ? intAllSubStyleId : doubleAllSubStyleId));
                            }
                            index++;
                            //level2 category units
                            foreach (CategoryLevelInfo sub in level2s)
                            {
                                unitSum = new StringBuilder(); subtotalSum = new StringBuilder();
                                SetCellValue(sheet, "A" + (index + SummaryCategoryStartRow).ToString(), "U");
                                SetCellValue(sheet, "B" + (index + SummaryCategoryStartRow).ToString(), sub.CategoryID);
                                //SetCellValue(cell, , 1, textNullRightBoldStyleId);
                                cell = SetCellValue(sheet, summaryColList[0] + (index + SummaryCategoryStartRow).ToString(), sub.Category);
                                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.WhiteSmoke);
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                for (int i = 1; i < summaryColList.Count; i++)
                                {
                                    var tformula = string.Empty;
                                    if (i == summaryColList.Count - 1)
                                    {
                                        //SUM the units of this row
                                        tformula = "SUM(" + subtotalSum + ")";
                                    }
                                    else if (i == summaryColList.Count - 2)
                                    {
                                        //SUM the units of this row
                                        tformula = "SUM(" + unitSum + ")";
                                    }
                                    else
                                    {
                                        if (i % 2 == 1)
                                            unitSum.AppendFormat("{0}{1}{2}", (unitSum.Length > 0 ? "," : string.Empty), summaryColList[i], index + SummaryCategoryStartRow);
                                        else
                                            subtotalSum.AppendFormat("{0}{1}{2}", (subtotalSum.Length > 0 ? "," : string.Empty), summaryColList[i], index + SummaryCategoryStartRow);
                                    }
                                    //SetCellValue(cell, string.Empty, 3, (i % 2 == 1 ? intAllSubStyleId : doubleAllSubStyleId));
                                    var dataType = (i % 2 == 1 ? "1" : "2");
                                    cell = SetCellValue(sheet, summaryColList[i] + (index + SummaryCategoryStartRow).ToString(), string.Empty, dataType);
                                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.WhiteSmoke);
                                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    if (!string.IsNullOrEmpty(tformula))
                                        cell.Formula = tformula;
                                }
                                index++;
                            }
                            #endregion
                        }
                        else
                        {
                            var level2s = from c in CategoryLevelList where c.ParentID == category.CategoryID select c;

                            #region top category units
                            if(level2s.Count() > 0)
                                SetCellValue(sheet, "A" + (index + SummaryCategoryStartRow).ToString(), "U0");
                            else
                                SetCellValue(sheet, "A" + (index + SummaryCategoryStartRow).ToString(), "US");
                            SetCellValue(sheet, "B" + (index + SummaryCategoryStartRow).ToString(), category.CategoryID);
                            var cell = SetCellValue(sheet, categoryNameColumn + (index + SummaryCategoryStartRow).ToString(), category.Category + " Total Units");//textNullBoldStyleId
                            SetCellStyle(cell, System.Drawing.Color.LightGray, false, false, false, false);

                            SUMUnitsFormat = string.IsNullOrEmpty(SUMUnitsFormat) ? "COLUMN" + (index + SummaryCategoryStartRow).ToString() : SUMUnitsFormat + ", COLUMN" + (index + SummaryCategoryStartRow).ToString();

                            //total unites for top level category
                            var tformula = string.Empty;
                            for (int i = 0; i < OrderColumnList.Count; i++)
                            {
                                cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + (index + SummaryCategoryStartRow).ToString(), string.Empty, "1");
                                if (level2s.Count() > 0)
                                {
                                    tformula = "SUM(" + OrderColumnList[i].SummaryColumn + (index + 1 + SummaryCategoryStartRow).ToString() + ":" + OrderColumnList[i].SummaryColumn + (index + level2s.Count() + SummaryCategoryStartRow).ToString() + ")";
                                    cell.Formula = tformula;
                                }
                                SetCellStyle(cell, System.Drawing.Color.LightGray);
                            }
                            //total unites of all order columns for top level category
                            tformula = "SUM(" + OrderColumnList[0].SummaryColumn + (index + SummaryCategoryStartRow).ToString() + ":" + lastOrderColumn + (index + SummaryCategoryStartRow).ToString() + ")";
                            cell = SetCellValue(sheet, topCategoryTotalColumn + (index + SummaryCategoryStartRow).ToString(), string.Empty, "1");
                            cell.Formula = tformula;
                            SetCellStyle(cell, System.Drawing.Color.LightGray);

                            index++;
                            #endregion

                            #region level2 category units
                            foreach (CategoryLevelInfo sub in level2s)
                            {
                                SetCellValue(sheet, "A" + (index + SummaryCategoryStartRow).ToString(), "U");
                                SetCellValue(sheet, "B" + (index + SummaryCategoryStartRow).ToString(), sub.CategoryID);
                                //SetCellValue(cell, , 1, textNullRightBoldStyleId);
                                cell = SetCellValue(sheet, categoryNameColumn + (index + SummaryCategoryStartRow).ToString(), sub.Category);
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                SetCellStyle(cell, System.Drawing.Color.WhiteSmoke, false, false, false, false);

                                for (int i = 0; i < OrderColumnList.Count; i++)
                                {
                                    cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + (index + SummaryCategoryStartRow).ToString(), string.Empty, "1");
                                    SetCellStyle(cell, System.Drawing.Color.WhiteSmoke);
                                }
                                tformula = "SUM(" + OrderColumnList[0].SummaryColumn + (index + SummaryCategoryStartRow).ToString() + ":" + lastOrderColumn + (index + SummaryCategoryStartRow).ToString() + ")"; ;
                                cell = SetCellValue(sheet, topCategoryTotalColumn + (index + SummaryCategoryStartRow).ToString(), string.Empty, "1");
                                SetCellStyle(cell, System.Drawing.Color.WhiteSmoke);
                                cell.Formula = tformula;

                                index++;
                            }
                            #endregion

                            #region top category amount value
                            if (level2s.Count() > 0)
                                SetCellValue(sheet, "A" + (index + SummaryCategoryStartRow).ToString(), "V0");
                            else
                                SetCellValue(sheet, "A" + (index + SummaryCategoryStartRow).ToString(), "VS");
                            SetCellValue(sheet, "B" + (index + SummaryCategoryStartRow).ToString(), category.CategoryID);
                            //SetCellValue(cell, category.Category + " Total Value", 1, textNullBoldStyleId);
                            cell = SetCellValue(sheet, categoryNameColumn + (index + SummaryCategoryStartRow).ToString(), category.Category + " Total Value");
                            SetCellStyle(cell, System.Drawing.Color.LightGray, false, false, false, false);

                            SUMValueFormat = string.IsNullOrEmpty(SUMValueFormat) ? "COLUMN" + (index + SummaryCategoryStartRow).ToString() : SUMValueFormat + ", COLUMN" + (index + SummaryCategoryStartRow).ToString();

                            //total amounts for top level category
                            tformula = string.Empty;
                            for (int i = 0; i < OrderColumnList.Count; i++)
                            {
                                cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + (index + SummaryCategoryStartRow).ToString(), string.Empty, "2");
                                if (level2s.Count() > 0)
                                {
                                    tformula = "SUM(" + OrderColumnList[i].SummaryColumn + (index + 1 + SummaryCategoryStartRow).ToString() + ":" + OrderColumnList[i].SummaryColumn + (index + level2s.Count() + SummaryCategoryStartRow).ToString() + ")";
                                    cell.Formula = tformula;
                                }
                                SetCellStyle(cell, System.Drawing.Color.LightGray);
                            }
                            //total amounts of all order columns for top level category
                            tformula = "SUM(" + OrderColumnList[0].SummaryColumn + (index + SummaryCategoryStartRow).ToString() + ":" + lastOrderColumn + (index + SummaryCategoryStartRow).ToString() + ")";
                            cell = SetCellValue(sheet, topCategoryTotalColumn + (index + SummaryCategoryStartRow).ToString(), string.Empty, "2");
                            cell.Formula = tformula;
                            SetCellStyle(cell, System.Drawing.Color.LightGray);

                            //set currency to cell SetCellValue(cell, CurrencyCode, 1, CommonTextFormat);
                            SetCellValue(sheet, currencyColumn + (index + SummaryCategoryStartRow).ToString(), CurrencyCode);
                            index++;
                            #endregion

                            #region level2 category amount value
                            foreach (CategoryLevelInfo sub in level2s)
                            {
                                SetCellValue(sheet, "A" + (index + SummaryCategoryStartRow).ToString(), "V");
                                SetCellValue(sheet, "B" + (index + SummaryCategoryStartRow).ToString(), sub.CategoryID);
                                //SetCellValue(cell, , 1, textNullRightBoldStyleId);
                                cell = SetCellValue(sheet, categoryNameColumn + (index + SummaryCategoryStartRow).ToString(), sub.Category);
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                SetCellStyle(cell, System.Drawing.Color.WhiteSmoke, false, false, false, false);

                                for (int i = 0; i < OrderColumnList.Count; i++)
                                {
                                    cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + (index + SummaryCategoryStartRow).ToString(), string.Empty, "2");
                                    SetCellStyle(cell, System.Drawing.Color.WhiteSmoke);
                                }
                                tformula = "SUM(" + OrderColumnList[0].SummaryColumn + (index + SummaryCategoryStartRow).ToString() + ":" + lastOrderColumn + (index + SummaryCategoryStartRow).ToString() + ")"; ;
                                cell = SetCellValue(sheet, topCategoryTotalColumn + (index + SummaryCategoryStartRow).ToString(), string.Empty, "2");
                                SetCellStyle(cell, System.Drawing.Color.WhiteSmoke);
                                cell.Formula = tformula;

                                SetCellValue(sheet, currencyColumn + (index + SummaryCategoryStartRow).ToString(), CurrencyCode);
                                index++;
                            }
                            #endregion
                        }
                    }
                    if (CategorySubTotalForStyle == "1")
                    {
                        #region FOR BAUER
                        //SetCellValue(cell, , 1, textNullBoldStyleId);
                        var cell = SetCellValue(sheet, summaryColList[0] + (SummaryOrderTotalUnitsRow + index + 2).ToString(), "Order Total");
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        #endregion
                    }

                    if (CategorySubTotalForStyle == "1")
                    {
                        for (int i = 1; i < summaryColList.Count; i++)
                        {
                            #region FOR BAUER
                            var tformula = string.Empty;
                            tformula = "SUM(" + SUMUnitsFormat.Replace("COLUMN", summaryColList[i]) + ")";
                            //SetCellValue(cell, string.Empty, 3, (i % 2 == 1 ? intPinkAllMainStyleId : doublePinkAllSubStyleId));
                            var dataType = (i % 2 == 1 ? "1" : "2");
                            var cell = SetCellValue(sheet, summaryColList[i] + (SummaryOrderTotalUnitsRow + index + 2).ToString(), string.Empty, dataType);
                            cell.Formula = tformula;
                            //cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 221, 220));
                            //cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            #endregion
                        }
                    }
                    else
                    {
                        var shipToHeaderRow = SummaryEnableShipTo == "1" ? (int.Parse(SummaryShipDateRow) - 1).ToString() : string.Empty;
                        var shipHeaderRow = string.IsNullOrEmpty(shipToHeaderRow) ? (int.Parse(SummaryShipDateRow) - 1).ToString() : (int.Parse(shipToHeaderRow) - 1).ToString();
                        var ponumberRow = (int.Parse(SummaryShipDateRow) + 1).ToString();
                        var orderHeaderRow = (SummaryOrderTotalUnitsRow - 1).ToString();
                        for (int i = 0; i < OrderColumnList.Count; i++)
                        {
                            #region ship info header
                            var cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + shipHeaderRow, "Order " + (i + 1).ToString());
                            SetCellStyle(cell, System.Drawing.Color.Black);
                            cell.Style.Font.Color.SetColor(System.Drawing.Color.White);

                            if (!string.IsNullOrEmpty(shipToHeaderRow))
                            {
                                cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + shipToHeaderRow, string.Empty);
                                cell.Style.Locked = false;
                                SetCellStyle(cell);
                            }

                            cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + SummaryShipDateRow, string.IsNullOrEmpty(OrderColumnList[i].DateShip) ? string.Empty : OrderColumnList[i].DateShip, "3");
                            cell.Style.Locked = string.IsNullOrEmpty(OrderColumnList[i].DateShip) ? false : true;
                            SetCellStyle(cell);

                            var ponumberAddress = OrderColumnList[i].SummaryColumn + ponumberRow;
                            cell = SetCellValue(sheet, ponumberAddress, string.Empty);
                            cell.Style.Locked = false;
                            SetCellStyle(cell);
                            var SummaryPONumberValidationFormula = ConfigurationManager.AppSettings["SummaryPONumberValidationFormula"] ?? string.Empty;
                            //=ISNUMBER(SUMPRODUCT(FIND(MID(CELLADDRESS,ROW(INDIRECT("1:"&LEN(CELLADDRESS))),1),"0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-"))) for CCM
                            if (!string.IsNullOrEmpty(SummaryPONumberValidationFormula))
                            {
                                var validationCell = sheet.DataValidations.AddCustomValidation(ponumberAddress);
                                SummaryPONumberValidationFormula = SummaryPONumberValidationFormula.Replace("CELLADDRESS", ponumberAddress);
                                validationCell.Formula.ExcelFormula = SummaryPONumberValidationFormula;
                                validationCell.ShowErrorMessage = true;
                                validationCell.Error = ConfigurationManager.AppSettings["SummaryPONumberValidationErrorMessage"] ?? "Your PO Number is invalid.";
                                WriteToLog(SummaryPONumberValidationFormula);
                            }

                            if (SummaryEnableComments == "1")
                            {
                                cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + (int.Parse(ponumberRow) + 1).ToString(), string.Empty);
                                cell.Style.Locked = false;
                                SetCellStyle(cell);
                            }
                            #endregion

                            #region all total units for order 1,2,3,4... on the top
                            cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + orderHeaderRow, "Order " + (i + 1).ToString());
                            SetCellStyle(cell, System.Drawing.Color.Black);
                            cell.Style.Font.Color.SetColor(System.Drawing.Color.White);

                            var tformula = string.Empty;
                            tformula = "SUM(" + SUMUnitsFormat.Replace("COLUMN", OrderColumnList[i].SummaryColumn) + ")";
                            //SetCellValue(cell, string.Empty, 3, intPinkAllMainStyleId);
                            cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + (SummaryOrderTotalUnitsRow).ToString(), string.Empty, "1");
                            cell.Formula = tformula;
                            SetCellStyle(cell, System.Drawing.Color.LightPink);//System.Drawing.Color.FromArgb(242, 221, 220)

                            //all total value
                            tformula = "SUM(" + SUMValueFormat.Replace("COLUMN", OrderColumnList[i].SummaryColumn) + ")";
                            //SetCellValue(cell, string.Empty, 3, doublePinkAllSubStyleId);
                            cell = SetCellValue(sheet, OrderColumnList[i].SummaryColumn + (SummaryOrderTotalValueRow).ToString(), string.Empty, "2");
                            cell.Formula = tformula;
                            SetCellStyle(cell, System.Drawing.Color.LightPink);//System.Drawing.Color.FromArgb(242, 221, 220)
                            #endregion
                        }

                        #region TOTALS after order1-6
                        var tcell = SetCellValue(sheet, topCategoryTotalColumn + orderHeaderRow, "TOTALS");
                        SetCellStyle(tcell, System.Drawing.Color.Black);
                        tcell.Style.Font.Color.SetColor(System.Drawing.Color.White);

                        var formula = "SUM(" + OrderColumnList[0].SummaryColumn + (SummaryOrderTotalUnitsRow).ToString() + ":" + OrderColumnList[OrderColumnList.Count - 1].SummaryColumn + (SummaryOrderTotalUnitsRow).ToString() + ")";
                        tcell = SetCellValue(sheet, topCategoryTotalColumn + (SummaryOrderTotalUnitsRow).ToString(), string.Empty, "1");
                        SetCellStyle(tcell, System.Drawing.Color.LightPink);
                        tcell.Formula = formula;

                        formula = "SUM(" + OrderColumnList[0].SummaryColumn + (SummaryOrderTotalValueRow).ToString() + ":" + OrderColumnList[OrderColumnList.Count - 1].SummaryColumn + (SummaryOrderTotalValueRow).ToString() + ")";
                        tcell = SetCellValue(sheet, topCategoryTotalColumn + (SummaryOrderTotalValueRow).ToString(), string.Empty, "2");
                        SetCellStyle(tcell, System.Drawing.Color.LightPink);
                        tcell.Formula = formula;
                        #endregion
                    }
                    //sheet.Save();
                }
                book.Save();
            }
        }

        #region SAVE MAIN DATA TO SKU SHEETS
        private void SaveCategoryProductsToSheet(string soldto, string catalog, string filePath, string savedirectory)
        {
            var tops = from c in CategoryLevelList where string.IsNullOrEmpty(c.ParentID) select c;
            int i = 1;
            DataTable SKUTable = null;
            foreach (CategoryLevelInfo category in tops)
            {
                InsertWorksheet(category.Category, filePath, i);
                if (CategorySKUATPSheet == "1")
                    InsertWorksheet(category.Category + "_ATP", filePath, i);
                DataTable dt = SaveCategorySKUsToSheet(soldto, catalog, category, filePath, savedirectory, i);

                if (TabularCatalogUPCSheetName.Length > 0)
                {
                    if (SKUTable == null)
                    {
                        SKUTable = dt;
                    }
                    else if (dt != null)
                    {
                        SKUTable.Merge(dt);
                    }
                }
                i++;
            }
            
            if (TabularCatalogUPCSheetName.Length > 0 && SKUTable != null && SKUTable.Rows.Count > 0)
            {
                DataTable SortedSKUTable = DataTableSort(SKUTable, ConfigurationManager.AppSettings["TabularCatalogUPCSort"] ?? "SheetNameSortOrder ASC,RowIdx ASC");
                uint TabularCatalogUPCStartRowIdx = uint.Parse(ConfigurationManager.AppSettings["TabularCatalogUPCStartRowIdx"] ?? "9");
                Dictionary<string, string> columnList = (ConfigurationManager.AppSettings["TabularCatalogUPCColumnList"] ?? "A|SKU,B|UPC,C|SheetName,D|Style,E|ProductName,F|AttributeValue2,G|PriceWholesale").ToDic(new char[] { '|', ',' });
                Dictionary<string, string> formulacolumnList = (ConfigurationManager.AppSettings["TabularCatalogUPCFormulaColumns"] ?? "I|B{0},J|'{1}'!L{2},L|B{0},M|'{1}'!M{2},O|B{0},P|'{1}'!N{2},R|B{0},S|'{1}'!O{2},U|B{0},V|'{1}'!P{2},X|B{0},Y|'{1}'!Q{2},AA|B{0},AB|J{0}+M{0}+P{0}+S{0}+V{0}+Y{0}").ToDic(new char[] { '|', ',' });

                using (ExcelPackage spreadSheet = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet sheet = null;
                    sheet = spreadSheet.Workbook.Worksheets[TabularCatalogUPCSheetName];
                    if (sheet == null)
                        WriteToLog("Wrong template order form file, cannot find sheet " + TabularCatalogUPCSheetName + ", file path:" + filePath);

                    if (sheet != null)
                    {
                        //var styleSheet = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
                        //uint sheetBorder = CreateBorderFormat(styleSheet, false, false, false, true);
                        //var doubleAllMainStyleId = CreateCellFormat(styleSheet, null, null, null, numformatid);

                        uint k = 1;
                        foreach (DataRow dr in SortedSKUTable.Rows)
                        {
                            uint rowIdx = (uint)dr["RowIdx"];
                            uint categoryIdx = (uint)dr["SheetNameSortOrder"];
                            
                            if (rowIdx > 0)
                            {
                                foreach (KeyValuePair<string, string> kvp in columnList)
                                {
                                    string cvalue = string.Empty; 
                                    bool dateformat = false;
                                    if (kvp.Value == "RowSortOrder")
                                    {
                                        cvalue = (categoryIdx * 100000 + rowIdx).ToString();
                                    }
                                    else if (kvp.Value == "LaunchDate")
                                    {
                                        dateformat = true;
                                        cvalue = dr[kvp.Value].ToString();
                                    }
                                    else if (kvp.Value.Contains("+"))
                                    {
                                        string[] arr = kvp.Value.Split(new char[] { '+' });
                                        foreach (string s in arr)
                                            cvalue = cvalue + dr[s].ToString() + " ";
                                    }
                                    else
                                    {
                                        cvalue = dr[kvp.Value].ToString();
                                    }
                                    //SetCellValue(cCell, cvalue, (new string[] { "RowSortOrder", "PriceWholesale", "RowNo" }).Contains(kvp.Value) ? 2 : 1, doubleAllMainStyleId);
                                    ExcelRange cell = null;
                                    if (dateformat)
                                    {
                                        DateTime dt;
                                        if (DateTime.TryParse(cvalue, out dt))
                                        {
                                            var tformula = string.Format("DATE({0},{1},{2})", dt.Year, dt.Month, dt.Day);
                                            cell = SetCellValue(sheet, kvp.Key + (k + TabularCatalogUPCStartRowIdx).ToString(), string.Empty, "3");
                                            cell.Formula = tformula;
                                            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                        }
                                    }
                                    else
                                    {
                                        if ((new string[] { "RowSortOrder", "PriceWholesale", "RowNo" }).Contains(kvp.Value))
                                            cell = SetCellValue(sheet, kvp.Key + (k + TabularCatalogUPCStartRowIdx).ToString(), cvalue, "2");
                                        else
                                            cell = SetCellValue(sheet, kvp.Key + (k + TabularCatalogUPCStartRowIdx).ToString(), cvalue);
                                    }
                                }
                                foreach (KeyValuePair<string, string> kvp in formulacolumnList)
                                {
                                    var tformula = string.Format(kvp.Value, (k + TabularCatalogUPCStartRowIdx).ToString(), dr["SheetName"].ToString(), dr["RowIdx"].ToString());
                                    var cell = GetCell(sheet, kvp.Key + (k + TabularCatalogUPCStartRowIdx).ToString());
                                    cell.Formula = tformula;
                                }
                                k++;
                            }
                        }
                    }

                    spreadSheet.Save();
                }
            }
        }

        private static DataTable DataTableSort(DataTable dt, string sort)
        {
            DataView dv = dt.DefaultView;
            dv.Sort = sort;
            DataTable newdt = dv.ToTable();
            return newdt;
        }

        /*private void InsertWorksheet(string sheetName, string filePath, int idx)
        {
            // Open the document for editing.
            using (ExcelPackage spreadSheet = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = null;
                var tempName = string.Format(CategoryTemplateSheetName, idx);
                if (CategoryTemplateSheetName.Contains("{0}"))
                {
                    sheet = spreadSheet.Workbook.Worksheets[tempName];
                    sheet.Name = sheetName;
                    sheet.Hidden = eWorkSheetHidden.Visible;
                }
                else
                {
                    sheet = spreadSheet.Workbook.Worksheets.Copy(tempName, sheetName);
                    sheet.Hidden = eWorkSheetHidden.Visible;
                }
                //change sheet total titles
                var Locations = CategoryTotalUnitsTitleCell.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                var title = GetCellValue(sheet, Locations[0] + Locations[1]);
                if (!string.IsNullOrEmpty(title))
                    title = string.Format(title, sheetName);
                SetCellValue(sheet, Locations[0] + Locations[1], title);

                Locations = CategoryTotalValueTitleCell.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                title = GetCellValue(sheet, Locations[0] + Locations[1]);
                if (!string.IsNullOrEmpty(title))
                    title = string.Format(title, sheetName);
                SetCellValue(sheet, Locations[0] + Locations[1], title);

                spreadSheet.Save();
            }
        }*/

        /// <summary>
        /// EPPLUS CANNOT COPY SHEET WITH FORM CONTROLS, SO WE HAVE TO USE OPENXML METHOD
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="filePath"></param>
        /// <param name="idx"></param>
        private void InsertWorksheet(string sheetName, string filePath, int idx)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
            {
                //get template
                WorksheetPart clonedSheet = null;
                var tempSheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == string.Format(CategoryTemplateSheetName, idx));
                if (CategoryTemplateSheetName.Contains("{0}"))
                {
                    Sheet tmp = tempSheets.FirstOrDefault();
                    if (tmp != null)
                    {
                        clonedSheet = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(tmp.Id) as WorksheetPart;
                        tmp.Name = sheetName;
                        tmp.State = SheetStateValues.Visible;
                    }
                }
                else
                {
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

            using (ExcelPackage spreadSheet = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = spreadSheet.Workbook.Worksheets[sheetName];
                //change sheet total titles
                var Locations = CategoryTotalUnitsTitleCell.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                var title = GetCellValue(sheet, Locations[0] + Locations[1]);
                if (!string.IsNullOrEmpty(title))
                    title = string.Format(title, sheetName);
                SetCellValue(sheet, Locations[0] + Locations[1], title);

                Locations = CategoryTotalValueTitleCell.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                title = GetCellValue(sheet, Locations[0] + Locations[1]);
                if (!string.IsNullOrEmpty(title))
                    title = string.Format(title, sheetName);
                SetCellValue(sheet, Locations[0] + Locations[1], title);

                spreadSheet.Save();
            }
        }

        private DataTable SaveCategorySKUsToSheet(string soldto, string catalog, CategoryLevelInfo category, string filePath, string savedirectory, int categoryIdx)
        {
            var CategorySKUShipDateAboveRows = string.IsNullOrEmpty(ConfigurationManager.AppSettings["CategorySKUShipDateAboveRows"]) ?
                3 : int.Parse(ConfigurationManager.AppSettings["CategorySKUShipDateAboveRows"]);//FOR BAUER, IT'S 4
            var CategorySKUTotalColor = ConfigurationManager.AppSettings["CategorySKUTotalColor"];//"#00A4E4" -- BAUER COLOR;
            var cellTotalColor = System.Drawing.Color.Yellow;
            if(!string.IsNullOrEmpty(CategorySKUTotalColor))
                cellTotalColor = System.Drawing.ColorTranslator.FromHtml(CategorySKUTotalColor);
            //var cellCommonColor = System.Drawing.Color.WhiteSmoke;
            //if(CategorySubTotalForStyle == "1")
            //    cellCommonColor = System.Drawing.Color.LightGray;

            var orderlist = CategorySKUOrderColumnList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            var dt = GetCategoryTabularOrderFormData(soldto, catalog, category.CategoryID, savedirectory);
            dt.Columns.Add("SheetName", typeof(string));
            dt.Columns.Add("SheetNameSortOrder", typeof(uint));
            dt.Columns.Add("RowIdx", typeof(uint));
            int totalSKUs = dt.Rows.Count;
            if (dt != null && totalSKUs > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];
                    dr["SheetName"] = category.Category;
                    dr["SheetNameSortOrder"] = categoryIdx;
                    dr["RowIdx"] = 0;
                }

                int styleSubtotalCount = 0, styleSubtotalIndex = 0;
                if (CategorySubTotalForStyle == "1")
                    styleSubtotalCount = dt.DefaultView.ToTable(true, "Level2DeptID", "Style", "DepartmentID").Rows.Count;

                //category.SKUBeginRow = 0;
                category.SKURowNumber = (uint)totalSKUs;
                using (ExcelPackage spreadSheet = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet sheet = null;
                    sheet = spreadSheet.Workbook.Worksheets[category.Category];
                    if(sheet == null)
                        WriteToLog("Wrong template order form file, cannot find sheet " + category.Category + ", file path:" + filePath);

                    ExcelWorksheet atpSheet = null;
                    if (CategorySKUATPSheet == "1")
                    {
                        atpSheet = spreadSheet.Workbook.Worksheets[category.Category + "_ATP"];
                        if (atpSheet == null)
                            WriteToLog("Wrong template order form file, cannot find ATP sheet " + category.Category + "_ATP" + ", file path:" + filePath);
                        else
                            atpSheet.Hidden = eWorkSheetHidden.Hidden;
                    }

                    if (sheet != null)
                    {
                        //Style Format
                        //var styleSheet = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
                        
                        //var borderLRId = CreateBorderFormat(styleSheet, true, true, false, false);
                        //var borderLRTId = CreateBorderFormat(styleSheet, true, true, true, false);
                        //var borderLRBId = CreateBorderFormat(styleSheet, true, true, false, true);
                        //var doubleStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, numformatid); // UInt32Value.FromUInt32(2));
                        //var doubleUnlockedStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, numformatid, false, false); // UInt32Value.FromUInt32(2));
                        //var doubleTStyleId = CreateCellFormat(styleSheet, CategoryTabTotalFontId.HasValue ? CategoryTabTotalFontId.Value : CommonFontId, CategoryTabTotalFillId, CommonBorderAllId, numformatid); //UInt32Value.FromUInt32(2));
                        //var intStyleId = CreateCellFormat(styleSheet, CategoryTabTotalFontId.HasValue ? CategoryTabTotalFontId.Value : CommonFontId, CategoryTabTotalFillId, CommonBorderAllId, intformatid); //UInt32Value.FromUInt32(1));
                        //var intLStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, intformatid, false, false); //UInt32Value.FromUInt32(1), false, false);
                        //var textAllStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, UInt32Value.FromUInt32(49));
                        //var textLRStyleId = CreateCellFormat(styleSheet, CommonFontId, null, borderLRId, UInt32Value.FromUInt32(49));
                        //var intLockedLStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid2, CommonBorderAllId, intformatid, true, false);
                        //var textLRTStyleId = CreateCellFormat(styleSheet, CommonFontId, null, borderLRTId, UInt32Value.FromUInt32(49));
                        //var textLRBStyleId = CreateCellFormat(styleSheet, CommonFontId, null, borderLRBId, UInt32Value.FromUInt32(49));
                        //var textNullStyleId = CreateCellFormat(styleSheet, CommonFontId, null, null, UInt32Value.FromUInt32(49));
                        //var textDAllStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, UInt32Value.FromUInt32(49), true, true);
                        //var textDLRTStyleId = CreateCellFormat(styleSheet, CommonFontId, null, borderLRTId, UInt32Value.FromUInt32(49), true, true);
                        //var textNullBlackStyleId = CreateCellFormat(styleSheet, CommonFontId, blackFillID, null, UInt32Value.FromUInt32(49));
                        //var textAllSubStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid2, CommonBorderAllId, UInt32Value.FromUInt32(49));
                        //var doubleAllSubStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid2, CommonBorderAllId, numformatid);
                        //var intAllSubStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid2, CommonBorderAllId, intformatid);
                        //var doubleSubTotalStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid1, CommonBorderAllId, numformatid);
                        //var intSubTotalStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid1, CommonBorderAllId, intformatid);
                        //Check previous departmentid, style, color
                        var preCategoryId = string.Empty;
                        var preDept = string.Empty;
                        var preStyle = string.Empty;
                        var preColor = string.Empty;
                        var multipleList = new List<OrderMultipleDataValidation>();

                        ExcelRange cCell = null;

                        #region set currency in category sheet
                        var currencys = CategotySKUCurrencyLocations.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        var location = currencys[0].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        cCell = SetCellValue(sheet, location[0] + location[1], CurrencyCode); //, 1, textAllStyleId);
                        SetCellStyle(cCell);
                        var lastOrderColumn = OrderColumnList[OrderColumnList.Count - 1].SKUColumn;
                        var currencyColumn = ((char)(char.Parse(lastOrderColumn) + 1)).ToString();
                        cCell = SetCellValue(sheet, currencyColumn + location[1], CurrencyCode); //, 1, textAllStyleId);
                        SetCellStyle(cCell);
                        #endregion

                        var styleTotalUnitsColumn = currencyColumn;
                        var styleTotalAmountColumn = ((char)(char.Parse(currencyColumn) + 1)).ToString();
                        
                        #region total units and amount for order 1-6, CELL G8, G9
                        var totalunits = CategorySKUTotalUnitsLocation.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        var tformula = "SUM(" + styleTotalUnitsColumn + startRow.ToString() + ":" + styleTotalUnitsColumn + (totalSKUs + startRow + (uint)styleSubtotalCount - 1).ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                        //SetCellValue(cCell, , 3, intStyleId);
                        cCell = SetCellValue(sheet, totalunits[0] + totalunits[1], string.Empty, "1");
                        cCell.Formula = tformula;
                        SetCellStyle(cCell, cellTotalColor);
                        if (CategorySubTotalForStyle == "1") //FOR BAUER
                            cCell.Style.Font.Color.SetColor(System.Drawing.Color.White);

                        //var totalAmountRow = (int.Parse(totalunits[1]) + 1).ToString();
                        tformula = "SUM(" + OrderColumnList[0].SKUColumn + CategorySKUSubTotalRow + ":" + OrderColumnList[OrderColumnList.Count - 1].SKUColumn + CategorySKUSubTotalRow + ")";
                        cCell = SetCellValue(sheet, totalunits[0] + CategorySKUSubTotalRow, string.Empty, "2");
                        cCell.Formula = tformula;
                        SetCellStyle(cCell, cellTotalColor);
                        if (CategorySubTotalForStyle == "1") //FOR BAUER
                            cCell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        #endregion
                        
                        #region total header, total amount value for order1-6 CELL L9-Q9, 6 total cells
                        var shipHeaderRow = (CategorySKUSubTotalRow - CategorySKUShipDateAboveRows).ToString();
                        var shipDateRow = (CategorySKUSubTotalRow - 2).ToString();
                        var skuHeaderRow = (startRow - 1).ToString();
                        for (int o = 0; o < OrderColumnList.Count; o++)
                        {
                            //ship date header
                            cCell = SetCellValue(sheet, OrderColumnList[o].SKUColumn + shipHeaderRow, "Order " + (o + 1).ToString());
                            SetCellStyle(cCell);
                            cCell.Style.Font.Bold = true;
                            if (SummaryEnableShipTo == "1" && int.Parse(shipHeaderRow) < (int.Parse(shipDateRow) -1))
                            {
                                cCell = SetCellValue(sheet, OrderColumnList[o].SKUColumn + (int.Parse(shipHeaderRow) + 1).ToString(), string.Empty);
                                SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                            }
                            //ship date
                            cCell = SetCellValue(sheet, OrderColumnList[o].SKUColumn + shipDateRow, string.IsNullOrEmpty(OrderColumnList[o].DateShip) ? string.Empty : OrderColumnList[o].DateShip, "3");
                            SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                            //ponumber
                            cCell = SetCellValue(sheet, OrderColumnList[o].SKUColumn + totalunits[1], string.Empty);
                            SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                            //sku header
                            cCell = SetCellValue(sheet, OrderColumnList[o].SKUColumn + skuHeaderRow, "Order " + (o + 1).ToString());
                            SetCellStyle(cCell, System.Drawing.Color.Black);
                            cCell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                            cCell.Style.Font.Bold = true;
                            //total amount
                            var endrow = totalSKUs + startRow + styleSubtotalCount - 1;
                            if (totalSKUs <= 1)
                                endrow++;
                            tformula = "SUMPRODUCT(" + CategorySKUPriceColumn + startRow.ToString() + ":" + CategorySKUPriceColumn + (endrow).ToString() + "," + OrderColumnList[o].SKUColumn + startRow.ToString() + ":" + OrderColumnList[o].SKUColumn + (endrow).ToString() + ")";
                            //SetCellValue(cCell, string.Empty, 3, doubleTStyleId);
                            cCell = SetCellValue(sheet, OrderColumnList[o].SKUColumn + CategorySKUSubTotalRow.ToString(), string.Empty, "2");
                            cCell.Formula = tformula;
                            SetCellStyle(cCell, cellTotalColor);
                            if (CategorySubTotalForStyle == "1") //FOR BAUER
                                cCell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        }
                        //Style Total Units header
                        cCell = SetCellValue(sheet, styleTotalUnitsColumn + skuHeaderRow, "Style Total Units");
                        SetCellStyle(cCell, System.Drawing.Color.Black);
                        cCell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        cCell.Style.Font.Bold = true;
                        cCell.AutoFitColumns(14.5);
                        //Style Total Value header
                        cCell = SetCellValue(sheet, styleTotalAmountColumn + skuHeaderRow, "Style Total Value");
                        SetCellStyle(cCell, System.Drawing.Color.Black);
                        cCell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        cCell.Style.Font.Bold = true;
                        cCell.AutoFitColumns(14.5);
                        #endregion

                        if (!CategorySKUComColumnList.Contains(CategorySKUPriceColumn + "|PriceWholesale"))
                            CategorySKUComColumnList = CategorySKUComColumnList + "," + CategorySKUPriceColumn + "|PriceWholesale";
                        var comColumnList = CategorySKUComColumnList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        
                        uint currentStyleStartRows = startRow;
                        for (int i = 0; i < totalSKUs + (CategorySubTotalForStyle == "1" ? 1 : 0); i++) // +1 to Add the Last style-level subtotal line
                        {
                            ExcelRange cell = null;
                            var row = (i == totalSKUs ? null : dt.Rows[i]);
                            var categoryId = GetColumnValue(row, "Level2DeptID");
                            var deptId = GetColumnValue(row, "DepartmentID");
                            var style = GetColumnValue(row, "Style");
                            var color = GetColumnValue(row, "AttributeValue2");

                            #region FOR BAUER
                            if (styleSubtotalCount > 0 && preStyle.Length > 0 && (style != preStyle || i == totalSKUs))  // Add style-level subtotal line
                            {
                                uint currentStyleEndRows = (uint)i + startRow + (uint)styleSubtotalIndex - 1;
                                foreach (string column in comColumnList)
                                {
                                    LoadTabularDataToCell(column, sheet, (uint)i + startRow + (uint)styleSubtotalIndex, null, preDept, preDept, "", "", "", "", i, totalSKUs, styleSubtotalCount,
                                        category.Category);
                                }
                                for (int o = 0; o < orderlist.Count; o++)
                                {
                                    var cellf = "SUM(" + orderlist[o] + currentStyleStartRows.ToString() + ":" + orderlist[o] + currentStyleEndRows.ToString() + ")";
                                    //SetCellValue(cell, , 5, intSubTotalStyleId);
                                    cell = SetCellValue(sheet,  orderlist[o] + ((uint)i + startRow + (uint)styleSubtotalIndex).ToString(),string.Empty, "1");
                                    cell.Formula = cellf;
                                    SetCellStyle(cell, System.Drawing.Color.LightGray);
                                }
                                //total unite funcational
                                var cellformula1 = "SUM(" + CategorySKUUnitTotalColumn + currentStyleStartRows.ToString() + ":" + CategorySKUUnitTotalColumn + currentStyleEndRows.ToString() + ")";
                                //SetCellValue(cell, string.Empty, 3, intSubTotalStyleId);
                                cell = SetCellValue(sheet,  CategorySKUUnitTotalColumn + ((uint)i + startRow + (uint)styleSubtotalIndex).ToString(), string.Empty, "1");
                                cell.Formula = cellformula1;
                                SetCellStyle(cell, System.Drawing.Color.LightGray);

                                //total value funcational
                                var cellformulb1 = "SUM(" + CategorySKUValueTotalColumn + currentStyleStartRows.ToString() + ":" + CategorySKUValueTotalColumn + currentStyleEndRows.ToString() + ")";
                                //SetCellValue(cell, string.Empty, 3, doubleSubTotalStyleId);
                                cell = SetCellValue(sheet, CategorySKUValueTotalColumn + ((uint)i + startRow + (uint)styleSubtotalIndex).ToString(), string.Empty, "2");
                                cell.Formula = cellformulb1;
                                SetCellStyle(cell, System.Drawing.Color.LightGray);

                                // Hide the Subtotal Line if only One SKU for a Style
                                if (currentStyleStartRows == currentStyleEndRows)
                                {
                                    ExcelRow subtotalRow = sheet.Row((int)currentStyleEndRows + 1);
                                    //subtotalRow.CustomHeight = true;
                                    //subtotalRow.Height = 0; //THIS SET WILL MODIFY THE SIZE OF LOGO IMAGE
                                    subtotalRow.Hidden = true;
                                }

                                styleSubtotalIndex++;

                                currentStyleStartRows = (uint)i + startRow + (uint)styleSubtotalIndex;

                                if (i == totalSKUs)  // the last subtotal line
                                {
                                    break;
                                }
                            }
                            #endregion

                            #region set product cell values for the row
                            foreach (string column in comColumnList)
                            {
                                if (column == CategorySKUPriceColumn + "|PriceWholesale")
                                {
                                    uint rowidx = (uint)i + startRow + (uint)styleSubtotalIndex;
                                    //SetCellValue(cell, , 2, CategoryWSPriceEditable ? doubleUnlockedStyleId : doubleStyleId);
                                    cell = SetCellValue(sheet, CategorySKUPriceColumn + rowidx.ToString(), GetColumnValue(row, "PriceWholesale"), "2");
                                    SetCellStyle(cell);
                                    if (CategoryWSPriceEditable)
                                        cell.Style.Locked = false;
                                        
                                    row["RowIdx"] = rowidx;
                                }
                                else
                                {
                                    LoadTabularDataToCell(column, sheet, (uint)i + startRow + (uint)styleSubtotalIndex, row, deptId, preDept, style, preStyle, color, preColor, i, totalSKUs, styleSubtotalCount,
                                        category.Category);
                                }
                            }
                            #endregion

                            #region update begin/end row number for each level2 category
                            if (categoryId != preCategoryId)
                            {
                                var preCategorys = from c in CategoryLevelList where c.CategoryID == preCategoryId && c.ParentID == category.CategoryID select c;
                                if (preCategorys.Count() > 0)
                                {
                                    var presubc = preCategorys.First();
                                    presubc.SKUEndRow = startRow + (uint)i + (uint)styleSubtotalIndex - 1;
                                }
                                var categorys = from c in CategoryLevelList where c.CategoryID == categoryId && c.ParentID == category.CategoryID select c;
                                if (categorys.Count() > 0)
                                {
                                    var subc = categorys.First();
                                    subc.SKUBeginRow = startRow + (uint)i + (uint)styleSubtotalIndex;
                                }
                            }
                            #endregion

                            preCategoryId = categoryId;
                            preDept = deptId;
                            preStyle = style;
                            preColor = color;

                            #region set order columns 1-6 for each sku
                            for (int o = 0; o < OrderColumnList.Count; o++)
                            {
                                cell = SetCellValue(sheet, OrderColumnList[o].SKUColumn + ((uint)i + startRow + (uint)styleSubtotalIndex).ToString(), string.Empty, "1");
                                if (ShipDateRequiredForStart == "1")
                                {
                                    cell.Style.Locked = true;
                                    SetCellStyle(cell, System.Drawing.Color.LightGray);
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(OrderColumnList[o].DateShip))
                                    {
                                        var atpDate = row.IsNull("ATPDateToCompare") ? DateTime.MaxValue : DateTime.Parse(row["ATPDateToCompare"].ToString());
                                        var shipDate = DateTime.Parse(OrderColumnList[o].DateShip);
                                        if (shipDate < atpDate)
                                        {
                                            cell.Style.Locked = true;
                                            SetCellStyle(cell, System.Drawing.Color.LightGray);
                                            continue;
                                        }
                                    }
                                    cell.Style.Locked = false;
                                    SetCellStyle(cell);
                                }
                            }
                            #endregion

                            #region set total units and amount for each sku row
                            //total unite funcational
                            var cellformula = "SUM(" + OrderColumnList[0].SKUColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + OrderColumnList[OrderColumnList.Count - 1].SKUColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ")";
                            //SetCellValue(cell, string.Empty, 3, intStyleId);
                            cell = SetCellValue(sheet,  styleTotalUnitsColumn, (uint)i + startRow + (uint)styleSubtotalIndex, string.Empty, "1");
                            cell.Formula = cellformula;
                            SetCellStyle(cell, cellTotalColor);
                            if (CategorySubTotalForStyle == "1") //FOR BAUER
                                cell.Style.Font.Color.SetColor(System.Drawing.Color.White);

                            //total value funcational
                            var cellformulb = "=SUMPRODUCT(" + CategorySKUPriceColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + CategorySKUPriceColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + "," + styleTotalUnitsColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + styleTotalUnitsColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ")";
                            //SetCellValue(cell, string.Empty, 3, doubleTStyleId);
                            cell = SetCellValue(sheet, styleTotalAmountColumn, (uint)i + startRow + (uint)styleSubtotalIndex, string.Empty, "2");
                            cell.Formula = cellformulb;
                            SetCellStyle(cell, cellTotalColor);
                            if (CategorySubTotalForStyle == "1") //FOR BAUER
                                cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                            #endregion

                            #region ORDER MULTIPLE generate order multiple list so can add to sheet
                            if (OrderMultipleCheck == "1")
                            {
                                var OrderMultipleStr = GetColumnValue(row, "OrderMultiple");
                                decimal OrderMultipleD = 0;
                                try
                                {
                                    decimal.TryParse(OrderMultipleStr, out OrderMultipleD);
                                }
                                catch { OrderMultipleD = 1; }
                                int OrderMultiple = OrderMultipleD <= 0 ? 1 : (int)OrderMultipleD;
                                OrderMultipleDataValidation multiple = null;
                                var multips = from m in multipleList where m.Multiple == OrderMultiple select m;
                                if (multips.Count() <= 0)
                                {
                                    multiple = new OrderMultipleDataValidation { Multiple = OrderMultiple, SequenceOfReferences = OrderColumnList[0].SKUColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + OrderColumnList[OrderColumnList.Count - 1].SKUColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() };
                                    multipleList.Add(multiple);
                                }
                                else
                                {
                                    multiple = multips.First();
                                    multiple.SequenceOfReferences = multiple.SequenceOfReferences + " " + OrderColumnList[0].SKUColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + OrderColumnList[OrderColumnList.Count - 1].SKUColumn + (i + startRow + (uint)styleSubtotalIndex).ToString();
                                }
                            }
                            #endregion

                            if (atpSheet != null && row != null)
                            {
                                var atpDate = row.IsNull("ATPDateToCompare") ? string.Empty : row["ATPDateToCompare"].ToString();
                                SetCellValue(atpSheet, "A", (uint)i + startRow + (uint)styleSubtotalIndex, atpDate);
                            }
                            
                        }

                        #region Footer summary
                        var summaryCols = CategorySKUSummaryColumnList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        var footerBegin = totalSKUs + startRow + (uint)styleSubtotalCount;

                        #region set black row in sheet footer
                        foreach (var col in summaryCols)
                        {
                            //SetCellValue(cCell, string.Empty, 3, textNullBlackStyleId);
                            cCell = SetCellValue(sheet, col, (uint)footerBegin, string.Empty);
                            SetCellStyle(cCell, System.Drawing.Color.Black, false, false, false, false);
                        }
                        for (int o = 0; o < OrderColumnList.Count; o++)
                        {
                            if (!summaryCols.Contains(OrderColumnList[o].SKUColumn))
                            {
                                cCell = SetCellValue(sheet, OrderColumnList[o].SKUColumn, (uint)footerBegin, string.Empty);
                                SetCellStyle(cCell, System.Drawing.Color.Black, false, false, false, false);
                            }
                        }
                        if (!summaryCols.Contains(styleTotalUnitsColumn))
                        {
                            cCell = SetCellValue(sheet, styleTotalUnitsColumn, (uint)footerBegin, string.Empty);
                            SetCellStyle(cCell, System.Drawing.Color.Black, false, false, false, false);
                        }
                        if (!summaryCols.Contains(styleTotalAmountColumn))
                        {
                            cCell = SetCellValue(sheet, styleTotalAmountColumn, (uint)footerBegin, string.Empty);
                            SetCellStyle(cCell, System.Drawing.Color.Black, false, false, false, false);
                        }
                        #endregion

                        var level2s = from c in CategoryLevelList where c.ParentID == category.CategoryID select c;
                        for (int c = 0; c < level2s.Count(); c++)
                        {
                            level2s.ElementAt(c).CategoryRow = (uint)(footerBegin + c + 1);
                            if (level2s.ElementAt(c).SKUEndRow <= 0)
                                level2s.ElementAt(c).SKUEndRow = (uint)(totalSKUs + startRow + (uint)styleSubtotalCount - 1);
                            cCell = null;
                            foreach (var col in summaryCols)
                            {
                                var orderColExists = from oc in OrderColumnList where oc.SKUColumn == col select oc;
                                if (col == CategorySKUSummaryCategoryColumn || col == CategorySKUSummaryCategoryIDColumn ||
                                    col == styleTotalUnitsColumn || col == styleTotalAmountColumn ||
                                    orderColExists.Count() > 0)
                                {
                                    if (col == CategorySKUSummaryCategoryColumn)
                                    {
                                        //SetCellValue(cCell, level2s.ElementAt(c).Category, 1, textAllSubStyleId);
                                        cCell = SetCellValue(sheet, col, (uint)(footerBegin + c + 1), level2s.ElementAt(c).Category);
                                        SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                                    }
                                    if (col == CategorySKUSummaryCategoryIDColumn)
                                    {
                                        //SetCellValue(cCell, , 1, textAllStyleId);
                                        cCell = SetCellValue(sheet, CategorySKUSummaryCategoryIDColumn, (uint)(footerBegin + c + 1), level2s.ElementAt(c).CategoryID);
                                        SetCellStyle(cCell);
                                    }
                                    if (orderColExists.Count() > 0 || col == styleTotalUnitsColumn)
                                    {
                                        var cellformula = "SUM(" + col + level2s.ElementAt(c).SKUBeginRow.ToString() + ":" + col + level2s.ElementAt(c).SKUEndRow.ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                                        //SetCellValue(cCell, string.Empty, 3, intAllSubStyleId);
                                        cCell = SetCellValue(sheet, col, (uint)(footerBegin + c + 1), string.Empty, "1");
                                        cCell.Formula = cellformula;
                                        SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                                    }
                                    if (col == styleTotalAmountColumn)
                                    {
                                        var cellformula = "SUM(" + col + level2s.ElementAt(c).SKUBeginRow.ToString() + ":" + col + level2s.ElementAt(c).SKUEndRow.ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                                        //SetCellValue(cCell, string.Empty, 3, doubleAllSubStyleId);
                                        cCell = SetCellValue(sheet, col, (uint)(footerBegin + c + 1), string.Empty, "2");
                                        cCell.Formula = cellformula;
                                        SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                                    }
                                }
                                else
                                {
                                    //SetCellValue(cCell, string.Empty, 3, textAllSubStyleId);
                                    cCell = SetCellValue(sheet, col, (uint)(footerBegin + c + 1), string.Empty);
                                    SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                                }
                            }
                            for (int o = 0; o < OrderColumnList.Count; o++)
                            {
                                if (!summaryCols.Contains(OrderColumnList[o].SKUColumn))
                                {
                                    var cellformula = "SUM(" + OrderColumnList[o].SKUColumn + level2s.ElementAt(c).SKUBeginRow.ToString() + ":" + OrderColumnList[o].SKUColumn + level2s.ElementAt(c).SKUEndRow.ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                                    cCell = SetCellValue(sheet, OrderColumnList[o].SKUColumn, (uint)(footerBegin + c + 1), string.Empty, "1");
                                    cCell.Formula = cellformula;
                                    SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                                }
                            }
                            if (!summaryCols.Contains(styleTotalUnitsColumn))
                            {
                                var cellformula = "SUM(" + styleTotalUnitsColumn + level2s.ElementAt(c).SKUBeginRow.ToString() + ":" + styleTotalUnitsColumn + level2s.ElementAt(c).SKUEndRow.ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                                cCell = SetCellValue(sheet, styleTotalUnitsColumn, (uint)(footerBegin + c + 1), string.Empty, "1");
                                cCell.Formula = cellformula;
                                SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                            }
                            if (!summaryCols.Contains(styleTotalAmountColumn))
                            {
                                var cellformula = "SUM(" + styleTotalAmountColumn + level2s.ElementAt(c).SKUBeginRow.ToString() + ":" + styleTotalAmountColumn + level2s.ElementAt(c).SKUEndRow.ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                                cCell = SetCellValue(sheet, styleTotalAmountColumn, (uint)(footerBegin + c + 1), string.Empty, "2");
                                cCell.Formula = cellformula;
                                SetCellStyle(cCell, System.Drawing.Color.WhiteSmoke);
                            }
                        }
                        #endregion

                        #region ORDER MULTIPLE Update Numeric Validation
                        if (OrderMultipleCheck == "1")
                        {
                            multipleList.ForEach(m =>
                            {
                                var validation = sheet.DataValidations.AddCustomValidation(m.SequenceOfReferences);
                                validation.AllowBlank = true;
                                validation.ShowInputMessage = true;
                                validation.ShowErrorMessage = true;
                                validation.ErrorTitle = "Multiple Value Cell";
                                validation.Error = "You must enter a multiple of " + m.Multiple.ToString() + " in this cell.";
                                validation.Formula.ExcelFormula = "(MOD(INDIRECT(ADDRESS(ROW(),COLUMN()))," + m.Multiple.ToString() + ")=0)";
                            });
                        }
                        #endregion

                        if (!string.IsNullOrEmpty(CategorySKUAutoFilterRange))
                        {
                            sheet.Cells[CategorySKUAutoFilterRange].AutoFilter = true;
                        }
                    }
                    spreadSheet.Save();
                }
            }
            return dt;
        }

        #region GetCategoryTabularOrderFormData [p_Offline_CategoryTabularGridView]
        private DataTable GetCategoryTabularOrderFormData(string soldto, string catalog, string Category, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_CategoryTabularGridView", conn);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
                cmd.Parameters["@catcd"].Value = catalog;

                cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
                cmd.Parameters["@SoldTo"].Value = soldto;

                cmd.Parameters.Add(new SqlParameter("@Category", SqlDbType.VarChar, 80));
                cmd.Parameters["@Category"].Value = Category;

                cmd.Parameters.Add(new SqlParameter("@PriceType", SqlDbType.VarChar, 80));
                cmd.Parameters["@PriceType"].Value = string.IsNullOrEmpty(PriceType) ? "W" : PriceType;

                try
                {
                    DataSet outds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(outds);
                    conn.Close();
                    if (outds.Tables.Count > 0)
                    {
                        return outds.Tables[0];
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    //throw;
                    WriteToLog("p_Offline_TabularGridView | GetTabularOrderFormData |" + soldto + "|" + catalog + "|" + Category, ex, savedirectory);
                    return null;
                }
            }
        }
        #endregion

        private void LoadTabularDataToCell(string columnInfo, ExcelWorksheet sheet, uint rowIndex, DataRow row,
            string deptId, string preDept, string style, string preStyle, string color, string preColor, int loopIndex, int totalCount, int subtotalCount,
            string topCategory)
        {
            ExcelRange cell = null;
            var colInfo = columnInfo.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            switch (colInfo[1])
            {
                case "Catalog":
                    string catalogcd = GetColumnValue(row, "CatalogCode");
                    if (catalogcd == string.Empty)
                        //SetCellValue(cell, "*", 1, textAllStyleId);
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, "*");
                    else
                        //SetCellValue(cell, , 1, textAllStyleId);
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, GetColumnValue(row, "CatalogCode"));
                    SetCellStyle(cell);
                    break;
                case "Department":
                    var NavigationBar = GetColumnValue(row, "NavigationBar");
                    var topCategoryReg = topCategory.Replace("(", "\\(").Replace(")", "\\)");
                    Regex regEx = new Regex(topCategoryReg + " > ", RegexOptions.Multiline);
                    NavigationBar = regEx.Replace(NavigationBar, "", 1);
                    if (CategorySKUCombineDepartment == "1")
                    {
                        if (deptId != preDept)
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                //SetCellValue(cell, NavigationBar, 1, textDAllStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, NavigationBar);
                                SetCellStyle(cell);
                                cell.Style.WrapText = true;
                            }
                            else
                            {
                                //SetCellValue(cell, NavigationBar, 1, textDLRTStyleId);//DepartmentName
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, NavigationBar);
                                SetCellStyle(cell, null, true, true, true, false);
                                cell.Style.WrapText = true;
                            }
                        }
                        else
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                //SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                                SetCellStyle(cell, null, true, true, false, true);
                            }
                            else
                            {
                                //SetCellValue(cell, string.Empty, 1, textLRStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                                SetCellStyle(cell, null, true, true, false, false);
                            }
                        }
                    }
                    else
                    {
                        //SetCellValue(cell, NavigationBar, 1, textDAllStyleId);
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, NavigationBar);
                        SetCellStyle(cell);
                    }
                    break;
                case "Style":
                    if (CategorySKUCombineStyle == "1")
                    {
                        if ((style != preStyle) || (style == preStyle && deptId != preDept))
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                //SetCellValue(cell, style, 1, textAllStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, style);
                                SetCellStyle(cell);
                            }
                            else
                            {
                                //SetCellValue(cell, style, 1, textLRTStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, style);
                                SetCellStyle(cell, null, true, true, true, false);
                            }
                        }
                        else
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                //SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                                SetCellStyle(cell, null, true, true, false, true);
                            }
                            else
                            {
                                //SetCellValue(cell, string.Empty, 1, textLRStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                                SetCellStyle(cell, null, true, true, false, false);
                            }
                        }
                    }
                    else
                    {
                        //SetCellValue(cell, style, 1, textAllStyleId);
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, style);
                        SetCellStyle(cell);
                    }
                    break;
                case "ProductName":
                    var productName = GetColumnValue(row, "ProductName");
                    if (CategorySKUCombineProductName == "1")
                    {
                        if ((style != preStyle) || (style == preStyle && deptId != preDept))
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                //SetCellValue(cell, productName, 1, textAllStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, productName);
                                SetCellStyle(cell);
                            }
                            else
                            {
                                //SetCellValue(cell, productName, 1, textLRTStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, productName);
                                SetCellStyle(cell, null, true, true, true, false);
                            }
                        }
                        else
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                //SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                                SetCellStyle(cell, null, true, true, false, true);
                            }
                            else
                            {
                                //SetCellValue(cell, string.Empty, 1, textLRStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                                SetCellStyle(cell, null, true, true, false, false);
                            }
                        }
                    }
                    else
                    {
                        //SetCellValue(cell, productName, 1, textDAllStyleId); //textAllStyleId);
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, productName);
                        SetCellStyle(cell);
                    }
                    break;
                case "AttributeValue2":
                    if (CategorySKUCombineColor == "1")
                    {
                        if ((color != preColor) || (color == preColor && deptId != preDept) || (color == preColor && style != preStyle))
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                //SetCellValue(cell, color, 1, textAllStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, color);
                                SetCellStyle(cell);
                            }
                            else
                            {
                                //SetCellValue(cell, color, 1, textLRTStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, color);
                                SetCellStyle(cell, null, true, true, true, false);
                            }
                        }
                        else
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                //SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                                SetCellStyle(cell, null, true, true, false, true);
                            }
                            else
                            {
                                //SetCellValue(cell, string.Empty, 1, textLRStyleId);
                                cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                                SetCellStyle(cell, null, true, true, false, false);
                            }
                        }
                    }
                    else
                    {
                        //SetCellValue(cell, color, 1, textAllStyleId);
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, color);
                        SetCellStyle(cell);
                    }
                    break;
                case "LaunchDate":
                    var cvalue = GetColumnValue(row, colInfo[1]);
                    DateTime dt;
                    if (cvalue.Length > 0 && DateTime.TryParse(cvalue, out dt))
                    {
                        var tformula = string.Format("DATE({0},{1},{2})", dt.Year, dt.Month, dt.Day);
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty, "3");
                        cell.Formula = tformula;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    break;
                case "HistoryData":
                    {
                        var historyRealValue = 0;
                        var historyData = GetColumnValue(row, colInfo[1]);
                        if (int.TryParse(historyData, out historyRealValue))
                        {
                            cell = SetCellValue(sheet, colInfo[0], rowIndex, historyData, "1");
                            SetCellStyle(cell);
                        }
                        else
                        {
                            cell = SetCellValue(sheet, colInfo[0], rowIndex, historyData);
                            if (!string.IsNullOrEmpty(historyData))
                            {
                                SetCellStyle(cell, System.Drawing.Color.Yellow);
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            }
                            else
                                SetCellStyle(cell);
                        }
                        break;
                    }
                case "Level2DeptID":
                case "SKU":
                case "UPC":
                case "AttributeValue1":
                case "AttributeValue5":
                case "Gender":
                default:
                    //SetCellValue(cell, , 1, textAllStyleId);
                    cell = SetCellValue(sheet, colInfo[0], rowIndex, GetColumnValue(row, colInfo[1]));
                    SetCellStyle(cell);
                    break;
            }
        }

        #endregion

        #region SaveCategoryListToSheet
        private void SaveCategoryListToSheet(string filePath)
        {
            using (ExcelPackage spreadSheet = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = null;
                sheet = spreadSheet.Workbook.Worksheets[CategoryListSheetName];
                if(sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + CategoryListSheetName + ", file path:" + filePath);

                if (sheet != null)
                {
                    uint index = 2;
                    foreach (CategoryLevelInfo category in CategoryLevelList)
                    {
                        SetCellValue(sheet, "A" + index.ToString(), category.CategoryID);
                        SetCellValue(sheet, "B" + index.ToString(), category.Category);
                        SetCellValue(sheet, "C" + index.ToString(), category.ParentID);
                        SetCellValue(sheet, "D" + index.ToString(), category.ParentCategory);
                        SetCellValue(sheet, "E" + index.ToString(), category.CategoryRow.ToString());
                        SetCellValue(sheet, "F" + index.ToString(), category.SKUBeginRow.ToString());
                        SetCellValue(sheet, "G" + index.ToString(), category.SKUEndRow.ToString());
                        SetCellValue(sheet, "H" + index.ToString(), category.SKURowNumber.ToString());
                        SetCellValue(sheet, "I" + index.ToString(), category.IsTopLevel ? "1" : "0");
                        index++;
                    }
                }
                spreadSheet.Save();
            }
        }
        #endregion 

        #region SaveTabularCustomerDataToSheet
        private DataTable GetTabularCustomerData(string soldto, string catalog, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_TabularCustomerList", conn);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
                cmd.Parameters["@catcd"].Value = catalog;

                cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
                cmd.Parameters["@SoldTo"].Value = soldto;

                try
                {
                    DataSet outds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(outds);
                    conn.Close();
                    if (outds.Tables.Count > 0)
                    {
                        return outds.Tables[0];
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    //throw;
                    WriteToLog("p_Offline_TabularCustomerList | GetTabularCustomerData |" + catalog, ex, savedirectory);
                    return null;
                }
            }
        }

        private void SaveTabularCustomerDataToSheet(DataTable customerData, string filePath)
        {
            var sheetName = ConfigurationManager.AppSettings["CustomerSheetName"] == null ? "SoldToShipTo" : ConfigurationManager.AppSettings["CustomerSheetName"];
            var startColumn = ConfigurationManager.AppSettings["CustomerStartColumn"] == null ? "A" : ConfigurationManager.AppSettings["CustomerStartColumn"];
            var columnLength = ConfigurationManager.AppSettings["CustomerColumnLength"] == null ? 16 : int.Parse(ConfigurationManager.AppSettings["CustomerColumnLength"]);

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = null;
                sheet = package.Workbook.Worksheets[sheetName];
                if (sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + filePath);
                if (sheet != null)
                {
                    var columnList = new List<string>();
                    for (int i = 0; i < columnLength; i++)
                    {
                        var col = ((char)(char.Parse(startColumn) + i)).ToString();
                        WriteToLog(col);
                        var cell = GetCell(sheet, col + "1");
                        if (cell != null && cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                        {
                            columnList.Add(cell.Value.ToString());
                            WriteToLog(cell.Value.ToString());
                        }
                        else
                            columnList.Add(string.Empty);
                    }
                    for (int i = 0; i < customerData.Rows.Count; i++)
                    {
                        var row = customerData.Rows[i];
                        for (int j = 0; j < columnLength; j++)
                        {
                            var col = ((char)(char.Parse(startColumn) + j)).ToString();
                            SetCellValue(sheet, col, (uint)(i + 2), GetColumnValue(row, columnList[j]));
                        }
                        sheet.Row(i + 2).Hidden = true;
                    }
                    sheet.Row(1).Hidden = true;
                    package.Save();
                }
            }
        }
        #endregion

        private void SaveCategoryDateValidationToSheet(DataRow catalogRow, string filePath)
        {
            var sheetName = ConfigurationManager.AppSettings["GlobalSheetName"] == null ? "GlobalVariables" : ConfigurationManager.AppSettings["GlobalSheetName"];
            using (ExcelPackage book = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = null;
                sheet = book.Workbook.Worksheets[sheetName];
                if(sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + filePath);

                if (sheet != null)
                {
                    //BOOKING ORDER
                    var bookingOrder = GetColumnValue(catalogRow, "BookingCatalog");
                    SetCellValue(sheet, "B3", (string.IsNullOrEmpty(bookingOrder) ? "0" : bookingOrder));
                    //DateShipStart
                    if (!catalogRow.IsNull("DateShipStart"))
                    {
                        var startdate = (DateTime)catalogRow["DateShipStart"];
                        var windowdates = from o in OrderColumnList where o.DateShipDate != null && o.DateShipDate > DateTime.MinValue select o.DateShipDate;
                        if (windowdates != null && windowdates.Count() > 0)
                        {
                            var windowmax = windowdates.Max();
                            if (windowmax > startdate)
                                startdate = windowmax;
                        }
                        var date = (startdate).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);//1033
                        SetCellValue(sheet, "B4", date);
                    }
                    //DateShipEnd
                    if (!catalogRow.IsNull("DateShipEnd"))
                    {
                        var date = ((DateTime)catalogRow["DateShipEnd"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);//1033
                        SetCellValue(sheet, "B5", date);
                    }
                    //ReqDateDays
                    var ReqDateDays = (int)catalogRow["ReqDateDays"];
                    SetCellValue(sheet, "B6", ReqDateDays.ToString());
                    //ReqDefaultDays
                    var ReqDefaultDays = (int)catalogRow["DefaultReqDays"];
                    SetCellValue(sheet, "B7", ReqDefaultDays.ToString());
                    //CancelDateDays
                    var CancelDateDays = (int)catalogRow["CancelDateDays"];
                    SetCellValue(sheet, "B8", CancelDateDays.ToString());
                    //CancelDefaultDays
                    var CancelDefaultDays = (int)catalogRow["CancelDefaultDays"];
                    SetCellValue(sheet, "B9", CancelDefaultDays.ToString());
                    //Category Numbers
                    SetCellValue(sheet, "B10", CategoryLevelList.Count.ToString());
                    //Ship Window Count
                    SetCellValue(sheet, "B11", OrderColumnList.Count.ToString());
                }
                book.Save();
            }
        }

        private void ProtectWorkbook(string filename, string savedirectory)
        {
            var pwd = ConfigurationManager.AppSettings["ProtectPassword"] == null ? "Plumriver" : ConfigurationManager.AppSettings["ProtectPassword"];
            var password = HashPassword(pwd);
            var filePath = Path.Combine(savedirectory, filename);
            using (ExcelPackage book = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = null;
                sheet = book.Workbook.Worksheets[SummarySheetName];
                if (sheet != null)
                {
                    sheet.Protection.AllowAutoFilter = true;
                    sheet.Protection.IsProtected = true;
                    sheet.Protection.AllowEditObject = true;
                    sheet.Protection.AllowEditScenarios = true;
                    sheet.Protection.SetPassword(password);
                }

                var tops = from c in CategoryLevelList where string.IsNullOrEmpty(c.ParentID) select c;
                foreach (CategoryLevelInfo category in tops)
                {
                    sheet = null;
                    sheet = book.Workbook.Worksheets[category.Category];
                    if (sheet != null)
                    {
                        //SET PERCENTAGE OF SHEET, CCM 21358
                        var CategorySKUSheetPercentage = ConfigurationManager.AppSettings["CategorySKUSheetPercentage"];
                        if (!string.IsNullOrEmpty(CategorySKUSheetPercentage))
                        {
                            var CategorySKUSheetPercentageInt = 0;
                            if (int.TryParse(CategorySKUSheetPercentage, out CategorySKUSheetPercentageInt))
                                sheet.View.ZoomScale = CategorySKUSheetPercentageInt;
                        }
                        //autofit columns OF SHEET, CCM 21358
                        var CategorySKUSheetAutoFitAddress = ConfigurationManager.AppSettings["CategorySKUSheetAutoFitAddress"];
                        if (!string.IsNullOrEmpty(CategorySKUSheetAutoFitAddress))
                        {
                            sheet.Cells[CategorySKUSheetAutoFitAddress].AutoFitColumns();
                        }
                        //HIDE UPC COLUMN OF SHEET, CCM 21358
                        var CategorySKUSheetHideUPCColumn = ConfigurationManager.AppSettings["CategorySKUSheetHideUPCColumn"]; //4
                        if (!string.IsNullOrEmpty(CategorySKUSheetHideUPCColumn))
                        {
                             var CategorySKUSheetHideUPCColumnInt = 0;
                             if (int.TryParse(CategorySKUSheetHideUPCColumn, out CategorySKUSheetHideUPCColumnInt))
                                 sheet.Column(CategorySKUSheetHideUPCColumnInt).Hidden = true;
                        }
                        sheet.Protection.AllowAutoFilter = true;
                        sheet.Protection.IsProtected = true;
                        sheet.Protection.AllowEditObject = true;
                        sheet.Protection.AllowEditScenarios = true;
                        sheet.Protection.SetPassword(password);
                    }
                }
                book.Workbook.Protection.LockStructure = true;

                var gsheetName = ConfigurationManager.AppSettings["GlobalSheetName"] == null ? "GlobalVariables" : ConfigurationManager.AppSettings["GlobalSheetName"];
                var pwdColumn = ConfigurationManager.AppSettings["GlobalPWDColumn"] == null ? "B" : ConfigurationManager.AppSettings["GlobalPWDColumn"];
                var pwdRow = ConfigurationManager.AppSettings["GlobalPWDRow"] == null ? 200 : int.Parse(ConfigurationManager.AppSettings["GlobalPWDRow"]);

                sheet = null;
                sheet = book.Workbook.Worksheets[gsheetName];
                if (sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + gsheetName + ", file path:" + filePath);
                if (sheet != null)
                {
                    SetCellValue(sheet,  pwdColumn + pwdRow.ToString(), pwd);
                }
                book.Save();
            }
        }

        private void GetOrderColumns(string catalog, string savedirectory, DataRow catalogRow)
        {
            OrderColumnList = new List<ShipWindowInfo>();
            var orderlist = CategorySKUOrderColumnList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            var summaryColList = SummaryCategoryColumns.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (CategoryEnableShipWindow == "1") //for CCM EPC
            {
                var CategoryMinSheetCount = ConfigurationManager.AppSettings["CategoryMinSheetCount"] == null ? 6 : int.Parse(ConfigurationManager.AppSettings["CategoryMinSheetCount"]);
                var dt = GetCatalogShipWindow(catalog, savedirectory);
                if (dt != null && dt.Rows.Count > 0)
                {
                    var CatalogSpreadsheetStartShipDate = catalogRow.Table.Columns.Contains("CatalogSpreadsheetStartShipDate") ?
                        catalogRow["CatalogSpreadsheetStartShipDate"].ToString() : string.Empty;
                    var catalogStartDate = DateTime.MinValue;
                    var firstSKUColumn = char.Parse(orderlist[0]);
                    var firstSummaryColumn = char.Parse(summaryColList[1]);
                    if (CatalogSpreadsheetStartShipDate == "1")
                    {
                        catalogStartDate = (DateTime)catalogRow["DateShipStart"];
                        var cshipinfo = new ShipWindowInfo
                                    {
                                        SKUColumn = ((char)(firstSKUColumn)).ToString(),
                                        SummaryColumn = ((char)(firstSummaryColumn)).ToString(),
                                        DateShip = string.Format(@"{0:MM/dd/yyyy}", catalogStartDate),
                                        DateShipDate = catalogStartDate
                                    };
                        OrderColumnList.Add(cshipinfo);
                    }
                    var wIndex = 0;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var row = dt.Rows[i];
                        var shipWindowDate = (DateTime)row["DateShip"];
                        if (shipWindowDate > catalogStartDate)
                        {
                            wIndex = i;
                            if (CatalogSpreadsheetStartShipDate == "1")
                                wIndex++;
                            var shipinfo = new ShipWindowInfo
                            {
                                SKUColumn = ((char)(firstSKUColumn + wIndex)).ToString(),
                                SummaryColumn = ((char)(firstSummaryColumn + wIndex)).ToString(),
                                DateShip = row.IsNull("DateShip") ? string.Empty : string.Format(@"{0:MM/dd/yyyy}", (DateTime)row["DateShip"]),
                                DateShipDate = row.IsNull("DateShip") ? DateTime.MinValue : (DateTime)row["DateShip"]
                            };
                            OrderColumnList.Add(shipinfo);
                        }
                    }
                }
                if (OrderColumnList.Count > 0 && CategoryMinSheetCount > OrderColumnList.Count)
                {
                    var lastSKUColumn = OrderColumnList.Last().SKUColumn;
                    var lastSummaryColumn = OrderColumnList.Last().SummaryColumn;
                    var shipWindowCount = OrderColumnList.Count;
                    for (int i = 1; i <= CategoryMinSheetCount - shipWindowCount; i++)
                    {
                        var newSKUColumn = ((char)(char.Parse(lastSKUColumn) + i)).ToString();
                        var newSummaryColumn = ((char)(char.Parse(lastSummaryColumn) + i)).ToString();
                        var shipinfo = new ShipWindowInfo
                        {
                            SKUColumn = newSKUColumn,
                            SummaryColumn = newSummaryColumn,
                            DateShipDate = DateTime.MinValue
                        };
                        OrderColumnList.Add(shipinfo);
                    }
                }
            }
            if(OrderColumnList.Count <= 0)
            {
                for(int i=0;i<orderlist.Count;i++)
                {
                    var shipinfo = new ShipWindowInfo { SKUColumn = orderlist[i], SummaryColumn = summaryColList[i + 1], DateShipDate = DateTime.MinValue };
                    OrderColumnList.Add(shipinfo);
                }
            }
        }

        private DataTable GetCatalogShipWindow(string catalog, string savedirectory)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("Search");
            doc.AppendChild(root);
            List<string> OrderInfo = new List<string>() { "CatalogCode", "Enabled" };
            foreach (string col in OrderInfo)
            {
                XmlAttribute colnode = doc.CreateAttribute(col);
                switch (col)
                {
                    case "CatalogCode":
                        colnode.Value = catalog;
                        break;
                    case "Enabled":
                        colnode.Value = "1";
                        break;
                }
                root.Attributes.Append(colnode);
            }

            string connString = ConfigurationManager.ConnectionStrings["HDIConnString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Admin_CMS_CatalogShipWindowGetData", conn);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@WebSite", SqlDbType.VarChar, 80));
                cmd.Parameters["@WebSite"].Value = string.Empty;

                cmd.Parameters.Add(new SqlParameter("@XML", SqlDbType.VarChar, -1));
                cmd.Parameters["@XML"].Value = doc.InnerXml;

                try
                {
                    DataSet outds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(outds);
                    conn.Close();
                    if (outds.Tables.Count > 0 && outds.Tables[0].Rows.Count > 0)
                    {
                        return outds.Tables[0];
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    //throw;
                    WriteToLog("p_Admin_CMS_CatalogShipWindowGetData |" + catalog, ex, savedirectory);
                    return null;
                }
            }
        }

        private DataTable CodeMethods(string keyword, string savedirectory)
        {
            SqlParameter parmParentTable = new SqlParameter("@parent_table", SqlDbType.VarChar, 50);
            SqlParameter parmKeyword = new SqlParameter("@keyword", SqlDbType.VarChar, 50);

            parmParentTable.Value = string.Empty;
            parmKeyword.Value = keyword;

            DataTable dtCollection = null;
            DataSet dsCollection = new DataSet();

            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            try
            {
                using (SqlConnection conn = new SqlConnection(connString))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("pp_GetCodeValues", conn);
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add(new SqlParameter("@parent_table", SqlDbType.NVarChar, 80));
                    cmd.Parameters["@parent_table"].Value = string.Empty;

                    cmd.Parameters.Add(new SqlParameter("@keyword", SqlDbType.NVarChar, 80));
                    cmd.Parameters["@keyword"].Value = keyword;

                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(dsCollection);
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                //throw;
                WriteToLog("CodeMethods |" + keyword, ex, savedirectory);
                return null;
            }

            if (dsCollection != null && dsCollection.Tables.Count > 0)
            {
                dtCollection = dsCollection.Tables[0];
            }

            return dtCollection;
        }


        /********************************************************************************/
        /*************COMMON METHODS*********************/
        /********************************************************************************/
        private string GetColumnValue(DataRow row, string ColumnName)
        {
            if (row == null)
            {
                return string.Empty;
            }
            else if (row.Table.Columns.Contains(ColumnName))
            {
                if (row.IsNull(ColumnName))
                    return string.Empty;
                return row[ColumnName].ToString();
            }
            return string.Empty;
        }

        private ExcelRange SetCellValue(ExcelWorksheet sheet, string column, uint row, string value)
        {
            return SetCellValue(sheet, column, row, value, string.Empty);
        }

        private ExcelRange SetCellValue(ExcelWorksheet sheet, string column, uint row, string value, string dataType)
        {
            var address = column + row.ToString();
            return SetCellValue(sheet, address, value, dataType);
        }

        private ExcelRange SetCellValue(ExcelWorksheet sheet, string address, string value)
        {
            return SetCellValue(sheet, address, value, string.Empty);
        }

        private ExcelRange SetCellValue(ExcelWorksheet sheet, string address, string value, string dataType)
        {
            ExcelRange cell = sheet.Cells[address];
            switch (dataType)
            {
                case "1": //int
                    cell.Style.Numberformat.Format = "#,##0";
                    if(!string.IsNullOrEmpty(value))
                        cell.Value = int.Parse(value);
                    break;
                case "2": //double
                    cell.Style.Numberformat.Format = "#,##0.00";
                    if (!string.IsNullOrEmpty(value))
                        cell.Value = double.Parse(value);
                    break;
                case "3": //date
                    cell.Style.Numberformat.Format = "yyyy/m/d";
                    cell.Value = value;
                    break;
                default:
                    cell.Value = value;
                    break;
            }
                
            return cell;
        }

        private string GetCellValue(ExcelWorksheet sheet, string address)
        {
            var cellvalue = sheet.Cells[address].Value;
            return cellvalue.ToString();
        }

        private ExcelRange GetCell(ExcelWorksheet sheet, string address)
        {
            var cell = sheet.Cells[address];
            return cell;
        }

        private void SetCellStyle(ExcelRange cell)
        {
            SetCellStyle(cell, null);
        }

        private void SetCellStyle(ExcelRange cell, System.Drawing.Color? bgColor)
        {
            SetCellStyle(cell, bgColor, true, true, true, true);
        }

        private void SetCellStyle(ExcelRange cell, System.Drawing.Color? bgColor, bool borderLeft, bool borderRight, bool borderTop, bool borderBottom)
        {
            if (bgColor != null)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor((System.Drawing.Color)bgColor);
            }
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            if(!borderLeft)
                cell.Style.Border.Left.Style = ExcelBorderStyle.None;
            if (!borderRight)
                cell.Style.Border.Right.Style = ExcelBorderStyle.None;
            if (!borderTop)
                cell.Style.Border.Top.Style = ExcelBorderStyle.None;
            if (!borderBottom)
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.None;
        }

        protected string HashPassword(string password)
        {
            /*byte[] passwordCharacters = System.Text.Encoding.ASCII.GetBytes(password);
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

            return Convert.ToString(hash, 16).ToUpperInvariant();*/
            return password.Trim();
        }


        /*
        public void GenerateCategoryTabularOrderForm(string templateFile, string filename, string soldto, string catalog, string pricecode, string savedirectory)
        {
            WriteToLog(DateTime.Now.ToString() + " Begin " + filename);
            if (File.Exists(templateFile))
            {
                try
                {
                    //save new file
                    var filePath = Path.Combine(savedirectory, filename);
                    File.Copy(templateFile, filePath, true);
                    //set catalog name
                    var catalogRow = GetCatalogInfo(soldto, catalog, savedirectory);
                    CurrencyCode = GetColumnValue(catalogRow, "CurrencyCode");
                    //get level1 and level2 category list
                    GetTabularCategoryData(soldto, catalog, savedirectory);
                    //save category list to Summary sheet, //----and also currency
                    SaveCategoryDataToSummarySheet(catalogRow, filename, savedirectory);
                    //generate top category sheets
                    SaveCategoryProductsToSheet(soldto, catalog, filePath, savedirectory);
                    //save category list
                    SaveCategoryListToSheet(filePath);
                    //validation
                    SaveCategoryDateValidationToSheet(catalogRow, filePath);
                    //protect
                    ProtectWorkbook(filename, savedirectory);
                }
                catch (Exception ex)
                {
                    WriteToLog(ex.Message);
                    WriteToLog(ex.StackTrace);
                }
            }
            else
                WriteToLog("Cannot find template file " + templateFile);
            WriteToLog(DateTime.Now.ToString() + " End " + filename);
        }

        #region GetCatalogInfo
        private DataRow GetCatalogInfo(string soldto, string catalog, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_TabularCatalogInfo", conn);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
                cmd.Parameters["@catcd"].Value = catalog;

                cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
                cmd.Parameters["@SoldTo"].Value = soldto;

                try
                {
                    DataSet outds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(outds);
                    conn.Close();
                    if (outds.Tables.Count > 0 && outds.Tables[0].Rows.Count > 0)
                    {
                        return outds.Tables[0].Rows[0];
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    //throw;
                    WriteToLog("p_Offline_TabularCatalogInfo |" + catalog, ex, savedirectory);
                    return null;
                }
            }
        }
        #endregion

        #region GetTabularCategoryData [p_Offline_CategoryTabularCategoryList]
        private void GetTabularCategoryData(string soldto, string catalog, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_CategoryTabularCategoryList", conn);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
                cmd.Parameters["@catcd"].Value = catalog;

                cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
                cmd.Parameters["@SoldTo"].Value = soldto;

                try
                {
                    DataSet outds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(outds);
                    conn.Close();
                    CategoryLevelList = new List<CategoryLevelInfo>();
                    if (outds.Tables.Count > 0 && outds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow row in outds.Tables[0].Rows)
                        {
                            CategoryLevelInfo category = new CategoryLevelInfo
                            {
                                CategoryID = row["CategoryID"].ToString(),
                                Category = row["Category"].ToString(),
                                ParentID = row.IsNull("ParentID") ? string.Empty : row["ParentID"].ToString(),
                                ParentCategory = row.IsNull("ParentCategory") ? string.Empty : row["ParentCategory"].ToString(),
                                IsTopLevel = row["TopLevel"].ToString() == "1" ? true : false
                            };
                            CategoryLevelList.Add(category);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //throw;
                    WriteToLog("p_Offline_CategoryTabularCategoryList | " + catalog + " | " + soldto, ex, savedirectory);
                }
            }
        }
        #endregion

        private void SaveCategoryDataToSummarySheet(DataRow catalogRow, string filename, string savedirectory)
        {
            var filePath = Path.Combine(savedirectory, filename);
            uint index = 0;
            List<string> summaryColList = SummaryCategoryColumns.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            Worksheet sheet = null;
            SpreadsheetDocument book = SpreadsheetDocument.Open(filePath, true);
            var sheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == SummarySheetName);
            if (sheets.Count() > 0)
            {
                WorksheetPart worksheetPart = (WorksheetPart)book.WorkbookPart.GetPartById(sheets.First().Id.Value);
                sheet = worksheetPart.Worksheet;
            }
            else
                WriteToLog("Wrong template order form file, cannot find sheet " + SummarySheetName + ", file path:" + filePath);

            if (sheet != null && CategoryLevelList.Count > 0)
            {
                var styleSheet = book.WorkbookPart.WorkbookStylesPart.Stylesheet;
                CommonFontId = CreateFontFormat(styleSheet);
                CommonBorderAllId = CreateBorderFormat(styleSheet, true, true, true, true);
                //var textAllStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, UInt32Value.FromUInt32(49));
                CommonTextFormat = CreateCellFormat(styleSheet, CommonFontId, null, null, UInt32Value.FromUInt32(49));
                CommonBoldTextId = CreateCellFormat(styleSheet, bolderfont, null, CommonBorderAllId, UInt32Value.FromUInt32(49));
                var textNullBoldStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid1, null, UInt32Value.FromUInt32(49));
                var textNullRightBoldStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid2, null, UInt32Value.FromUInt32(49), true, false, HorizontalAlignmentValues.Right);
                var doubleAllMainStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid1, CommonBorderAllId, numformatid);
                var doubleAllSubStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid2, CommonBorderAllId, numformatid);
                var intAllMainStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid1, CommonBorderAllId, intformatid);
                var intAllSubStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid2, CommonBorderAllId, intformatid);
                var doublePinkAllSubStyleId = CreateCellFormat(styleSheet, CommonFontId, pinkfileid, CommonBorderAllId, numformatid);
                var intPinkAllMainStyleId = CreateCellFormat(styleSheet, CommonFontId, pinkfileid, CommonBorderAllId, intformatid);

                //save catalog code to A6, save catalog name to B6
                var Locations = CatalogCodeLocation.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                var ccell = GetCell(sheet, Locations[0], uint.Parse(Locations[1]));
                SetCellValue(ccell, GetColumnValue(catalogRow, "CatalogCode"), 1, textNullBoldStyleId);
                Locations = CatalogNameLocation.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                ccell = GetCell(sheet, Locations[0], uint.Parse(Locations[1]));
                SetCellValue(ccell, GetColumnValue(catalogRow, "CatalogName"), 1, textNullBoldStyleId);

                //total currency
                ccell = GetCell(sheet, SummaryCurrencyColumn, SummaryCategoryStartRow - 1);
                SetCellValue(ccell, CurrencyCode, 1, CommonTextFormat);
                //save category level list
                var SUMUnitsFormat = string.Empty;
                var SUMValueFormat = string.Empty;
                var tops = from c in CategoryLevelList where string.IsNullOrEmpty(c.ParentID) select c;
                foreach (CategoryLevelInfo category in tops)
                {
                    if (CategorySubTotalForStyle == "1")
                    {
                        //top category units
                        var cell = GetCell(sheet, "A", index + SummaryCategoryStartRow);
                        SetCellValue(cell, "U0", 1, textNullBoldStyleId);
                        cell = GetCell(sheet, "B", index + SummaryCategoryStartRow);
                        SetCellValue(cell, category.CategoryID, 1, textNullBoldStyleId);
                        cell = GetCell(sheet, summaryColList[0], index + SummaryCategoryStartRow);
                        SetCellValue(cell, category.Category + " Total", 1, textNullBoldStyleId);
                        SUMUnitsFormat = string.IsNullOrEmpty(SUMUnitsFormat) ? "COLUMN" + (index + SummaryCategoryStartRow).ToString() : SUMUnitsFormat + ", COLUMN" + (index + SummaryCategoryStartRow).ToString();
                        var level2s = from c in CategoryLevelList where c.ParentID == category.CategoryID select c;
                        StringBuilder unitSum = new StringBuilder(), subtotalSum = new StringBuilder();
                        for (int i = 1; i < summaryColList.Count; i++)
                        {
                            cell = GetCell(sheet, summaryColList[i], index + SummaryCategoryStartRow);
                            CellFormula tformula = new CellFormula();
                            if (i == summaryColList.Count - 1)
                            {
                                //SUM the units of this row
                                tformula.Text = "SUM(" + subtotalSum + ")";
                            }
                            else if (i == summaryColList.Count - 2)
                            {
                                //SUM the units of this row
                                tformula.Text = "SUM(" + unitSum + ")";
                            }
                            else
                            {
                                if (i % 2 == 1)
                                    unitSum.AppendFormat("{0}{1}{2}", (unitSum.Length > 0 ? "," : string.Empty), summaryColList[i], index + SummaryCategoryStartRow);
                                else
                                    subtotalSum.AppendFormat("{0}{1}{2}", (subtotalSum.Length > 0 ? "," : string.Empty), summaryColList[i], index + SummaryCategoryStartRow);
                                tformula.Text = "SUM(" + summaryColList[i] + (index + 1 + SummaryCategoryStartRow).ToString() + ":" + summaryColList[i] + (index + level2s.Count() + SummaryCategoryStartRow).ToString() + ")";
                            }
                            cell.Append(tformula);
                            SetCellValue(cell, string.Empty, 3, (i % 2 == 1 ? intAllMainStyleId : doubleAllMainStyleId));
                        }
                        index++;
                        //level2 category units
                        foreach (CategoryLevelInfo sub in level2s)
                        {
                            unitSum = new StringBuilder(); subtotalSum = new StringBuilder();
                            cell = GetCell(sheet, "A", index + SummaryCategoryStartRow);
                            SetCellValue(cell, "U", 1, textNullBoldStyleId);
                            cell = GetCell(sheet, "B", index + SummaryCategoryStartRow);
                            SetCellValue(cell, sub.CategoryID, 1, textNullBoldStyleId);
                            cell = GetCell(sheet, summaryColList[0], index + SummaryCategoryStartRow);
                            //SetCellValue(cell, category.Category + " > " + sub.Category, 1, textNullRightBoldStyleId);
                            SetCellValue(cell, sub.Category, 1, textNullRightBoldStyleId);
                            for (int i = 1; i < summaryColList.Count; i++)
                            {
                                cell = GetCell(sheet, summaryColList[i], index + SummaryCategoryStartRow);
                                if (i == summaryColList.Count - 1)
                                {
                                    //SUM the units of this row
                                    CellFormula tformula = new CellFormula();
                                    tformula.Text = "SUM(" + subtotalSum + ")";
                                    cell.Append(tformula);
                                }
                                else if (i == summaryColList.Count - 2)
                                {
                                    //SUM the units of this row
                                    CellFormula tformula = new CellFormula();
                                    tformula.Text = "SUM(" + unitSum + ")";
                                    cell.Append(tformula);
                                }
                                else
                                {
                                    if (i % 2 == 1)
                                        unitSum.AppendFormat("{0}{1}{2}", (unitSum.Length > 0 ? "," : string.Empty), summaryColList[i], index + SummaryCategoryStartRow);
                                    else
                                        subtotalSum.AppendFormat("{0}{1}{2}", (subtotalSum.Length > 0 ? "," : string.Empty), summaryColList[i], index + SummaryCategoryStartRow);
                                }
                                SetCellValue(cell, string.Empty, 3, (i % 2 == 1 ? intAllSubStyleId : doubleAllSubStyleId));
                            }
                            index++;
                        }
                    }
                    else
                    {
                        //top category units
                        var cell = GetCell(sheet, "A", index + SummaryCategoryStartRow);
                        SetCellValue(cell, "U0", 1, textNullBoldStyleId);
                        cell = GetCell(sheet, "B", index + SummaryCategoryStartRow);
                        SetCellValue(cell, category.CategoryID, 1, textNullBoldStyleId);
                        cell = GetCell(sheet, summaryColList[0], index + SummaryCategoryStartRow);
                        SetCellValue(cell, category.Category + " Total Units", 1, textNullBoldStyleId);
                        SUMUnitsFormat = string.IsNullOrEmpty(SUMUnitsFormat) ? "COLUMN" + (index + SummaryCategoryStartRow).ToString() : SUMUnitsFormat + ", COLUMN" + (index + SummaryCategoryStartRow).ToString();
                        var level2s = from c in CategoryLevelList where c.ParentID == category.CategoryID select c;
                        for (int i = 1; i < summaryColList.Count; i++)
                        {
                            cell = GetCell(sheet, summaryColList[i], index + SummaryCategoryStartRow);
                            CellFormula tformula = new CellFormula();
                            if (i == summaryColList.Count - 1)
                            {
                                //SUM the units of this row
                                tformula.Text = "SUM(" + summaryColList[1] + (index + SummaryCategoryStartRow).ToString() + ":" + summaryColList[i - 1] + (index + SummaryCategoryStartRow).ToString() + ")";
                            }
                            else
                            {
                                tformula.Text = "SUM(" + summaryColList[i] + (index + 1 + SummaryCategoryStartRow).ToString() + ":" + summaryColList[i] + (index + level2s.Count() + SummaryCategoryStartRow).ToString() + ")";
                            }
                            cell.Append(tformula);
                            SetCellValue(cell, string.Empty, 3, intAllMainStyleId);
                        }
                        index++;
                        //level2 category units
                        foreach (CategoryLevelInfo sub in level2s)
                        {
                            cell = GetCell(sheet, "A", index + SummaryCategoryStartRow);
                            SetCellValue(cell, "U", 1, textNullBoldStyleId);
                            cell = GetCell(sheet, "B", index + SummaryCategoryStartRow);
                            SetCellValue(cell, sub.CategoryID, 1, textNullBoldStyleId);
                            cell = GetCell(sheet, summaryColList[0], index + SummaryCategoryStartRow);
                            //SetCellValue(cell, category.Category + " > " + sub.Category, 1, textNullRightBoldStyleId);
                            SetCellValue(cell, sub.Category, 1, textNullRightBoldStyleId);
                            for (int i = 1; i < summaryColList.Count; i++)
                            {
                                cell = GetCell(sheet, summaryColList[i], index + SummaryCategoryStartRow);
                                if (i == summaryColList.Count - 1)
                                {
                                    //SUM the units of this row
                                    CellFormula tformula = new CellFormula();
                                    tformula.Text = "SUM(" + summaryColList[1] + (index + SummaryCategoryStartRow).ToString() + ":" + summaryColList[i - 1] + (index + SummaryCategoryStartRow).ToString() + ")";
                                    cell.Append(tformula);
                                }
                                SetCellValue(cell, string.Empty, 3, intAllSubStyleId);
                            }
                            index++;
                        }
                        //top category value
                        cell = GetCell(sheet, "A", index + SummaryCategoryStartRow);
                        SetCellValue(cell, "V0", 1, textNullBoldStyleId);
                        cell = GetCell(sheet, "B", index + SummaryCategoryStartRow);
                        SetCellValue(cell, category.CategoryID, 1, textNullBoldStyleId);
                        cell = GetCell(sheet, summaryColList[0], index + SummaryCategoryStartRow);
                        SetCellValue(cell, category.Category + " Total Value", 1, textNullBoldStyleId);
                        SUMValueFormat = string.IsNullOrEmpty(SUMValueFormat) ? "COLUMN" + (index + SummaryCategoryStartRow).ToString() : SUMValueFormat + ", COLUMN" + (index + SummaryCategoryStartRow).ToString();
                        for (int i = 1; i < summaryColList.Count; i++)
                        {
                            cell = GetCell(sheet, summaryColList[i], index + SummaryCategoryStartRow);
                            CellFormula tformula = new CellFormula();
                            if (i == summaryColList.Count - 1)
                            {
                                //SUM the value of this row
                                tformula.Text = "SUM(" + summaryColList[1] + (index + SummaryCategoryStartRow).ToString() + ":" + summaryColList[i - 1] + (index + SummaryCategoryStartRow).ToString() + ")";
                            }
                            else
                            {
                                tformula.Text = "SUM(" + summaryColList[i] + (index + 1 + SummaryCategoryStartRow).ToString() + ":" + summaryColList[i] + (index + level2s.Count() + SummaryCategoryStartRow).ToString() + ")";
                            }
                            cell.Append(tformula);
                            SetCellValue(cell, string.Empty, 3, doubleAllSubStyleId);
                        }
                        cell = GetCell(sheet, SummaryCurrencyColumn, index + SummaryCategoryStartRow);
                        SetCellValue(cell, CurrencyCode, 1, CommonTextFormat);
                        index++;
                        //level2 category value
                        foreach (CategoryLevelInfo sub in level2s)
                        {
                            cell = GetCell(sheet, "A", index + SummaryCategoryStartRow);
                            SetCellValue(cell, "V", 1, textNullBoldStyleId);
                            cell = GetCell(sheet, "B", index + SummaryCategoryStartRow);
                            SetCellValue(cell, sub.CategoryID, 1, textNullBoldStyleId);
                            cell = GetCell(sheet, summaryColList[0], index + SummaryCategoryStartRow);
                            //SetCellValue(cell, category.Category + " > " + sub.Category, 1, textNullRightBoldStyleId);
                            SetCellValue(cell, sub.Category, 1, textNullRightBoldStyleId);
                            for (int i = 1; i < summaryColList.Count; i++)
                            {
                                cell = GetCell(sheet, summaryColList[i], index + SummaryCategoryStartRow);
                                if (i == summaryColList.Count - 1)
                                {
                                    //SUM the values of this row
                                    CellFormula tformula = new CellFormula();
                                    tformula.Text = "SUM(" + summaryColList[1] + (index + SummaryCategoryStartRow).ToString() + ":" + summaryColList[i - 1] + (index + SummaryCategoryStartRow).ToString() + ")";
                                    cell.Append(tformula);
                                }
                                SetCellValue(cell, string.Empty, 3, doubleAllSubStyleId);
                            }
                            cell = GetCell(sheet, SummaryCurrencyColumn, index + SummaryCategoryStartRow);
                            SetCellValue(cell, CurrencyCode, 1, CommonTextFormat);
                            index++;
                        }
                    }
                }
                if (CategorySubTotalForStyle == "1")
                {
                    var cell = GetCell(sheet, summaryColList[0], SummaryOrderTotalUnitsRow + index + 2);
                    SetCellValue(cell, "Order Total", 1, textNullBoldStyleId);

                }
                for (int i = 1; i < summaryColList.Count - (CategorySubTotalForStyle == "1"? 0: 1); i++)
                {
                    if (CategorySubTotalForStyle == "1")
                    {
                        var cell = GetCell(sheet, summaryColList[i], SummaryOrderTotalUnitsRow + index + 2);
                        CellFormula tformula = new CellFormula();
                        tformula.Text = "SUM(" + SUMUnitsFormat.Replace("COLUMN", summaryColList[i]) + ")";
                        cell.Append(tformula);
                        SetCellValue(cell, string.Empty, 3, (i % 2 == 1 ? intPinkAllMainStyleId : doublePinkAllSubStyleId));
                    }
                    else
                    {
                        //all total units
                        var cell = GetCell(sheet, summaryColList[i], SummaryOrderTotalUnitsRow);
                        CellFormula tformula = new CellFormula();
                        tformula.Text = "SUM(" + SUMUnitsFormat.Replace("COLUMN", summaryColList[i]) + ")";
                        cell.Append(tformula);
                        SetCellValue(cell, string.Empty, 3, intPinkAllMainStyleId);
                        //all total value
                        cell = GetCell(sheet, summaryColList[i], SummaryOrderTotalValueRow);
                        tformula = new CellFormula();
                        tformula.Text = "SUM(" + SUMValueFormat.Replace("COLUMN", summaryColList[i]) + ")";
                        cell.Append(tformula);
                        SetCellValue(cell, string.Empty, 3, doublePinkAllSubStyleId);
                    }
                }
                sheet.Save();
            }
            book.Close();
        }

        private void SaveCategoryDateValidationToSheet(DataRow catalogRow, string filePath)
        {
            var sheetName = ConfigurationManager.AppSettings["GlobalSheetName"] == null ? "GlobalVariables" : ConfigurationManager.AppSettings["GlobalSheetName"];

            Worksheet sheet = null;
            SpreadsheetDocument book = SpreadsheetDocument.Open(filePath, true);
            var sheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() > 0)
            {
                WorksheetPart worksheetPart = (WorksheetPart)book.WorkbookPart.GetPartById(sheets.First().Id.Value);
                sheet = worksheetPart.Worksheet;
            }
            else
                WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + filePath);
            if (sheet != null)
            {
                Cell cell = null;
                //BOOKING ORDER
                cell = GetCell(sheet, "B", 3);
                var bookingOrder = GetColumnValue(catalogRow, "BookingCatalog");
                SetCellValue(cell, (string.IsNullOrEmpty(bookingOrder) ? "0" : bookingOrder), 1, CommonTextFormat);
                //DateShipStart
                if (!catalogRow.IsNull("DateShipStart"))
                {
                    cell = GetCell(sheet, "B", 4);
                    var date = ((DateTime)catalogRow["DateShipStart"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);//1033
                    SetCellValue(cell, date, 4, CommonDateFormat);
                }
                //DateShipEnd
                if (!catalogRow.IsNull("DateShipEnd"))
                {
                    cell = GetCell(sheet, "B", 5);
                    var date = ((DateTime)catalogRow["DateShipEnd"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);//1033
                    SetCellValue(cell, date, 4, CommonDateFormat);
                }
                //ReqDateDays
                var ReqDateDays = (int)catalogRow["ReqDateDays"];
                cell = GetCell(sheet, "B", 6);
                SetCellValue(cell, ReqDateDays.ToString(), 1, CommonTextFormat);
                //ReqDefaultDays
                var ReqDefaultDays = (int)catalogRow["DefaultReqDays"];
                cell = GetCell(sheet, "B", 7);
                SetCellValue(cell, ReqDefaultDays.ToString(), 1, CommonTextFormat);
                //CancelDateDays
                var CancelDateDays = (int)catalogRow["CancelDateDays"];
                cell = GetCell(sheet, "B", 8);
                SetCellValue(cell, CancelDateDays.ToString(), 1, CommonTextFormat);
                //CancelDefaultDays
                var CancelDefaultDays = (int)catalogRow["CancelDefaultDays"];
                cell = GetCell(sheet, "B", 9);
                SetCellValue(cell, CancelDefaultDays.ToString(), 1, CommonTextFormat);
                //Category Numbers
                cell = GetCell(sheet, "B", 10);
                SetCellValue(cell, CategoryLevelList.Count.ToString(), 1, CommonTextFormat);

                //SKU END ROW
                //cell = GetCell(sheet, "B", 10);
                //SetCellValue(cell, (SKURowNumber + startRow - 1).ToString(), 1, TabularCommonTextFormat);

                //category
                //cell = GetCell(sheet, "A", 11);
                //SetCellValue(cell, "Categories", 1, TabularCommonTextFormat);
                ////set for category configuration
                //for (int i = 0; i < CategoryList.Count; i++)
                //{
                //    cell = GetCell(sheet, "B", 11 + (uint)i);
                //    SetCellValue(cell, CategoryList[i].Category, 1, TabularCommonTextFormat);
                //    cell = GetCell(sheet, "C", 11 + (uint)i);
                //    SetCellValue(cell, CategoryList[i].Column, 1, TabularCommonTextFormat);
                //    cell = GetCell(sheet, "D", 11 + (uint)i);
                //    SetCellValue(cell, CategoryList[i].Row.ToString(), 1, TabularCommonTextFormat);
                //}
            }

            book.WorkbookPart.Workbook.Save();
            book.Close();
        }

        private void SaveCategoryProductsToSheet(string soldto, string catalog, string filePath, string savedirectory)
        {
            var tops = from c in CategoryLevelList where string.IsNullOrEmpty(c.ParentID) select c;
            int i = 1;
            DataTable SKUTable = null;
            foreach (CategoryLevelInfo category in tops)
            {
                InsertWorksheet(category.Category, filePath, i);
                if(CategorySKUATPSheet == "1")
                    InsertWorksheet(category.Category + "_ATP", filePath, i);
                DataTable dt = SaveCategorySKUsToSheet(soldto, catalog, category, filePath, savedirectory, i);

                if (TabularCatalogUPCSheetName.Length > 0)
                {
                    if (SKUTable == null)
                    {
                        SKUTable = dt;
                    }
                    else if (dt != null)
                    {
                        SKUTable.Merge(dt);
                    }
                }
                i++;
            }

            if (TabularCatalogUPCSheetName.Length > 0 && SKUTable != null && SKUTable.Rows.Count > 0)
            {
                DataTable SortedSKUTable = DataTableSort(SKUTable, ConfigurationManager.AppSettings["TabularCatalogUPCSort"] ?? "SheetNameSortOrder ASC,RowIdx ASC");
                uint TabularCatalogUPCStartRowIdx = uint.Parse(ConfigurationManager.AppSettings["TabularCatalogUPCStartRowIdx"] ?? "9");
                Dictionary<string, string> columnList = (ConfigurationManager.AppSettings["TabularCatalogUPCColumnList"] ?? "A|SKU,B|UPC,C|SheetName,D|Style,E|ProductName,F|AttributeValue2,G|PriceWholesale").ToDic(new char[] { '|', ',' });
                Dictionary<string, string> formulacolumnList = (ConfigurationManager.AppSettings["TabularCatalogUPCFormulaColumns"] ?? "I|B{0},J|'{1}'!L{2},L|B{0},M|'{1}'!M{2},O|B{0},P|'{1}'!N{2},R|B{0},S|'{1}'!O{2},U|B{0},V|'{1}'!P{2},X|B{0},Y|'{1}'!Q{2},AA|B{0},AB|J{0}+M{0}+P{0}+S{0}+V{0}+Y{0}").ToDic(new char[] { '|', ',' });

                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
                {
                    Worksheet sheet = null;
                    var sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == TabularCatalogUPCSheetName);
                    if (sheets.Count() > 0)
                    {
                        WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(sheets.First().Id.Value);
                        sheet = worksheetPart.Worksheet;
                    }
                    else
                        WriteToLog("Wrong template order form file, cannot find sheet " + TabularCatalogUPCSheetName + ", file path:" + filePath);

                    if (sheet != null)
                    {
                        var styleSheet = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
                        uint sheetBorder = CreateBorderFormat(styleSheet, false, false, false, true);
                        var doubleAllMainStyleId = CreateCellFormat(styleSheet, null, null, null, numformatid);

                        Cell cCell = null; uint k = 1;
                        foreach (DataRow dr in SortedSKUTable.Rows)
                        {
                            uint rowIdx = (uint)dr["RowIdx"];
                            uint categoryIdx = (uint)dr["SheetNameSortOrder"];
                            if (rowIdx > 0)
                            {
                                foreach (KeyValuePair<string, string> kvp in columnList)
                                {
                                    string cvalue = string.Empty;
                                    if (kvp.Value == "RowSortOrder")
                                    {
                                        cvalue = (categoryIdx * 100000 + rowIdx).ToString();
                                    }
                                    else if (kvp.Value.Contains("+"))
                                    {
                                        string[] arr = kvp.Value.Split(new char[]{'+'});
                                        foreach (string s in arr)
                                            cvalue = cvalue + dr[s].ToString() + " ";
                                    }
                                    else
                                    {
                                        cvalue = dr[kvp.Value].ToString();
                                    }
                                    cCell = GetCell(sheet, kvp.Key, k + TabularCatalogUPCStartRowIdx);
                                    SetCellValue(cCell, cvalue, (new string[] { "RowSortOrder", "PriceWholesale", "RowNo" }).Contains(kvp.Value) ? 2 : 1, doubleAllMainStyleId);
                                }
                                foreach (KeyValuePair<string, string> kvp in formulacolumnList)
                                {
                                    cCell = GetCell(sheet, kvp.Key, k + TabularCatalogUPCStartRowIdx);
                                    CellFormula tformula = new CellFormula();
                                    tformula.Text = string.Format(kvp.Value, (k + TabularCatalogUPCStartRowIdx).ToString(), dr["SheetName"].ToString(), dr["RowIdx"].ToString());
                                    cCell.Append(tformula);
                                }
                                k++;
                            }
                        }
                    }
                }
            }
        }

        private static DataTable DataTableSort(DataTable dt, string sort)
        {
            DataView dv = dt.DefaultView;
            dv.Sort = sort;
            DataTable newdt = dv.ToTable();
            return newdt;
        }
        #region InsertWorksheet
        private void InsertWorksheet(string sheetName, string filePath, int idx)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
            {
                //get template
                WorksheetPart clonedSheet = null;
                var tempSheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == string.Format(CategoryTemplateSheetName, idx));
                if (CategoryTemplateSheetName.Contains("{0}"))
                {
                    Sheet tmp = tempSheets.FirstOrDefault();
                    if (tmp != null)
                    {
                        clonedSheet = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(tmp.Id) as WorksheetPart;  
                        tmp.Name = sheetName;
                        tmp.State = SheetStateValues.Visible;
                    }
                }
                else
                {
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
                //change sheet total titles
                var Locations = CategoryTotalUnitsTitleCell.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                var cell = GetCell(clonedSheet.Worksheet, Locations[0], uint.Parse(Locations[1]));
                var title = GetCellValue(spreadSheet, clonedSheet.Worksheet, cell);
                if (!string.IsNullOrEmpty(title))
                    title = string.Format(title, sheetName);
                SetCellValue(cell, title, 1, CommonBoldTextId);

                Locations = CategoryTotalValueTitleCell.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                cell = GetCell(clonedSheet.Worksheet, Locations[0], uint.Parse(Locations[1]));
                title = GetCellValue(spreadSheet, clonedSheet.Worksheet, cell);
                if (!string.IsNullOrEmpty(title))
                    title = string.Format(title, sheetName);
                SetCellValue(cell, title, 1, CommonBoldTextId);
            }
        }
        #endregion

        private DataTable SaveCategorySKUsToSheet(string soldto, string catalog, CategoryLevelInfo category, string filePath, string savedirectory, int categoryIdx)
        {
            var orderlist = CategorySKUOrderColumnList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            var dt = GetCategoryTabularOrderFormData(soldto, catalog, category.CategoryID, savedirectory);
            dt.Columns.Add("SheetName", typeof(string));
            dt.Columns.Add("SheetNameSortOrder", typeof(uint));
            dt.Columns.Add("RowIdx", typeof(uint));
            int totalSKUs = dt.Rows.Count;
            if (dt != null && totalSKUs > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                { 
                    DataRow dr =dt.Rows[i];
                    dr["SheetName"] = category.Category;
                    dr["SheetNameSortOrder"] = categoryIdx;
                    dr["RowIdx"] = 0;
                }

                int styleSubtotalCount = 0, styleSubtotalIndex = 0;
                if (CategorySubTotalForStyle == "1")
                    styleSubtotalCount = dt.DefaultView.ToTable(true, "Level2DeptID", "Style", "DepartmentID").Rows.Count;

                //category.SKUBeginRow = 0;
                category.SKURowNumber = (uint)totalSKUs;
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
                {
                    Worksheet sheet = null;
                    var sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == category.Category);
                    if (sheets.Count() > 0)
                    {
                        WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(sheets.First().Id.Value);
                        sheet = worksheetPart.Worksheet;
                    }
                    else
                        WriteToLog("Wrong template order form file, cannot find sheet " + category.Category + ", file path:" + filePath);

                    Worksheet atpSheet = null;
                    if (CategorySKUATPSheet == "1")
                    {
                        var atpSheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == category.Category + "_ATP");
                        if (atpSheets.Count() > 0)
                        {
                            WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(atpSheets.First().Id.Value);
                            atpSheet = worksheetPart.Worksheet;
                            atpSheets.First().State = SheetStateValues.Hidden;
                        }
                        else
                            WriteToLog("Wrong template order form file, cannot find ATP sheet " + category.Category + "_ATP" + ", file path:" + filePath);
                    }

                    if (sheet != null)
                    {
                        //Style Format
                        var styleSheet = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
                        //var fontId = CreateFontFormat(styleSheet);
                        //var borderAllId = CreateBorderFormat(styleSheet, true, true, true, true);
                        var borderLRId = CreateBorderFormat(styleSheet, true, true, false, false);
                        var borderLRTId = CreateBorderFormat(styleSheet, true, true, true, false);
                        var borderLRBId = CreateBorderFormat(styleSheet, true, true, false, true);
                        var doubleStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, numformatid); // UInt32Value.FromUInt32(2));
                        var doubleUnlockedStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, numformatid, false, false); // UInt32Value.FromUInt32(2));
                        var doubleTStyleId = CreateCellFormat(styleSheet, CategoryTabTotalFontId.HasValue ? CategoryTabTotalFontId.Value : CommonFontId, CategoryTabTotalFillId, CommonBorderAllId, numformatid); //UInt32Value.FromUInt32(2));
                        var intStyleId = CreateCellFormat(styleSheet, CategoryTabTotalFontId.HasValue ? CategoryTabTotalFontId.Value : CommonFontId, CategoryTabTotalFillId, CommonBorderAllId, intformatid); //UInt32Value.FromUInt32(1));
                        var intLStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, intformatid, false, false); //UInt32Value.FromUInt32(1), false, false);
                        var textAllStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, UInt32Value.FromUInt32(49));
                        var textLRStyleId = CreateCellFormat(styleSheet, CommonFontId, null, borderLRId, UInt32Value.FromUInt32(49));
                        var intLockedLStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid2, CommonBorderAllId, intformatid, true, false);
                        var textLRTStyleId = CreateCellFormat(styleSheet, CommonFontId, null, borderLRTId, UInt32Value.FromUInt32(49));
                        var textLRBStyleId = CreateCellFormat(styleSheet, CommonFontId, null, borderLRBId, UInt32Value.FromUInt32(49));
                        var textNullStyleId = CreateCellFormat(styleSheet, CommonFontId, null, null, UInt32Value.FromUInt32(49));
                        var textDAllStyleId = CreateCellFormat(styleSheet, CommonFontId, null, CommonBorderAllId, UInt32Value.FromUInt32(49), true, true);
                        var textDLRTStyleId = CreateCellFormat(styleSheet, CommonFontId, null, borderLRTId, UInt32Value.FromUInt32(49), true, true);
                        var textNullBlackStyleId = CreateCellFormat(styleSheet, CommonFontId, blackFillID, null, UInt32Value.FromUInt32(49));
                        var textAllSubStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid2, CommonBorderAllId, UInt32Value.FromUInt32(49));
                        var doubleAllSubStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid2, CommonBorderAllId, numformatid);
                        var intAllSubStyleId = CreateCellFormat(styleSheet, bolderfont, bgfillid2, CommonBorderAllId, intformatid);
                        var doubleSubTotalStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid1, CommonBorderAllId, numformatid);
                        var intSubTotalStyleId = CreateCellFormat(styleSheet, CommonFontId, bgfillid1, CommonBorderAllId, intformatid);
                        //Check previous departmentid, style, color
                        var preCategoryId = string.Empty;
                        var preDept = string.Empty;
                        var preStyle = string.Empty;
                        var preColor = string.Empty;
                        var multipleList = new List<OrderMultipleDataValidation>();

                        Cell cCell = null;
                        var currencys = CategotySKUCurrencyLocations.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        foreach (var currency in currencys)
                        {
                            var location = currency.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                            cCell = GetCell(sheet, location[0], uint.Parse(location[1]));
                            SetCellValue(cCell, CurrencyCode, 1, textAllStyleId);
                        }
                        //total units
                        var totalunits = CategorySKUTotalUnitsLocation.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        cCell = GetCell(sheet, totalunits[0], uint.Parse(totalunits[1]));
                        CellFormula tformula = new CellFormula();
                        tformula.Text = "SUM(" + CategorySKUUnitTotalColumn + startRow.ToString() + ":" + CategorySKUUnitTotalColumn + (totalSKUs + startRow + (uint)styleSubtotalCount - 1).ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                        cCell.Append(tformula);
                        SetCellValue(cCell, string.Empty, 3, intStyleId);
                        //total value for order1-6
                        for (int o = 0; o < orderlist.Count; o++)
                        {
                            var endrow = totalSKUs + startRow + styleSubtotalCount - 1;
                            if (totalSKUs <= 1)
                                endrow++;
                            cCell = GetCell(sheet, orderlist[o], CategorySKUSubTotalRow);
                            tformula = new CellFormula();
                            tformula.Text = "SUMPRODUCT(" + CategorySKUPriceColumn + startRow.ToString() + ":" + CategorySKUPriceColumn + (endrow).ToString() + "," + orderlist[o] + startRow.ToString() + ":" + orderlist[o] + (endrow).ToString() + ")";
                            cCell.Append(tformula);
                            SetCellValue(cCell, string.Empty, 3, doubleTStyleId);
                        }

                        if (!CategorySKUComColumnList.Contains(CategorySKUPriceColumn + "|PriceWholesale"))
                            CategorySKUComColumnList = CategorySKUComColumnList + "," + CategorySKUPriceColumn + "|PriceWholesale";
                        var comColumnList = CategorySKUComColumnList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        uint currentStyleStartRows = startRow;
                        for (int i = 0; i < totalSKUs + (CategorySubTotalForStyle == "1" ? 1 : 0); i++) // +1 to Add the Last style-level subtotal line
                        {
                            Cell cell = null;
                            var row = (i == totalSKUs ? null : dt.Rows[i]);
                            var categoryId = GetColumnValue(row, "Level2DeptID");
                            var deptId = GetColumnValue(row, "DepartmentID");
                            var style = GetColumnValue(row, "Style");
                            var color = GetColumnValue(row, "AttributeValue2");

                            if (styleSubtotalCount > 0 && preStyle.Length > 0 && (style != preStyle || i == totalSKUs))  // Add style-level subtotal line
                            {
                                uint currentStyleEndRows = (uint)i + startRow + (uint)styleSubtotalIndex - 1;
                                foreach (string column in comColumnList)
                                {
                                    LoadTabularDataToCell(column, sheet, (uint)i + startRow + (uint)styleSubtotalIndex, null, preDept, preDept, "", "", "", "", i, totalSKUs, styleSubtotalCount,
                                        doubleSubTotalStyleId, doubleSubTotalStyleId, textDLRTStyleId, textLRBStyleId, textLRStyleId, textLRTStyleId, category.Category);
                                }
                                for (int o = 0; o < orderlist.Count; o++)
                                {
                                    cell = GetCell(sheet, orderlist[o], (uint)i + startRow + (uint)styleSubtotalIndex);
                                    CellFormula cellf = new CellFormula();
                                    cellf.Text = "SUM(" + orderlist[o] + currentStyleStartRows.ToString() + ":" + orderlist[o] + currentStyleEndRows.ToString() + ")";
                                    cell.Append(cellf);
                                    SetCellValue(cell, string.Empty, 5, intSubTotalStyleId);
                                }
                                //total unite funcational
                                cell = GetCell(sheet, CategorySKUUnitTotalColumn, (uint)i + startRow + (uint)styleSubtotalIndex);
                                CellFormula cellformula1 = new CellFormula();
                                cellformula1.Text = "SUM(" + CategorySKUUnitTotalColumn + currentStyleStartRows.ToString() + ":" + CategorySKUUnitTotalColumn + currentStyleEndRows.ToString() + ")";
                                cell.Append(cellformula1);
                                cell.CellValue = new CellValue(string.Empty);
                                SetCellValue(cell, string.Empty, 3, intSubTotalStyleId);

                                //total value funcational
                                cell = GetCell(sheet, CategorySKUValueTotalColumn, (uint)i + startRow + (uint)styleSubtotalIndex);
                                CellFormula cellformulb1 = new CellFormula();
                                cellformulb1.Text = "SUM(" + CategorySKUValueTotalColumn + currentStyleStartRows.ToString() + ":" + CategorySKUValueTotalColumn + currentStyleEndRows.ToString() + ")";
                                cell.Append(cellformulb1);
                                cell.CellValue = new CellValue(string.Empty);
                                SetCellValue(cell, string.Empty, 3, doubleSubTotalStyleId);

                                // Hide the Subtotal Line if only One SKU for a Style
                                if (currentStyleStartRows == currentStyleEndRows)
                                {
                                    Row subtotalRow = GetRow(sheet, currentStyleEndRows + 1);
                                    subtotalRow.CustomHeight = true;
                                    subtotalRow.Height = 0;
                                }

                                styleSubtotalIndex++;

                                currentStyleStartRows = (uint)i + startRow + (uint)styleSubtotalIndex;

                                if (i == totalSKUs)  // the last subtotal line
                                {
                                    break;
                                }
                            }

                            foreach (string column in comColumnList)
                            {
                                if (column == CategorySKUPriceColumn + "|PriceWholesale")
                                {
                                    uint rowidx = (uint)i + startRow + (uint)styleSubtotalIndex;
                                    cell = GetCell(sheet, CategorySKUPriceColumn, rowidx);
                                    SetCellValue(cell, GetColumnValue(row, "PriceWholesale"), 2, CategoryWSPriceEditable ? doubleUnlockedStyleId : doubleStyleId);

                                    row["RowIdx"] = rowidx;
                                }
                                else
                                {
                                    LoadTabularDataToCell(column, sheet, (uint)i + startRow + (uint)styleSubtotalIndex, row, deptId, preDept, style, preStyle, color, preColor, i, totalSKUs, styleSubtotalCount,
                                        textAllStyleId, textDAllStyleId, textDLRTStyleId, textLRBStyleId, textLRStyleId, textLRTStyleId, category.Category);
                                }
                            }
                            if (categoryId != preCategoryId)
                            {
                                var preCategorys = from c in CategoryLevelList where c.CategoryID == preCategoryId && c.ParentID == category.CategoryID select c;
                                if (preCategorys.Count() > 0)
                                {
                                    var presubc = preCategorys.First();
                                    presubc.SKUEndRow = startRow + (uint)i + (uint)styleSubtotalIndex - 1;
                                }
                                var categorys = from c in CategoryLevelList where c.CategoryID == categoryId && c.ParentID == category.CategoryID select c;
                                if (categorys.Count() > 0)
                                {
                                    var subc = categorys.First();
                                    subc.SKUBeginRow = startRow + (uint)i + (uint)styleSubtotalIndex;
                                }
                            }
                            preCategoryId = categoryId;
                            preDept = deptId;
                            preStyle = style;
                            preColor = color;

                            for (int o = 0; o < orderlist.Count; o++)
                            {
                                cell = GetCell(sheet, orderlist[o], (uint)i + startRow + (uint)styleSubtotalIndex);
                                if (ShipDateRequiredForStart == "1")
                                    SetCellValue(cell, string.Empty, 3, intLockedLStyleId);
                                else
                                    SetCellValue(cell, string.Empty, 3, intLStyleId);
                            }

                            //total unite funcational
                            cell = GetCell(sheet, CategorySKUUnitTotalColumn, (uint)i + startRow + (uint)styleSubtotalIndex);
                            CellFormula cellformula = new CellFormula();
                            cellformula.Text = "SUM(" + orderlist[0] + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + orderlist[orderlist.Count - 1] + (i + startRow + (uint)styleSubtotalIndex).ToString() + ")";
                            cell.Append(cellformula);
                            cell.CellValue = new CellValue(string.Empty);
                            SetCellValue(cell, string.Empty, 3, intStyleId);

                            //total value funcational
                            cell = GetCell(sheet, CategorySKUValueTotalColumn, (uint)i + startRow + (uint)styleSubtotalIndex);
                            CellFormula cellformulb = new CellFormula();
                            cellformulb.Text = "=SUMPRODUCT(" + CategorySKUPriceColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + CategorySKUPriceColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + "," + CategorySKUUnitTotalColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + CategorySKUUnitTotalColumn + (i + startRow + (uint)styleSubtotalIndex).ToString() + ")";
                            cell.Append(cellformulb);
                            cell.CellValue = new CellValue(string.Empty);
                            SetCellValue(cell, string.Empty, 3, doubleTStyleId);

                            //ORDER MULTIPLE
                            if (OrderMultipleCheck == "1")
                            {
                                var OrderMultipleStr = GetColumnValue(row, "OrderMultiple");
                                decimal OrderMultipleD = 0;
                                try
                                {
                                    decimal.TryParse(OrderMultipleStr, out OrderMultipleD);
                                }
                                catch { OrderMultipleD = 1; }
                                int OrderMultiple = OrderMultipleD <= 0 ? 1 : (int)OrderMultipleD;
                                OrderMultipleDataValidation multiple = null;
                                var multips = from m in multipleList where m.Multiple == OrderMultiple select m;
                                if (multips.Count() <= 0)
                                {
                                    multiple = new OrderMultipleDataValidation { Multiple = OrderMultiple, SequenceOfReferences = orderlist[0] + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + orderlist[orderlist.Count - 1] + (i + startRow + (uint)styleSubtotalIndex).ToString() };
                                    multipleList.Add(multiple);
                                }
                                else
                                {
                                    multiple = multips.First();
                                    multiple.SequenceOfReferences = multiple.SequenceOfReferences + " " + orderlist[0] + (i + startRow + (uint)styleSubtotalIndex).ToString() + ":" + orderlist[orderlist.Count - 1] + (i + startRow + (uint)styleSubtotalIndex).ToString();
                                }
                            }

                            if (atpSheet != null && row != null)
                            {
                                var atpDate = row.IsNull("ATPDateToCompare") ? string.Empty : row["ATPDateToCompare"].ToString();
                                var atpCell = GetCell(atpSheet, "A", (uint)i + startRow + (uint)styleSubtotalIndex);
                                SetCellValue(atpCell, atpDate, 1, textNullStyleId);
                            }
                        }
                        //Footer summary
                        var summaryCols = CategorySKUSummaryColumnList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        var footerBegin = totalSKUs + startRow + (uint)styleSubtotalCount;
                        foreach (var col in summaryCols)
                        {
                            cCell = GetCell(sheet, col, (uint)footerBegin);
                            SetCellValue(cCell, string.Empty, 3, textNullBlackStyleId);
                        }
                        var level2s = from c in CategoryLevelList where c.ParentID == category.CategoryID select c;
                        for (int c = 0; c < level2s.Count(); c++)
                        {
                            level2s.ElementAt(c).CategoryRow = (uint)(footerBegin + c + 1);
                            if (level2s.ElementAt(c).SKUEndRow <= 0)
                                level2s.ElementAt(c).SKUEndRow = (uint)(totalSKUs + startRow + (uint)styleSubtotalCount - 1);
                            cCell = null;
                            foreach (var col in summaryCols)
                            {
                                if (col == CategorySKUSummaryCategoryColumn || col == CategorySKUSummaryCategoryIDColumn ||
                                    col == CategorySKUUnitTotalColumn || col == CategorySKUValueTotalColumn ||
                                    orderlist.Contains(col))
                                {
                                    if (col == CategorySKUSummaryCategoryColumn)
                                    {
                                        cCell = GetCell(sheet, col, (uint)(footerBegin + c + 1));
                                        SetCellValue(cCell, level2s.ElementAt(c).Category, 1, textAllSubStyleId);
                                    }
                                    if (col == CategorySKUSummaryCategoryIDColumn)
                                    {
                                        cCell = GetCell(sheet, CategorySKUSummaryCategoryIDColumn, (uint)(footerBegin + c + 1));
                                        SetCellValue(cCell, level2s.ElementAt(c).CategoryID, 1, textAllStyleId);
                                    }
                                    if (orderlist.Contains(col) || col == CategorySKUUnitTotalColumn)
                                    {
                                        cCell = GetCell(sheet, col, (uint)(footerBegin + c + 1));
                                        CellFormula cellformula = new CellFormula();
                                        cellformula.Text = "SUM(" + col + level2s.ElementAt(c).SKUBeginRow.ToString() + ":" + col + level2s.ElementAt(c).SKUEndRow.ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                                        cCell.Append(cellformula);
                                        SetCellValue(cCell, string.Empty, 3, intAllSubStyleId);
                                    }
                                    if (col == CategorySKUValueTotalColumn)
                                    {
                                        cCell = GetCell(sheet, col, (uint)(footerBegin + c + 1));
                                        CellFormula cellformula = new CellFormula();
                                        cellformula.Text = "SUM(" + col + level2s.ElementAt(c).SKUBeginRow.ToString() + ":" + col + level2s.ElementAt(c).SKUEndRow.ToString() + ")" + (CategorySubTotalForStyle == "1" ? "/2" : string.Empty);
                                        cCell.Append(cellformula);
                                        SetCellValue(cCell, string.Empty, 3, doubleAllSubStyleId);
                                    }
                                }
                                else
                                {
                                    cCell = GetCell(sheet, col, (uint)(footerBegin + c + 1));
                                    SetCellValue(cCell, string.Empty, 3, textAllSubStyleId);
                                }
                            }
                        }
                        //Update Numeric Validation
                        DataValidation validation = null;
                        if (OrderMultipleCheck != "1")
                        {
                            var validations = from v in sheet.Descendants<DataValidation>() where v.SequenceOfReferences.InnerText == TabularNumericValidationCell select v;
                            validation = validations.Count() > 0 ? validations.First() : null;
                            if (validation != null)
                            {
                                validation.SequenceOfReferences.InnerText = orderlist[0] + startRow.ToString() + ":" + orderlist[orderlist.Count - 1] + (startRow + totalSKUs + (uint)styleSubtotalCount - 1).ToString();
                            }
                        }
                        else
                        {
                            var validations = sheet.Descendants<DataValidations>().Count() > 0 ? sheet.Descendants<DataValidations>().First() : null;
                            if (validations != null)
                            {
                                multipleList.ForEach(m =>
                                {
                                    validation = new DataValidation
                                    {
                                        Type = DataValidationValues.Custom,
                                        AllowBlank = true,
                                        ShowInputMessage = true,
                                        ShowErrorMessage = true,
                                        ErrorTitle = "Multiple Value Cell",
                                        Error = "You must enter a multiple of " + m.Multiple.ToString() + " in this cell.",
                                        SequenceOfReferences = new ListValue<StringValue> { InnerText = m.SequenceOfReferences },
                                        Formula1 = new Formula1 { Text = "(MOD(INDIRECT(ADDRESS(ROW(),COLUMN()))," + m.Multiple.ToString() + ")=0)" }
                                    };
                                    validations.Append(validation);
                                });
                            }
                        }

                        if (!string.IsNullOrEmpty(CategorySKUAutoFilterRange))
                        {
                            var sheetData = sheet.Descendants<SheetData>().Count() > 0 ? sheet.Descendants<SheetData>().First() : null;
                            if (sheetData != null)
                            {
                                AutoFilter autoFilter = new AutoFilter() { Reference = CategorySKUAutoFilterRange };
                                sheet.InsertAfter(autoFilter, sheetData);
                            }
                        }

                        sheet.Save();
                    }
                }
            }
            return dt;
        }

        #region GetCategoryTabularOrderFormData [p_Offline_CategoryTabularGridView]
        private DataTable GetCategoryTabularOrderFormData(string soldto, string catalog, string Category, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_CategoryTabularGridView", conn);
                cmd.CommandTimeout = 0;
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
                cmd.Parameters["@catcd"].Value = catalog;

                cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
                cmd.Parameters["@SoldTo"].Value = soldto;

                cmd.Parameters.Add(new SqlParameter("@Category", SqlDbType.VarChar, 80));
                cmd.Parameters["@Category"].Value = Category;

                cmd.Parameters.Add(new SqlParameter("@PriceType", SqlDbType.VarChar, 80));
                cmd.Parameters["@PriceType"].Value = string.IsNullOrEmpty(PriceType) ? "W" : PriceType;

                try
                {
                    DataSet outds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(outds);
                    conn.Close();
                    if (outds.Tables.Count > 0)
                    {
                        return outds.Tables[0];
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    //throw;
                    WriteToLog("p_Offline_TabularGridView | GetTabularOrderFormData |" + soldto + "|" + catalog + "|" + Category, ex, savedirectory);
                    return null;
                }
            }
        }
        #endregion

        private void LoadTabularDataToCell(string columnInfo, Worksheet sheet, uint rowIndex, DataRow row,
            string deptId, string preDept, string style, string preStyle, string color, string preColor, int loopIndex, int totalCount, int subtotalCount,
            UInt32Value textAllStyleId, UInt32Value textDAllStyleId, UInt32Value textDLRTStyleId, UInt32Value textLRBStyleId, UInt32Value textLRStyleId, UInt32Value textLRTStyleId, string topCategory)
        {
            Cell cell = null;
            var colInfo = columnInfo.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            cell = GetCell(sheet, colInfo[0], rowIndex);
            switch (colInfo[1])
            {
                case "Catalog":
                    string catalogcd = GetColumnValue(row, "CatalogCode");
                    if (catalogcd == string.Empty)
                        SetCellValue(cell, "*", 1, textAllStyleId);
                    else
                        SetCellValue(cell, GetColumnValue(row, "CatalogCode"), 1, textAllStyleId);
                    break;
                case "Department":
                    var NavigationBar = GetColumnValue(row, "NavigationBar");
                    Regex regEx = new Regex(topCategory + " > ", RegexOptions.Multiline);
                    NavigationBar = regEx.Replace(NavigationBar, "", 1);
                    if (CategorySKUCombineDepartment == "1")
                    {
                        if (deptId != preDept)
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                            {
                                SetCellValue(cell, NavigationBar, 1, textDAllStyleId);
                            }
                            else
                            {
                                SetCellValue(cell, NavigationBar, 1, textDLRTStyleId);//DepartmentName
                            }
                        }
                        else
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                                SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                            else
                                SetCellValue(cell, string.Empty, 1, textLRStyleId);
                        }
                    }
                    else
                        SetCellValue(cell, NavigationBar, 1, textDAllStyleId);
                    break;
                case "Style":
                    if (CategorySKUCombineStyle == "1")
                    {
                        if ((style != preStyle) || (style == preStyle && deptId != preDept))
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                                SetCellValue(cell, style, 1, textAllStyleId);
                            else
                                SetCellValue(cell, style, 1, textLRTStyleId);//DepartmentName
                        }
                        else
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                                SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                            else
                                SetCellValue(cell, string.Empty, 1, textLRStyleId);
                        }
                    }
                    else
                        SetCellValue(cell, style, 1, textAllStyleId);
                    break;
                case "ProductName":
                    var productName = GetColumnValue(row, "ProductName");
                    if (CategorySKUCombineProductName == "1")
                    {
                        if ((style != preStyle) || (style == preStyle && deptId != preDept))
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                                SetCellValue(cell, productName, 1, textAllStyleId);
                            else
                                SetCellValue(cell, productName, 1, textLRTStyleId);//DepartmentName
                        }
                        else
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                                SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                            else
                                SetCellValue(cell, string.Empty, 1, textLRStyleId);
                        }
                    }
                    else
                        SetCellValue(cell, productName, 1, textDAllStyleId); //textAllStyleId);
                    break;
                case "AttributeValue2":
                    if (CategorySKUCombineColor == "1")
                    {
                        if ((color != preColor) || (color == preColor && deptId != preDept) || (color == preColor && style != preStyle))
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                                SetCellValue(cell, color, 1, textAllStyleId);
                            else
                                SetCellValue(cell, color, 1, textLRTStyleId);//DepartmentName
                        }
                        else
                        {
                            if (loopIndex == totalCount + subtotalCount - 1)
                                SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                            else
                                SetCellValue(cell, string.Empty, 1, textLRStyleId);
                        }
                    }
                    else
                        SetCellValue(cell, color, 1, textAllStyleId);
                    break;
                case "Level2DeptID":
                case "SKU":
                case "UPC":
                case "AttributeValue1":
                case "AttributeValue5":
                case "Gender":
                default:
                    SetCellValue(cell, GetColumnValue(row, colInfo[1]), 1, textAllStyleId);
                    break;
            }
        }

        #region SaveCategoryListToSheet
        private void SaveCategoryListToSheet(string filePath)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
            {
                Worksheet sheet = null;
                var sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == CategoryListSheetName);
                if (sheets.Count() > 0)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(sheets.First().Id.Value);
                    sheet = worksheetPart.Worksheet;
                }
                else
                    WriteToLog("Wrong template order form file, cannot find sheet " + CategoryListSheetName + ", file path:" + filePath);

                if (sheet != null)
                {
                    uint index = 2;
                    foreach (CategoryLevelInfo category in CategoryLevelList)
                    {
                        var cell = GetCell(sheet, "A", index);
                        SetCellValue(cell, category.CategoryID, 1, CommonTextFormat);
                        cell = GetCell(sheet, "B", index);
                        SetCellValue(cell, category.Category, 1, CommonTextFormat);
                        cell = GetCell(sheet, "C", index);
                        SetCellValue(cell, category.ParentID, 1, CommonTextFormat);
                        cell = GetCell(sheet, "D", index);
                        SetCellValue(cell, category.ParentCategory, 1, CommonTextFormat);
                        cell = GetCell(sheet, "E", index);
                        SetCellValue(cell, category.CategoryRow.ToString(), 1, CommonTextFormat);
                        cell = GetCell(sheet, "F", index);
                        SetCellValue(cell, category.SKUBeginRow.ToString(), 1, CommonTextFormat);
                        cell = GetCell(sheet, "G", index);
                        SetCellValue(cell, category.SKUEndRow.ToString(), 1, CommonTextFormat);
                        cell = GetCell(sheet, "H", index);
                        SetCellValue(cell, category.SKURowNumber.ToString(), 1, CommonTextFormat);
                        cell = GetCell(sheet, "I", index);
                        SetCellValue(cell, category.IsTopLevel ? "1" : "0", 1, CommonTextFormat);
                        index++;
                    }
                    sheet.Save();
                }
            }
        }
        #endregion      
        
        private void ProtectWorkbook(string filename, string savedirectory)
        {
            var pwd = ConfigurationManager.AppSettings["ProtectPassword"] == null ? "Plumriver" : ConfigurationManager.AppSettings["ProtectPassword"];
            var password = HashPassword(pwd);
            var filePath = Path.Combine(savedirectory, filename);
            SpreadsheetDocument book = SpreadsheetDocument.Open(filePath, true);
            Worksheet sheet = null;

            var tops = from c in CategoryLevelList where string.IsNullOrEmpty(c.ParentID) select c;
            foreach (CategoryLevelInfo category in tops)
            {
                sheet = null;
                var sheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == category.Category);
                if (sheets.Count() > 0)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)book.WorkbookPart.GetPartById(sheets.First().Id.Value);
                    sheet = worksheetPart.Worksheet;
                }
                if (sheet != null)
                {
                    bool addNew = false;
                    SheetProtection sheetProtection = null;
                    sheetProtection = sheet.Descendants<SheetProtection>().Count() > 0 ? sheet.Descendants<SheetProtection>().First() : null;
                    if (sheetProtection == null)
                    {
                        sheetProtection = new SheetProtection() { Sheet = true, Objects = false, Scenarios = false, AutoFilter = false }; //, DeleteRows = false, DeleteColumns = false, FormatCells = false, FormatColumns = false, FormatRows = false, InsertColumns = false, InsertRows = false, SelectLockedCells = false};
                        addNew = true;
                    }
                    else
                    {
                        sheetProtection.Sheet = true;
                        sheetProtection.Objects = false;
                        sheetProtection.Scenarios = false;
                        sheetProtection.AutoFilter = false;
                    }
                    //{ Sheet = true, Objects = true, Scenarios = true };
                    sheetProtection.Password = password;
                    if (addNew)
                        sheet.InsertAfter(sheetProtection, sheet.Descendants<SheetData>().LastOrDefault());
                }
            }

            book.WorkbookPart.Workbook.WorkbookProtection = new WorkbookProtection() { LockStructure = true };
            ////book.WorkbookPart.Workbook.WorkbookProtection.LockWindows = true;
            //book.WorkbookPart.Workbook.WorkbookProtection.WorkbookAlgorithmName = "SHA-1";
            //book.WorkbookPart.Workbook.WorkbookProtection.WorkbookPassword = password;//new HexBinaryValue() { Value = password };

            var gsheetName = ConfigurationManager.AppSettings["GlobalSheetName"] == null ? "GlobalVariables" : ConfigurationManager.AppSettings["GlobalSheetName"];
            var pwdColumn = ConfigurationManager.AppSettings["GlobalPWDColumn"] == null ? "B" : ConfigurationManager.AppSettings["GlobalPWDColumn"];
            var pwdRow = ConfigurationManager.AppSettings["GlobalPWDRow"] == null ? 200 : int.Parse(ConfigurationManager.AppSettings["GlobalPWDRow"]);

            sheet = null;
            var gsheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == gsheetName);
            if (gsheets.Count() > 0)
            {
                WorksheetPart worksheetPart = (WorksheetPart)book.WorkbookPart.GetPartById(gsheets.First().Id.Value);
                sheet = worksheetPart.Worksheet;
            }
            else
                WriteToLog("Wrong template order form file, cannot find sheet " + gsheetName + ", file path:" + filePath);
            if (sheet != null)
            {
                var cell = GetCell(sheet, pwdColumn, (uint)(pwdRow));
                SetCellValue(cell, pwd, 1, CommonTextFormat);
            }

            book.WorkbookPart.Workbook.Save();
            book.Close();
        }
        */




        /*************************SUPPORT METHODS**************************/
        /*
        private Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            Cell cell = null;
            var cells = row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() <= 0)
            {
                cell = new Cell() { CellReference = columnName + rowIndex };
                row.Append(cell);
            }
            else
                cell = cells.First();
            return cell;
        }

        private Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            Row row = null;
            var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex);
            if (rows.Count() <= 0)
            {
                var sheetData = worksheet.GetFirstChild<SheetData>();
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            else
                row = rows.First();
            return row;
        }

        private UInt32Value CreateFontFormat(Stylesheet styleSheet)
        {
            Font newFont = new Font();
            FontSize newFontSize = new FontSize() { Val = 9 };
            FontName fontName = new FontName() { Val = "Arial" };
            FontFamily fontFamily = new FontFamily() { Val = 2 };
            newFont.Append(newFontSize);
            newFont.Append(fontName);
            newFont.Append(fontFamily);
            styleSheet.Fonts.Append(newFont);

            UInt32Value result = styleSheet.Fonts.Count;
            styleSheet.Fonts.Count++;
            return result;
        }

        private UInt32Value CreateBorderFormat(Stylesheet styleSheet, bool left, bool right, bool top, bool bottom)
        {
            var borderContent = new Border();
            if (left)
                borderContent.LeftBorder = new LeftBorder(new Color() { Indexed = 64 }) { Style = BorderStyleValues.Thin };//Auto = true
            if (right)
                borderContent.RightBorder = new RightBorder(new Color() { Indexed = 64 }) { Style = BorderStyleValues.Thin };
            if (top)
                borderContent.TopBorder = new TopBorder(new Color() { Indexed = 64 }) { Style = BorderStyleValues.Thin };
            if (bottom)
                borderContent.BottomBorder = new BottomBorder(new Color() { Indexed = 64 }) { Style = BorderStyleValues.Thin };
            borderContent.DiagonalBorder = new DiagonalBorder() { Style = BorderStyleValues.Thin };
            styleSheet.Borders.Append(borderContent);
            UInt32Value result = styleSheet.Borders.Count;
            styleSheet.Borders.Count++;
            return result;
        }

        private UInt32Value CreateCellFormat(Stylesheet styleSheet, UInt32Value fontIndex, UInt32Value fillIndex, UInt32Value borderIndex, UInt32Value numberFormatId)
        {
            return CreateCellFormat(styleSheet, fontIndex, fillIndex, borderIndex, numberFormatId, true, false);
        }

        private UInt32Value CreateCellFormat(Stylesheet styleSheet, UInt32Value fontIndex, UInt32Value fillIndex, UInt32Value borderIndex, UInt32Value numberFormatId, bool locked, bool alignment)
        {
            return CreateCellFormat(styleSheet, fontIndex, fillIndex, borderIndex, numberFormatId, locked, alignment, HorizontalAlignmentValues.General);
        }

        private UInt32Value CreateCellFormat(Stylesheet styleSheet, UInt32Value fontIndex, UInt32Value fillIndex, UInt32Value borderIndex, UInt32Value numberFormatId, bool locked, bool alignment, HorizontalAlignmentValues halignment)
        {
            CellFormat cellFormat = new CellFormat();

            if (fontIndex != null)
            {
                cellFormat.FontId = fontIndex;
                cellFormat.ApplyFont = true;
            }

            if (fillIndex != null)
            {
                cellFormat.FillId = fillIndex;
                cellFormat.ApplyFill = true;
            }

            if (borderIndex != null)
            {
                cellFormat.BorderId = borderIndex;
                cellFormat.ApplyBorder = true;
            }

            if (numberFormatId != null)
            {
                cellFormat.NumberFormatId = numberFormatId;
                cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            }

            cellFormat.ApplyProtection = true;
            if (!locked)
            {
                cellFormat.Protection = new Protection();
                cellFormat.Protection.Locked = new BooleanValue(locked);
            }

            cellFormat.ApplyAlignment = true;
            cellFormat.Alignment = new Alignment { Horizontal = halignment };
            if (alignment)
            {
                cellFormat.Alignment.WrapText = true;
            }
            styleSheet.CellFormats.Append(cellFormat);

            UInt32Value result = styleSheet.CellFormats.Count;
            styleSheet.CellFormats.Count++;
            return result;
        }

        private void SetCellValue(Cell cell, string value, int format)
        {
            SetCellValue(cell, value, format, 999999);
        }

        private void SetCellValue(Cell cell, string value, int format, UInt32Value styleId)
        {
            switch (format)
            {
                case 1:
                    if (styleId != 999999) cell.StyleIndex = styleId;
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    break;
                case 2:
                    if (styleId != 999999) cell.StyleIndex = styleId;
                    cell.CellValue = new CellValue(string.Format("{0:N2}", value));
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case 3:
                    if (styleId != 999999) cell.StyleIndex = styleId;
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case 4:
                    if (styleId != 999999) cell.StyleIndex = styleId;
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    break;
                case 5:
                    if (styleId != 999999) cell.StyleIndex = styleId;
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);

                    break;
            }
        }

        private string GetColumnValue(DataRow row, string ColumnName)
        {
            if (row == null)
            {
                return string.Empty;
            }
            else if (row.Table.Columns.Contains(ColumnName))
            {
                if (row.IsNull(ColumnName))
                    return string.Empty;
                return row[ColumnName].ToString();
            }
            return string.Empty;
        }

        private static double GetWidth(string text)
        {
            System.Drawing.Font stringFont = new System.Drawing.Font("Arial", 9);
            System.Drawing.Size textSize = System.Windows.Forms.TextRenderer.MeasureText(text, stringFont);
            double width = (double)(((textSize.Width / (double)7) * 256) - (128 / 7)) / 256;
            width = (double)decimal.Round((decimal)width + 0.2M, 2);

            return width;
        }

        private string GetCellValue(SpreadsheetDocument book, Worksheet worksheet, Cell cell)
        {
            string cellValue = string.Empty;
            if (cell != null)
            {
                if (cell.DataType == null)
                    cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.Text;
                else
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.Number:
                            cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.Text;
                            break;
                        default:
                            if (cell.CellValue != null)
                            {
                                var id = int.Parse(cell.CellValue.Text);
                                cellValue = book.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id).InnerText;
                            }
                            break;
                    }
                }
            }
            return cellValue;
        }

        public String GetSheetPassword(String password)
        {
            Int32 pLength = password.Length;
            Int32 hash = 0;
            if (pLength == 0) return hash.ToString("X");

            for (Int32 i = pLength - 1; i >= 0; i--)
            {
                hash = hash >> 14 & 0x01 | hash << 1 & 0x7fff;
                hash ^= password[i];
            }
            hash = hash >> 14 & 0x01 | hash << 1 & 0x7fff;
            hash ^= 0x8000 | 'N' << 8 | 'K';
            hash ^= pLength;
            return hash.ToString("X");
        }

        protected string HashPassword(string password)
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
        }*/


        private void WriteToLog(string msg, Exception e, string savedirectory)
        {
            var logFile = "OfflineLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            StreamWriter sw = new StreamWriter(@savedirectory + logFile, true);
            sw.WriteLine("At " + DateTime.Now.Date + ", " + DateTime.Now.TimeOfDay + ":");
            sw.WriteLine(e.Message);
            sw.WriteLine(e.StackTrace);
            if (e.InnerException != null)
                sw.WriteLine(e.InnerException.Message);
            sw.WriteLine(msg);
            sw.WriteLine();
            sw.Close();
        }

        private static void WriteToLog(string msg)
        {
            //string templatefile = ConfigurationManager.AppSettings["templatefile"].ToString();
            var logFile = "OfflineLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            string savefolder = ConfigurationManager.AppSettings["savefolder"].ToString();
            StreamWriter sw = new StreamWriter(savefolder + logFile, true);
            sw.WriteLine(msg);
            sw.Close();
        }
        


        #region copy sheet DO NOT NEED NOW CANNOT COPY MACRO WHEN COPY SHEET WITH OPENXML
        /**
        private void CopySheet(string filePath, string sheetName, int pNum, string type)
        {
            SpreadsheetDocument book = SpreadsheetDocument.Open(filePath, true);
            var tempSheet = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
            WorkbookPart tempWBP = tempSheet.AddWorkbookPart();
            var asheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            var part = (WorksheetPart)book.WorkbookPart.GetPartById(asheets.First().Id.Value);

            WorksheetPart tempWSP = tempWBP.AddPart<WorksheetPart>(part);
            var copy = book.WorkbookPart.AddPart<WorksheetPart>(tempWSP);
            var sheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            var sheet = new Sheet();
            sheet.Id = book.WorkbookPart.GetIdOfPart(copy);
            sheet.Name = "Phase " + pNum + " " + type;
            uint id = 1;
            bool valid = false;
            while (!valid)
            {
                uint temp = id;
                foreach (OpenXmlElement e in sheets.ChildElements)
                {
                    var s = e as Sheet;
                    if (id == s.SheetId.Value)
                    {
                        id++;
                        break;
                    }
                }
                if (temp == id)
                    valid = true;
            }
            sheet.SheetId = id;
            sheets.Append(sheet);
            book.Close();
        }

        private void CopySheet(string filename, string sheetName, string clonedSheetName)
        {
            //Open workbook
            using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = mySpreadsheet.WorkbookPart;
                //Get the source sheet to be copied
                WorksheetPart sourceSheetPart = GetWorkSheetPart(workbookPart, sheetName);
                //Take advantage of AddPart for deep cloning
                SpreadsheetDocument tempSheet = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook); //mySpreadsheet.DocumentType);
                WorkbookPart tempWorkbookPart = tempSheet.AddWorkbookPart();
                WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart<WorksheetPart>(sourceSheetPart);
                //Add cloned sheet and all associated parts to workbook
                WorksheetPart clonedSheet = workbookPart.AddPart<WorksheetPart>(tempWorksheetPart);
                //Table definition parts are somewhat special and need unique ids...so let's make an id based on count
                int numTableDefParts = sourceSheetPart.GetPartsCountOfType<TableDefinitionPart>();
                //Clean up table definition parts (tables need unique ids)
                if (numTableDefParts != 0)
                    FixupTableParts(clonedSheet, numTableDefParts);
                //There should only be one sheet that has focus
                CleanView(clonedSheet);

                //Add new sheet to main workbook part
                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                Sheet copiedSheet = new Sheet();
                copiedSheet.Name = clonedSheetName;
                copiedSheet.Id = workbookPart.GetIdOfPart(clonedSheet);
                copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 4;
                sheets.Append(copiedSheet);
                //Save Changes
                //workbookPart.Workbook.Save();
            }
        }

        private WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)
        {
            //Get the relationship id of the sheetname
            string relId = workbookPart.Workbook.Descendants<Sheet>()
            .Where(s => s.Name.Value.Equals(sheetName))
            .First()
            .Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        }

        private void CleanView(WorksheetPart worksheetPart)
        {
            //There can only be one sheet that has focus
            SheetViews views = worksheetPart.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                worksheetPart.Worksheet.Save();
            }
        }

        private void FixupTableParts(WorksheetPart worksheetPart, int numTableDefParts)
        {
            var tableId = numTableDefParts;
            //Every table needs a unique id and name
            foreach (TableDefinitionPart tableDefPart in worksheetPart.TableDefinitionParts)
            {
                tableId++;
                tableDefPart.Table.Id = (uint)tableId;
                tableDefPart.Table.DisplayName = "CopiedTable" + tableId;
                tableDefPart.Table.Name = "CopiedTable" + tableId;
                tableDefPart.Table.Save();
            }
        }
         * */
        #endregion
    }

    public class CategoryLevelInfo
    {
        public string CategoryID { get; set; }
        public string Category { get; set; }
        public string ParentCategory { get; set; }
        public string ParentID { get; set; }
        public bool IsTopLevel { get; set; }
        public uint CategoryRow { get; set; }
        public uint SKURowNumber { get; set; }
        public uint SKUBeginRow { get; set; }
        public uint SKUEndRow { get; set; }
    }

    public class ShipWindowInfo
    {
        public string ShipWindow { get; set; }
        public string ShipWindowName { get; set; }
        public string DateWindowStart { get; set; }
        public string DateWindowEnd { get; set; }
        public string DateShip { get; set; }
        public string SummaryColumn { get; set; }
        public string SKUColumn { get; set; }
        public DateTime DateShipDate { get; set; }
    }
}
