using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Linq;
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
    public class TabularFormGenerator
    {
        public string BookingCatalog { get; set; }
        string TabularColorWithImageLink = ConfigurationManager.AppSettings["TabularColorWithImageLink"] == null ? "0" : ConfigurationManager.AppSettings["TabularColorWithImageLink"];
        string TabularImageLinkFormat = ConfigurationManager.AppSettings["TabularImageLinkFormat"];

        public void GenerateTabularOrderForm(string templateFile, string filename, string soldto, string catalog, string pricecode, string savedirectory)
        {
            WriteToLog(DateTime.Now.ToString() + " Begin " + filename);
            var catalogRow = GetCatalogInfo(soldto, catalog, savedirectory);
            //OrderType support, for BDEU UIUX
            if (catalogRow != null)
            {
                var orderType = GetColumnValue(catalogRow, "OrderType");
                if (!string.IsNullOrEmpty(orderType))
                {
                    var orderTypePartName = Path.GetInvalidFileNameChars().Aggregate(orderType, (current, c) => current.Replace(c.ToString(), ""));
                    var orderTypeTemplatefileSetting = "TabularTemplateFile_OrderType_" + orderTypePartName.Replace(" ", string.Empty);
                    WriteToLog(orderTypeTemplatefileSetting);
                    if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings[orderTypeTemplatefileSetting]))
                    {
                        templateFile = ConfigurationManager.AppSettings[orderTypeTemplatefileSetting];
                        if (!File.Exists(templateFile))
                            templateFile = ConfigurationManager.AppSettings["TabularTemplateFile"].ToString();
                        WriteToLog("catalog temp: " + templateFile);
                    }
                }
                BookingCatalog = GetColumnValue(catalogRow, "BookingCatalog");
            }
            //catalog support; for blackdimand UIUX DEV
            var catalogPartName = Path.GetInvalidFileNameChars().Aggregate(catalog, (current, c) => current.Replace(c.ToString(), ""));
            var catalogTemplatefileSetting = "TabularTemplateFile_" + catalogPartName.Replace(" ", string.Empty);
            WriteToLog(catalogTemplatefileSetting);
            if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings[catalogTemplatefileSetting]))
            {
                templateFile = ConfigurationManager.AppSettings[catalogTemplatefileSetting];
                if (!File.Exists(templateFile))
                    templateFile = ConfigurationManager.AppSettings["TabularTemplateFile"].ToString();
                WriteToLog("catalog temp: " + templateFile);
            }
            if (File.Exists(templateFile))
            {
                //get gridview data
                var formData = GetTabularOrderFormData(soldto, catalog, savedirectory);
                if (formData != null && formData.Rows.Count > 0)
                {
                    var fileInfo = new FileInfo(filename);
                    var tempFileInfo = new FileInfo(templateFile);
                    filename = filename.Replace(fileInfo.Extension, tempFileInfo.Extension);
                    //get shipment windows
                    DataTable windowDt = null;
                    var SupportShipWindow = ConfigurationManager.AppSettings["ShipmentWindowSupport"] == null ? "0" : ConfigurationManager.AppSettings["ShipmentWindowSupport"];
                    if (SupportShipWindow == "1")
                    {
                        windowDt = GetTabularShipmentWindow(soldto, catalog, savedirectory);
                    }
                    //save product data to excel cells
                    SaveTabularOrderFormDataToSheet(soldto, formData, templateFile, filename, catalog, windowDt, savedirectory);
                    //save customer data to sheet
                    var customerData = GetTabularCustomerData(soldto, catalog, savedirectory);
                    if (customerData != null && customerData.Rows.Count > 0)
                    {
                        SaveTabularCustomerDataToSheet(customerData, Path.Combine(savedirectory, filename));
                    }
                    //save shipmethod data to sheet
                    var methodData = GetAllShipMethod();
                    if (methodData != null && methodData.Rows.Count > 0)
                    {
                        SaveTabularShipMethodsDataToSheet(methodData, Path.Combine(savedirectory, filename));
                    }
                    //ship date validation
                    if (catalogRow != null)
                    {
                        SaveTabularDateValidationToSheet(catalogRow, Path.Combine(savedirectory, filename), formData.Rows.Count, windowDt);
                    }
                    //security
                    ProtectWorkbook(filename, savedirectory);
                }
            }
            else
                WriteToLog("Cannot find template file " + templateFile);
            WriteToLog(DateTime.Now.ToString() + " End " + filename);
        }

        private DataTable GetTabularOrderFormData(string soldto, string catalog, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_TabularGridView", conn);
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
                    WriteToLog("p_Offline_TabularGridView | GetTabularOrderFormData |" + soldto + "|" + catalog, ex, savedirectory);
                    return null;
                }
            }
        }

        private DataTable GetTabularShipmentWindow(string soldto, string catalog, string savedirectory)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("p_Offline_TabularShipmentWindow", conn);
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
                    WriteToLog("p_Offline_TabularShipmentWindow | GetTabularShipmentWindow |" + catalog, ex, savedirectory);
                    return null;
                }
            }
        }

        private void SaveTabularOrderFormDataToSheet(string soldto, DataTable formData, string templateFile, string filename, string catalog, DataTable windowDt, string savedirectory)
        {
            var filePath = Path.Combine(savedirectory, filename);
            File.Copy(templateFile, filePath, true);

            var sheetName = ConfigurationManager.AppSettings["TabularSheetName"] == null ? "TabularOfflineOrderForm" : ConfigurationManager.AppSettings["TabularSheetName"];
            var lineSheetName = ConfigurationManager.AppSettings["ProductLineSheetName"] == null ? "ProductLine_Style" : ConfigurationManager.AppSettings["ProductLineSheetName"];
            uint startRow = ConfigurationManager.AppSettings["TabularSKUStartRow"] == null ? 16 : uint.Parse(ConfigurationManager.AppSettings["TabularSKUStartRow"]);
            var skuColumn = ConfigurationManager.AppSettings["TabularSKUColumn"] == null ? "B" : ConfigurationManager.AppSettings["TabularSKUColumn"];
            var catalogColumn = ConfigurationManager.AppSettings["TabularCatalogColumn"] == null ? "F" : ConfigurationManager.AppSettings["TabularCatalogColumn"]; ;
            uint catalogRow = ConfigurationManager.AppSettings["TabularCatalogRow"] == null ? 2 : uint.Parse(ConfigurationManager.AppSettings["TabularCatalogRow"]);

            var TabularTotalCurrencyColumn = ConfigurationManager.AppSettings["TabularTotalCurrencyColumn"] == null ? "G" : ConfigurationManager.AppSettings["TabularTotalCurrencyColumn"];
            var TabularTotalCurrencyRow = ConfigurationManager.AppSettings["TabularTotalCurrencyRow"] == null ? 9 : int.Parse(ConfigurationManager.AppSettings["TabularTotalCurrencyRow"]);
            var TabularSubCurrencyColumn = ConfigurationManager.AppSettings["TabularSubCurrencyColumn"] == null ? "O" : ConfigurationManager.AppSettings["TabularSubCurrencyColumn"];
            var TabularSubCurrencyRow = ConfigurationManager.AppSettings["TabularSubCurrencyRow"] == null ? 9 : int.Parse(ConfigurationManager.AppSettings["TabularSubCurrencyRow"]);
            var TabularNumericValidationCell = ConfigurationManager.AppSettings["TabularNumericValidationCell"] == null ? "D1" : ConfigurationManager.AppSettings["TabularNumericValidationCell"];
            
            var maxDeptWidth = ConfigurationManager.AppSettings["TabularMaxDeptWidth"] == null ? 0 : double.Parse(ConfigurationManager.AppSettings["TabularMaxDeptWidth"]); ;
            var TabularCancelDate = ConfigurationManager.AppSettings["TabularCancelDate"] == null ? "0" : ConfigurationManager.AppSettings["TabularCancelDate"];
            uint TabularCancelDateRow = ConfigurationManager.AppSettings["TabularCancelDateRow"] == null ? 7 : uint.Parse(ConfigurationManager.AppSettings["TabularCancelDateRow"]);
            var TabularShipMethods = ConfigurationManager.AppSettings["TabularShipMethods"] == null ? "0" : ConfigurationManager.AppSettings["TabularShipMethods"];
            uint TabularShipMethodsRow = ConfigurationManager.AppSettings["TabularShipMethodsRow"] == null ? 8 : uint.Parse(ConfigurationManager.AppSettings["TabularShipMethodsRow"]);

           
            var OrderMultipleCheck = ConfigurationManager.AppSettings["OrderMultipleCheck"] == null ? "0" : ConfigurationManager.AppSettings["OrderMultipleCheck"];

            var TabularComColumnList = ConfigurationManager.AppSettings["TabularComColumnList"] == null ? "A|Catalog,B|SKU,C|UPC,D|Department,E|Style,F|ProductName,G|AttributeValue2,H|AttributeValue1" : ConfigurationManager.AppSettings["TabularComColumnList"];
            var TabularAlternateFillID = ConfigurationManager.AppSettings["TabularAlternateFillID"] == null ? -1 : int.Parse(ConfigurationManager.AppSettings["TabularAlternateFillID"]);

            var TabularPriceColumn = ConfigurationManager.AppSettings["TabularPriceColumn"] == null ? "I" : ConfigurationManager.AppSettings["TabularPriceColumn"];
            var TabularSKUOrderColumnList = ConfigurationManager.AppSettings["TabularSKUOrderColumnList"] == null ? "J|K|L|M|N" : ConfigurationManager.AppSettings["TabularSKUOrderColumnList"];
            var TabularRowTotalColumn = ConfigurationManager.AppSettings["TabularRowTotalColumn"] == null ? "O" : ConfigurationManager.AppSettings["TabularRowTotalColumn"];

            var TabularTotalMinimumOrderAmountBlockColumn = ConfigurationManager.AppSettings["TabularTotalMinimumOrderAmountBlockColumn"] ?? "F";
            var TabularTotalMinimumOrderAmountBlockMessage = ConfigurationManager.AppSettings["TabularTotalMinimumOrderAmountBlockMessage"] ?? "The Minimum Order Amount is [MinimumOrderAmount]";

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = null;
                ExcelWorksheet lineSheet = null;
                sheet = package.Workbook.Worksheets[sheetName];
                lineSheet = package.Workbook.Worksheets[lineSheetName];

                if (sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + templateFile);
                if (lineSheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + lineSheetName + ", file path:" + templateFile);

                if (sheet != null)
                {
                    var CurrencyCode = string.Empty;
                    CategoryList = new List<CategoryInfo>();
                    //Check previous departmentid, style, color
                    var preDept = string.Empty;
                    var preStyle = string.Empty;
                    var preColor = string.Empty;
                    var multipleList = new List<OrderMultipleDataValidation>();

                    var comColumnList = TabularComColumnList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    var TabularSKUOrderColumnLists = TabularSKUOrderColumnList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    var styleIndex = 0;

                    //set shipto dropdownlist if use formula to select shiptos
                    var TabularShipToRowWithFormula = ConfigurationManager.AppSettings["TabularShipToRowWithFormula"] ?? string.Empty;
                    if (!string.IsNullOrEmpty(TabularShipToRowWithFormula))
                    {
                        //var TabularShipToRowWithFormulaValue = int.Parse(TabularShipToRowWithFormula);
                        foreach (string ordercol in TabularSKUOrderColumnLists)
                        {
                            //var scell = SetCellValue(sheet, ordercol, (uint)TabularShipToRowWithFormulaValue, string.Empty);
                            //scell.Formula = "OFFSET(SoldToShipTo!$A$1,MATCH(D3,SoldToShipTo!$A:$A,0)-1,2,COUNTIF(SoldToShipTo!$A:$A,D3))";
                            var validationCell = sheet.DataValidations.AddListValidation(ordercol + TabularShipToRowWithFormula);
                            validationCell.Formula.ExcelFormula = "OFFSET(SoldToShipTo!$A$1,MATCH(D3,SoldToShipTo!$A:$A,0)-1,2,COUNTIF(SoldToShipTo!$A:$A,D3))";
                            validationCell.ShowErrorMessage = true;
                            validationCell.Error = "Please select value from dropdown list";
                        }
                    }

                    for (int i = 0; i < formData.Rows.Count; i++)
                    {
                        ExcelRange cell = null;
                        var row = formData.Rows[i];
                        if (i == 0)
                        {
                            var catalogCode = GetColumnValue(row, "CatalogCode");
                            cell = SetCellValue(sheet, catalogColumn, catalogRow, catalogCode);
                            SetCellStyle(cell);
                            //CurrencyCode
                            CurrencyCode = GetColumnValue(row, "CurrencyCode");
                            cell = SetCellValue(sheet, TabularTotalCurrencyColumn, (uint)TabularTotalCurrencyRow, CurrencyCode);
                            cell = SetCellValue(sheet, TabularSubCurrencyColumn, (uint)TabularSubCurrencyRow, CurrencyCode);

                            //MIN AMOUNT VALIDATION FOR BDEU
                            var minAmountSetting = GetB2BSetting("SiteSetting", "MinimumOrderAmountBlock");
                            if (minAmountSetting.ToUpper() == "CURRENCY")
                            { //=IF(AND(D14 <> 0, D14 < 100), "The Minimum Order Amount is 100.00", ""); [MinimumOrderAmount]
                                var minAmount = CheckMinimumOrderAmountBlock(soldto, catalogCode, 0);
                                var minAmountMessage = TabularTotalMinimumOrderAmountBlockMessage.Replace("[MinimumOrderAmount]", minAmount.ToString() + " " + CurrencyCode);
                                cell = SetCellValue(sheet, TabularTotalMinimumOrderAmountBlockColumn, (uint)TabularTotalCurrencyRow, string.Empty);
                                cell.Formula = "IF(AND(D14 <> 0, D14 < " + minAmount.ToString() + "), \"" + minAmountMessage + "\", \"\")";
                            }
                        }

                        var deptId = GetColumnValue(row, "DepartmentID");
                        var style = GetColumnValue(row, "Style");
                        var color = GetColumnValue(row, "AttributeValue2");

                        if ((style != preStyle) || (style == preStyle && deptId != preDept))
                            styleIndex++;
                        if (styleIndex % 2 == 1)
                        {
                            foreach (string column in comColumnList)
                            {
                                LoadTabularDataToCell(column, sheet, (uint)i + startRow, row, deptId, preDept, style, preStyle, color, preColor, i, formData.Rows.Count, skuColumn, null);
                            }
                            cell = SetCellValue(sheet,TabularPriceColumn, (uint)i + startRow, GetColumnValue(row, "PriceWholesale"), "2");
                            SetCellStyle(cell);

                            foreach (string ordercol in TabularSKUOrderColumnLists)
                            {
                                cell = SetCellValue(sheet, ordercol, (uint)i + startRow, string.Empty, "1");
                                cell.Style.Locked = false;
                                SetCellStyle(cell);
                            }
                        }
                        else
                        {
                            var TabularAlternateFillColor = string.Empty;
                            if (TabularAlternateFillID != -1)
                                TabularAlternateFillColor = ConfigurationManager.AppSettings["TabularAlternateFillColor"] == null ? "LightGray" : ConfigurationManager.AppSettings["TabularAlternateFillColor"];
                            System.Drawing.Color? bgcolor = null;
                            if (!string.IsNullOrEmpty(TabularAlternateFillColor))
                                bgcolor = System.Drawing.Color.FromName(TabularAlternateFillColor);

                            foreach (string column in comColumnList)
                            {
                                LoadTabularDataToCell(column, sheet, (uint)i + startRow, row, deptId, preDept, style, preStyle, color, preColor, i, formData.Rows.Count, skuColumn, bgcolor);
                            }
                            cell = SetCellValue(sheet, TabularPriceColumn, (uint)i + startRow, GetColumnValue(row, "PriceWholesale"), "2");
                            SetCellStyle(cell, bgcolor);

                            foreach (string ordercol in TabularSKUOrderColumnLists)
                            {
                                cell = SetCellValue(sheet,ordercol, (uint)i + startRow, string.Empty, "1");
                                cell.Style.Locked = false;
                                SetCellStyle(cell, bgcolor);
                            }
                        }

                        preDept = deptId;
                        preStyle = style;
                        preColor = color;

                        //total funcational
                        cell = SetCellValue(sheet, TabularRowTotalColumn, (uint)i + startRow, string.Empty, "1");
                        var cellformula = "SUM(" + TabularSKUOrderColumnLists[0] + (i + startRow).ToString() + ":" + TabularSKUOrderColumnLists[TabularSKUOrderColumnLists.Length - 1] + (i + startRow).ToString() + ")";
                        cell.Formula = cellformula;
                        SetCellStyle(cell, System.Drawing.Color.Yellow);

                        //ProductLine_Style
                        if (lineSheet != null)
                        {
                            //ProductLine
                            var productLine = GetColumnValue(row, "ProductLine");
                            cell = SetCellValue(lineSheet,  "A", (uint)i + startRow, productLine);
                            //Style
                            cell = SetCellValue(lineSheet, "B", (uint)i + startRow, style);
                            //set category
                            if (CategoryList.Count < 12 && !string.IsNullOrEmpty(productLine))
                            {
                                var results = from c in CategoryList where c.Category == productLine select c;
                                if (results.Count() <= 0)
                                    CategoryList.Add(new CategoryInfo() { Category = productLine });
                            }
                        }

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
                                multiple = new OrderMultipleDataValidation { Multiple = OrderMultiple, SequenceOfReferences = TabularSKUOrderColumnLists[0] + (i + startRow).ToString() + ":" + TabularSKUOrderColumnLists[TabularSKUOrderColumnLists.Length - 1] + (i + startRow).ToString() };
                                multipleList.Add(multiple);
                            }
                            else
                            {
                                multiple = multips.First();
                                multiple.SequenceOfReferences = multiple.SequenceOfReferences + " " + TabularSKUOrderColumnLists[0] + (i + startRow).ToString() + ":" + TabularSKUOrderColumnLists[TabularSKUOrderColumnLists.Length - 1] + (i + startRow).ToString();
                            }
                        }
                    }

                    if (TabularCancelDate == "0")
                    {
                        ExcelRow row = sheet.Row((int)TabularCancelDateRow);
                        row.Hidden = true;
                    }
                    if (TabularShipMethods == "0")
                    {
                        ExcelRow row = sheet.Row((int)TabularShipMethodsRow);
                        row.Hidden = true;
                    }

                    //Generate Category cells
                    if (lineSheet != null && CategoryList.Count > 0)
                    {
                        var TabularCategoryStartRow = 2;
                        var TabularCategory1Column = ConfigurationManager.AppSettings["TabularCategory1Column"] == null ? "J" : ConfigurationManager.AppSettings["TabularCategory1Column"];
                        var TabularCategory2Column = ConfigurationManager.AppSettings["TabularCategory2Column"] == null ? "M" : ConfigurationManager.AppSettings["TabularCategory2Column"];

                        if (CategoryList.Count > 10)
                            TabularCategoryStartRow = 1;
                        var category1Number = CategoryList.Count / 2 + CategoryList.Count % 2;
                        for (int i = 0; i < CategoryList.Count; i++)
                        {
                            if (i < category1Number)
                            {
                                CategoryList[i].Row = TabularCategoryStartRow + i;
                                CategoryList[i].Column = TabularCategory1Column;
                            }
                            else
                            {
                                CategoryList[i].Row = TabularCategoryStartRow + i - category1Number;
                                CategoryList[i].Column = TabularCategory2Column;
                            }
                            var cell = SetCellValue(sheet, CategoryList[i].Column, (uint)CategoryList[i].Row, CategoryList[i].Category);
                            SetCellStyle(cell);

                            CategoryList[i].Column = ((char)(char.Parse(CategoryList[i].Column) + 1)).ToString();
                            cell = SetCellValue(sheet, CategoryList[i].Column, (uint)CategoryList[i].Row, string.Empty, "2");
                            SetCellStyle(cell, System.Drawing.Color.Yellow);
                        }
                    }

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

                    if (windowDt != null && windowDt.Rows.Count > 0)
                    {
                        var TabularWindowsList = ConfigurationManager.AppSettings["TabularWindowsList"] == null ? "E|F|G" : ConfigurationManager.AppSettings["TabularWindowsList"];
                        var winlist = TabularWindowsList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        var currentIndex = 11;
                        var TabularWindownNameRightAlign = ConfigurationManager.AppSettings["TabularWindownNameRightAlign"] == null ? "0" : ConfigurationManager.AppSettings["TabularWindownNameRightAlign"];
                        foreach (DataRow row in windowDt.Rows)
                        {
                            var windowName = row.IsNull("WindowName") ? string.Empty : row["WindowName"].ToString();
                            if (!string.IsNullOrEmpty(windowName))
                            {
                                var cell = SetCellValue(sheet, winlist[0], (uint)currentIndex, windowName);
                                SetCellStyle(cell);
                                if (TabularWindownNameRightAlign == "1")
                                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                cell = SetCellValue(sheet, winlist[1], (uint)currentIndex, "0", "2");
                                SetCellStyle(cell, System.Drawing.Color.Yellow);
                                cell = SetCellValue(sheet, winlist[2], (uint)currentIndex, CurrencyCode);
                                SetCellStyle(cell);
                                currentIndex--;
                            }
                        }
                    }

                    package.Save();
                }
            }
        }

        private void LoadTabularDataToCell(string columnInfo, ExcelWorksheet sheet, uint rowIndex, DataRow row,
            string deptId, string preDept, string style, string preStyle, string color, string preColor, int loopIndex, int totalCount, string skuColumn, System.Drawing.Color? backColor)
        {
            ExcelRange cell = null;
            var colInfo = columnInfo.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

            var datatype = string.Empty;
            var TabularSpecialColumnType = ConfigurationManager.AppSettings["TabularSpecialColumnType"] == null ? "" : ConfigurationManager.AppSettings["TabularSpecialColumnType"];
            if (!string.IsNullOrEmpty(TabularSpecialColumnType))
            {
                var TabularSpecialColumnTypes = TabularSpecialColumnType.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string stype in TabularSpecialColumnTypes)
                {
                    var stypes = stype.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    if (stypes[0] == colInfo[1])
                    {
                        datatype = stypes[1];
                        break;
                    }
                }
            }
            //cell = GetCell(sheet, );
            switch (colInfo[1])
            {
                case "Catalog":
                    cell = SetCellValue(sheet, colInfo[0], rowIndex, GetColumnValue(row, "CatalogCode"), datatype);
                    SetCellStyle(cell, backColor);
                    break;
                case "SKU":
                    cell = SetCellValue(sheet, colInfo[0], rowIndex, GetColumnValue(row, "SKU"), datatype);
                    SetCellStyle(cell, backColor);
                    break;
                case "UPC":
                    cell = SetCellValue(sheet, colInfo[0], rowIndex, GetColumnValue(row, "UPC"), datatype);
                    SetCellStyle(cell, backColor);
                    break;
                case "Department":
                    if (deptId != preDept)
                    {
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, GetColumnValue(row, "NavigationBar"), datatype);
                        cell.Style.WrapText = true;
                        if (loopIndex == totalCount - 1)
                            SetCellStyle(cell, backColor);
                        else
                            SetCellStyle(cell, backColor, true, true, true, false);
                    }
                    else
                    {
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                        if (loopIndex == totalCount - 1)
                            SetCellStyle(cell, backColor, true, true, false, true);
                        else
                            SetCellStyle(cell, backColor, true, true, false, false);
                    }
                    break;
                case "Style":
                    if ((style != preStyle) || (style == preStyle && deptId != preDept))
                    {
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, style, datatype);
                        if (loopIndex == totalCount - 1)
                            SetCellStyle(cell, backColor);
                        else
                            SetCellStyle(cell, backColor, true, true, true, false);
                    }
                    else
                    {
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                        if (loopIndex == totalCount - 1)
                            SetCellStyle(cell, backColor, true, true, false, true);
                        else
                            SetCellStyle(cell, backColor, true, true, false, false);
                    }
                    break;
                case "ProductName":
                    if (skuColumn == "F")
                    {
                        cell = SetCellValue(sheet, skuColumn, rowIndex, GetColumnValue(row, "SKU"), datatype);
                        SetCellStyle(cell, backColor);
                    }
                    else
                    {
                        var productName = GetColumnValue(row, "ProductName");
                        if ((style != preStyle) || (style == preStyle && deptId != preDept))
                        {
                            cell = SetCellValue(sheet, colInfo[0], rowIndex, productName, datatype);
                            if (loopIndex == totalCount - 1)
                                SetCellStyle(cell, backColor);
                            else
                                SetCellStyle(cell, backColor, true, true, true, false);
                        }
                        else
                        {
                            cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                            if (loopIndex == totalCount - 1)
                                SetCellStyle(cell, backColor, true, true, false, true);
                            else
                                SetCellStyle(cell, backColor, true, true, false, false);
                        }
                    }
                    break;
                case "AttributeValue2":
                    if ((color != preColor) || (color == preColor && deptId != preDept) || (color == preColor && style != preStyle))
                    {
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, color, datatype);
                        if (loopIndex == totalCount - 1)
                            SetCellStyle(cell, backColor);
                        else
                            SetCellStyle(cell, backColor, true, true, true, false);

                        //add image link to cell, bdeu #1297
                        if (TabularColorWithImageLink == "1")
                        {
                            if (!string.IsNullOrEmpty(TabularImageLinkFormat))
                            {
                                var styleImageValue = GetColumnValue(row, "ImageName");
                                styleImageValue = styleImageValue.Replace("small/", string.Empty).Replace(".jpg", string.Empty).Replace(".png", string.Empty).Trim();
                                var cellImageLinkFormat = string.Format(TabularImageLinkFormat, styleImageValue, color);

                                cell.Formula = cellImageLinkFormat;
                                cell.Style.Font.UnderLine = true;
                                cell.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                            }
                        }
                    }
                    else
                    {
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty);
                        if (loopIndex == totalCount - 1)
                            SetCellStyle(cell, backColor, true, true, false, true);
                        else
                            SetCellStyle(cell, backColor, true, true, false, false);
                    }
                    break;
                case "AttributeValue1":
                    cell = SetCellValue(sheet, colInfo[0], rowIndex, GetColumnValue(row, "AttributeValue1"), datatype);
                    SetCellStyle(cell, backColor);
                    break;
                case "LaunchDate":
                    var cvalue = GetColumnValue(row, colInfo[1]);
                    DateTime dt;
                    if (cvalue.Length > 0 && DateTime.TryParse(cvalue, out dt))
                    {
                        var LaunchDateFormat = ConfigurationManager.AppSettings["LaunchDateFormat"]; //For BDEU
                        if (!string.IsNullOrEmpty(LaunchDateFormat))
                        {
                            cell = SetCellValue(sheet, colInfo[0], rowIndex, dt.ToString(LaunchDateFormat));
                        }
                        else
                        {
                            var tformula = string.Format("DATE({0},{1},{2})", dt.Year, dt.Month, dt.Day);
                            cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty, "3");
                            cell.Formula = tformula;
                        }
                        SetCellStyle(cell, backColor);
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                    else //#bdeu 186 set cell style when launch date is empty
                    {
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, string.Empty, "3");
                        SetCellStyle(cell, backColor);
                    }
                    break;
                default:
                    {
                        var cellValue = GetColumnValue(row, colInfo[1]);
                        if (colInfo[1] == "DateQty" || colInfo[1] == "ATPMaxQty")
                        {
                            if (BookingCatalog == "1") //bdeu #1297
                                cellValue = string.Empty;
                        }
                        cell = SetCellValue(sheet, colInfo[0], rowIndex, cellValue, datatype);
                        SetCellStyle(cell, backColor);
                    }
                    break;
            }
        }

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
                if(sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + filePath);
                if (sheet != null)
                {
                    var columnList = new List<string>();
                    for (int i = 0; i < columnLength; i++)
                    {
                        var col = ((char)(char.Parse(startColumn) + i)).ToString();
                        var cell = GetCell(sheet, col, 1);
                        if (cell != null && cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                            columnList.Add(cell.Value.ToString());
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

        private DataTable GetAllShipMethod()
        {
            var TabularShipMethods = ConfigurationManager.AppSettings["TabularShipMethods"] == null ? "0" : ConfigurationManager.AppSettings["TabularShipMethods"];
            if (TabularShipMethods == "0")
                return null;

            DataSet shipmethodds = new DataSet();
            string errMsg = "";
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand("p_OffLineGetShipMethod", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);

            //Add params 
            conn.Open();
            try
            {
                adapter.Fill(shipmethodds);
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                WriteToLog("p_OffLineGetShipMethod", ex, string.Empty);
            }
            finally
            {
                conn.Close();
            }
            if (shipmethodds.Tables.Count > 0)
            {
                return shipmethodds.Tables[0];
            }
            return null;
        }

        private void SaveTabularShipMethodsDataToSheet(DataTable methodsData, string filePath)
        {
            var sheetName = ConfigurationManager.AppSettings["TabularShipMethodsSheetName"] == null ? "ShipMethods" : ConfigurationManager.AppSettings["TabularShipMethodsSheetName"];

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = null;
                sheet = package.Workbook.Worksheets[sheetName];
                if(sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + filePath);
                if (sheet != null)
                {
                    for (int i = 0; i < methodsData.Rows.Count; i++)
                    {
                        var row = methodsData.Rows[i];
                        SetCellValue(sheet,"A", (uint)(i + 1), GetColumnValue(row, "code"));
                        SetCellValue(sheet,"B", (uint)(i + 1), GetColumnValue(row, "description")); 
                    }
                    package.Save();
                }
            }
        }

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

        private void SaveTabularDateValidationToSheet(DataRow catalogRow, string filePath, int SKURowNumber, DataTable windowDt)
        {
            var sheetName = ConfigurationManager.AppSettings["GlobalSheetName"] == null ? "GlobalVariables" : ConfigurationManager.AppSettings["GlobalSheetName"];
            var startRow = ConfigurationManager.AppSettings["TabularSKUStartRow"] == null ? 16 : int.Parse(ConfigurationManager.AppSettings["TabularSKUStartRow"]);

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet sheet = null;
                    sheet = package.Workbook.Worksheets[sheetName];
                    if (sheet == null)
                        WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + filePath);
                    if (sheet != null)
                    {
                        Cell cell = null;
                        //BOOKING ORDER
                        var bookingOrder = GetColumnValue(catalogRow, "BookingCatalog");
                        SetCellValue(sheet, "B", 3, (string.IsNullOrEmpty(bookingOrder) ? "0" : bookingOrder));
                        //DateShipStart
                        if (!catalogRow.IsNull("DateShipStart"))
                        {
                            var date = ((DateTime)catalogRow["DateShipStart"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);//1033
                            SetCellValue(sheet, "B", 4, date, "3");
                        }
                        //DateShipEnd
                        if (!catalogRow.IsNull("DateShipEnd"))
                        {
                            var date = ((DateTime)catalogRow["DateShipEnd"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);//1033
                            SetCellValue(sheet, "B", 5, date, "3");
                        }
                        //ReqDateDays
                        var ReqDateDays = (int)catalogRow["ReqDateDays"];
                        SetCellValue(sheet, "B", 6, ReqDateDays.ToString());
                        //ReqDate
                        //if (ReqDateDays > -1)
                        //{
                        //    cell = GetCell(sheet, "B", 7);
                        //    var reqDate = DateTime.Today.AddDays(ReqDateDays).ToString((new System.Globalization.CultureInfo(1033)).DateTimeFormat.ShortDatePattern);
                        //    SetCellValue(cell, reqDate, 4, TabularCommonDateFormat);
                        //}
                        //ReqDefaultDays
                        var ReqDefaultDays = (int)catalogRow["DefaultReqDays"];
                        SetCellValue(sheet, "B", 7, ReqDefaultDays.ToString());
                        //ReqDefaultDate
                        //if (ReqDefaultDays > -1)
                        //{
                        //    cell = GetCell(sheet, "B", 9);
                        //    var reqDate = DateTime.Today.AddDays(ReqDefaultDays).ToString((new System.Globalization.CultureInfo(1033)).DateTimeFormat.ShortDatePattern);
                        //    SetCellValue(cell, reqDate, 4, TabularCommonDateFormat);
                        //}
                        //CancelDateDays
                        var CancelDateDays = (int)catalogRow["CancelDateDays"];
                        SetCellValue(sheet, "B", 8, CancelDateDays.ToString());
                        //CancelDefaultDays
                        var CancelDefaultDays = (int)catalogRow["CancelDefaultDays"];
                        SetCellValue(sheet, "B", 9, CancelDefaultDays.ToString());
                        SetCellValue(sheet, "B", 10, (SKURowNumber + startRow - 1).ToString());
                        SetCellValue(sheet, "A", 11, "Categories");
                        //set for category configuration
                        for (int i = 0; i < CategoryList.Count; i++)
                        {
                            SetCellValue(sheet, "B", 11 + (uint)i, CategoryList[i].Category);
                            SetCellValue(sheet, "C", 11 + (uint)i, CategoryList[i].Column);
                            SetCellValue(sheet, "D", 11 + (uint)i, CategoryList[i].Row.ToString());
                        }

                        if (windowDt != null && windowDt.Rows.Count > 0)
                        {
                            var currentIndex = 3;
                            foreach (DataRow row in windowDt.Rows)
                            {
                                var windowName = row.IsNull("WindowName") ? string.Empty : row["WindowName"].ToString();
                                if (!string.IsNullOrEmpty(windowName))
                                {
                                    SetCellValue(sheet, "E", (uint)currentIndex, windowName);
                                    var date = ((DateTime)row["DateBegin"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);
                                    SetCellValue(sheet, "F", (uint)currentIndex, date, "3");
                                    date = ((DateTime)row["DateEnd"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);
                                    SetCellValue(sheet, "G", (uint)currentIndex, date, "3");
                                    currentIndex++;
                                }
                            }
                        }
                        package.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToLog("SaveTabularDateValidationToSheet", ex, string.Empty);
            }
        }

        private void ProtectWorkbook(string filename, string savedirectory)
        {
            var pwd = ConfigurationManager.AppSettings["ProtectPassword"] == null ? "Plumriver" : ConfigurationManager.AppSettings["ProtectPassword"];
            var password = pwd;//HashPassword(pwd);
            var filePath = Path.Combine(savedirectory, filename);
            var sheetName = ConfigurationManager.AppSettings["TabularSheetName"] == null ? "TabularOfflineOrderForm" : ConfigurationManager.AppSettings["TabularSheetName"];


            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet sheet = null;
                //TabularOfflineOrderForm
                sheet = package.Workbook.Worksheets[sheetName];
                if (sheet != null)
                {
                    sheet.Protection.AllowAutoFilter = true;
                    sheet.Protection.IsProtected = true;
                    sheet.Protection.AllowEditObject = true;
                    sheet.Protection.AllowEditScenarios = true;
                    sheet.Protection.SetPassword(password);
                }

                sheetName = ConfigurationManager.AppSettings["GlobalSheetName"] == null ? "GlobalVariables" : ConfigurationManager.AppSettings["GlobalSheetName"];
                var GlobalVariablesProtect = ConfigurationManager.AppSettings["GlobalVariablesProtect"] == null ? "0" : ConfigurationManager.AppSettings["GlobalVariablesProtect"];
                var pwdColumn = ConfigurationManager.AppSettings["GlobalPWDColumn"] == null ? "B" : ConfigurationManager.AppSettings["GlobalPWDColumn"];
                var pwdRow = ConfigurationManager.AppSettings["GlobalPWDRow"] == null ? 200 : int.Parse(ConfigurationManager.AppSettings["GlobalPWDRow"]);

                sheet = null;
                sheet = package.Workbook.Worksheets[sheetName];
                if (sheet == null)
                    WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + filePath);
                if (sheet != null)
                {
                    SetCellValue(sheet, pwdColumn, (uint)(pwdRow), pwd);
                    sheet.Row(pwdRow).Hidden = true;
                }

                //GlobalVariables
                if (GlobalVariablesProtect == "1")
                {
                    if (sheet != null)
                    {
                        sheet.Protection.AllowAutoFilter = true;
                        sheet.Protection.IsProtected = true;
                        sheet.Protection.AllowEditObject = true;
                        sheet.Protection.AllowEditScenarios = true;
                        sheet.Protection.SetPassword(password);
                    }
                }
                //SoldToShipTo
                sheetName = ConfigurationManager.AppSettings["CustomerSheetName"] == null ? "SoldToShipTo" : ConfigurationManager.AppSettings["CustomerSheetName"];
                sheet = package.Workbook.Worksheets[sheetName];
                if (sheet != null)
                {
                    sheet.Protection.AllowAutoFilter = true;
                    sheet.Protection.IsProtected = true;
                    sheet.Protection.AllowEditObject = true;
                    sheet.Protection.AllowEditScenarios = true;
                    sheet.Protection.SetPassword(password);
                }
                //THE WHOLE WORK BOOK
                package.Workbook.Protection.LockStructure = true;

                package.Save();
            }
        }


        public decimal CheckMinimumOrderAmountBlock(string SoldTo, string CatalogCode, decimal CartAmount)
        {
            decimal ret = 0;

            SqlParameter paramSoldTo = new SqlParameter("@DealerNumber", SqlDbType.NVarChar, 80);
            SqlParameter paramFromOLOF = new SqlParameter("@FromOLOF", SqlDbType.Int);
            SqlParameter paramCatalogCode = new SqlParameter("@CatalogCode", SqlDbType.NVarChar, 80);
            SqlParameter paramCartAmount = new SqlParameter("@CartAmount", SqlDbType.Decimal);
            SqlParameter paramMinOrderQuantityBlock = new SqlParameter("@MinOrderAmountBlock", SqlDbType.Int);
            SqlParameter paramErrorMessage = new SqlParameter("@ErrorMessage", SqlDbType.NVarChar, 2000);

            paramSoldTo.Value = SoldTo;
            paramCatalogCode.Value = CatalogCode;
            paramFromOLOF.Value = 1;
            paramCartAmount.Value = CartAmount;
            paramMinOrderQuantityBlock.Direction = ParameterDirection.Output;
            paramErrorMessage.Direction = ParameterDirection.Output;

            // Create & open a SqlConnection, and dispose of it after we are done
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["connString"].ConnectionString))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "p_MinimumOrderAmountBlock";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddRange(new SqlParameter[] { paramSoldTo, paramCatalogCode, paramFromOLOF, paramCartAmount, paramMinOrderQuantityBlock, paramErrorMessage });

                    connection.Open();


                    var ds = new DataSet();
                    SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                    adapt.Fill(ds);
                    if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        var row = ds.Tables[0].Rows[0];
                        ret = row.IsNull("MinimumOrderAmount") ? 0 : (decimal)row["MinimumOrderAmount"];
                    }
                }
                catch (Exception ex)
                {
                    WriteToLog("p_MinimumOrderAmountBlock", ex, string.Empty);
                }
                connection.Close();
            }
            return ret;
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
                    if (!string.IsNullOrEmpty(value))
                    {
                        Int64 cellIntValue = 0;
                        if (Int64.TryParse(value, out cellIntValue))
                            cell.Value = cellIntValue;
                        else
                            cell.Value = value;
                    }
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
                case "4": //int
                    cell.Style.Numberformat.Format = "#";
                    if (!string.IsNullOrEmpty(value))
                    {
                        Int64 cellIntValue = 0;
                        if (Int64.TryParse(value, out cellIntValue))
                            cell.Value = cellIntValue;
                        else
                            cell.Value = value;
                    }
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

        private ExcelRange GetCell(ExcelWorksheet sheet, string column, uint row)
        {
            var address = column + row.ToString();
            return GetCell(sheet, address);
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
            if (!borderLeft)
                cell.Style.Border.Left.Style = ExcelBorderStyle.None;
            if (!borderRight)
                cell.Style.Border.Right.Style = ExcelBorderStyle.None;
            if (!borderTop)
                cell.Style.Border.Top.Style = ExcelBorderStyle.None;
            if (!borderBottom)
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.None;
        }


        public string GetB2BSetting(string category, string setting)
        {
            string ret = string.Empty;

            SqlParameter paramCategory = new SqlParameter("@category", SqlDbType.VarChar, 80);
            SqlParameter paramSetting = new SqlParameter("@setting", SqlDbType.VarChar, 80);
            SqlParameter paramValue = new SqlParameter("@Value", SqlDbType.VarChar, 800);

            paramCategory.Value = category;
            paramSetting.Value = setting;
            paramValue.Direction = ParameterDirection.Output;

            // Create & open a SqlConnection, and dispose of it after we are done
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["connString"].ConnectionString))
            {

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = connection;
                cmd.CommandText = "p_B2B_Setting";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddRange(new SqlParameter[] { paramCategory, paramSetting, paramValue });

                connection.Open();
                cmd.ExecuteNonQuery();

                try
                {
                    ret = Convert.ToString(paramValue.Value);
                }
                catch (Exception ex)
                {
                    WriteToLog("p_B2B_Setting", ex, string.Empty);
                }
            }
            return ret;
        }





        static uint TabularCommonTextFormat = ConfigurationManager.AppSettings["TabularCommonTextFormat"] == null ? 4 : uint.Parse(ConfigurationManager.AppSettings["TabularCommonTextFormat"]);
        static uint TabularCommonDateFormat = ConfigurationManager.AppSettings["TabularCommonDateFormat"] == null ? 16 : uint.Parse(ConfigurationManager.AppSettings["TabularCommonDateFormat"]);
        static uint TabularCommonDoubleFormat = ConfigurationManager.AppSettings["TabularCommonDoubleFormat"] == null ? 37 : uint.Parse(ConfigurationManager.AppSettings["TabularCommonDoubleFormat"]);

        private List<CategoryInfo> CategoryList { get; set; }

        public void GenerateTabularOrderForm_old(string templateFile, string filename, string soldto, string catalog, string pricecode, string savedirectory)
        {
            WriteToLog(DateTime.Now.ToString() + " Begin " + filename);
            if (File.Exists(templateFile))
            {
                //get gridview data
                var formData = GetTabularOrderFormData(soldto, catalog, savedirectory);
                if (formData != null && formData.Rows.Count > 0)
                {
                    var fileInfo = new FileInfo(filename);
                    var tempFileInfo = new FileInfo(templateFile);
                    filename = filename.Replace(fileInfo.Extension, tempFileInfo.Extension);
                    //get shipment windows
                    DataTable windowDt = null;
                    var SupportShipWindow = ConfigurationManager.AppSettings["ShipmentWindowSupport"] == null ? "0" : ConfigurationManager.AppSettings["ShipmentWindowSupport"];
                    if (SupportShipWindow == "1")
                    {
                        windowDt = GetTabularShipmentWindow(soldto, catalog, savedirectory);
                    }
                    //save product data to excel cells
                    SaveTabularOrderFormDataToSheet_old(formData, templateFile, filename, catalog, windowDt, savedirectory);
                    //save customer data to sheet
                    var customerData = GetTabularCustomerData(soldto, catalog, savedirectory);
                    if (customerData != null && customerData.Rows.Count > 0)
                    {
                        SaveTabularCustomerDataToSheet(customerData, Path.Combine(savedirectory, filename));
                    }
                    //save shipmethod data to sheet
                    var methodData = GetAllShipMethod();
                    if (methodData != null && methodData.Rows.Count > 0)
                    {
                        SaveTabularShipMethodsDataToSheet(methodData, Path.Combine(savedirectory, filename));
                    }
                    //ship date validation
                    var catalogRow = GetCatalogInfo(soldto, catalog, savedirectory);
                    if (catalogRow != null)
                    {
                        SaveTabularDateValidationToSheet(catalogRow, Path.Combine(savedirectory, filename), formData.Rows.Count, windowDt);
                    }
                    //security
                    ProtectWorkbook(filename, savedirectory);
                }
            }
            else
                WriteToLog("Cannot find template file " + templateFile);
            WriteToLog(DateTime.Now.ToString() + " End " + filename);
        }

        private void SaveTabularOrderFormDataToSheet_old(DataTable formData, string templateFile, string filename, string catalog, DataTable windowDt, string savedirectory)
        {
            var filePath = Path.Combine(savedirectory, filename);
            File.Copy(templateFile, filePath, true);

            var sheetName = ConfigurationManager.AppSettings["TabularSheetName"] == null ? "TabularOfflineOrderForm" : ConfigurationManager.AppSettings["TabularSheetName"];
            uint startRow = ConfigurationManager.AppSettings["TabularSKUStartRow"] == null ? 16 : uint.Parse(ConfigurationManager.AppSettings["TabularSKUStartRow"]);
            var skuColumn = ConfigurationManager.AppSettings["TabularSKUColumn"] == null ? "B" : ConfigurationManager.AppSettings["TabularSKUColumn"];
            var catalogColumn = ConfigurationManager.AppSettings["TabularCatalogColumn"] == null ? "F" : ConfigurationManager.AppSettings["TabularCatalogColumn"]; ;
            uint catalogRow = ConfigurationManager.AppSettings["TabularCatalogRow"] == null ? 2 : uint.Parse(ConfigurationManager.AppSettings["TabularCatalogRow"]);
            var maxDeptWidth = ConfigurationManager.AppSettings["TabularMaxDeptWidth"] == null ? 0 : double.Parse(ConfigurationManager.AppSettings["TabularMaxDeptWidth"]); ;
            var TabularCancelDate = ConfigurationManager.AppSettings["TabularCancelDate"] == null ? "0" : ConfigurationManager.AppSettings["TabularCancelDate"];
            uint TabularCancelDateRow = ConfigurationManager.AppSettings["TabularCancelDateRow"] == null ? 7 : uint.Parse(ConfigurationManager.AppSettings["TabularCancelDateRow"]);
            var TabularShipMethods = ConfigurationManager.AppSettings["TabularShipMethods"] == null ? "0" : ConfigurationManager.AppSettings["TabularShipMethods"];
            uint TabularShipMethodsRow = ConfigurationManager.AppSettings["TabularShipMethodsRow"] == null ? 8 : uint.Parse(ConfigurationManager.AppSettings["TabularShipMethodsRow"]);

            var TabularTotalCurrencyColumn = ConfigurationManager.AppSettings["TabularTotalCurrencyColumn"] == null ? "G" : ConfigurationManager.AppSettings["TabularTotalCurrencyColumn"];
            var TabularTotalCurrencyRow = ConfigurationManager.AppSettings["TabularTotalCurrencyRow"] == null ? 9 : int.Parse(ConfigurationManager.AppSettings["TabularTotalCurrencyRow"]);
            var TabularSubCurrencyColumn = ConfigurationManager.AppSettings["TabularSubCurrencyColumn"] == null ? "O" : ConfigurationManager.AppSettings["TabularSubCurrencyColumn"];
            var TabularSubCurrencyRow = ConfigurationManager.AppSettings["TabularSubCurrencyRow"] == null ? 9 : int.Parse(ConfigurationManager.AppSettings["TabularSubCurrencyRow"]);
            var TabularNumericValidationCell = ConfigurationManager.AppSettings["TabularNumericValidationCell"] == null ? "D1" : ConfigurationManager.AppSettings["TabularNumericValidationCell"];
            var lineSheetName = ConfigurationManager.AppSettings["ProductLineSheetName"] == null ? "ProductLine_Style" : ConfigurationManager.AppSettings["ProductLineSheetName"];
            var OrderMultipleCheck = ConfigurationManager.AppSettings["OrderMultipleCheck"] == null ? "0" : ConfigurationManager.AppSettings["OrderMultipleCheck"];

            var TabularComColumnList = ConfigurationManager.AppSettings["TabularComColumnList"] == null ? "A|Catalog,B|SKU,C|UPC,D|Department,E|Style,F|ProductName,G|AttributeValue2,H|AttributeValue1" : ConfigurationManager.AppSettings["TabularComColumnList"];
            var TabularAlternateFillID = ConfigurationManager.AppSettings["TabularAlternateFillID"] == null ? -1 : int.Parse(ConfigurationManager.AppSettings["TabularAlternateFillID"]);

            var TabularPriceColumn = ConfigurationManager.AppSettings["TabularPriceColumn"] == null ? "I" : ConfigurationManager.AppSettings["TabularPriceColumn"];
            var TabularSKUOrderColumnList = ConfigurationManager.AppSettings["TabularSKUOrderColumnList"] == null ? "J|K|L|M|N" : ConfigurationManager.AppSettings["TabularSKUOrderColumnList"];
            var TabularRowTotalColumn = ConfigurationManager.AppSettings["TabularRowTotalColumn"] == null ? "O" : ConfigurationManager.AppSettings["TabularRowTotalColumn"];

            Worksheet sheet = null;
            Worksheet lineSheet = null;
            SpreadsheetDocument book = SpreadsheetDocument.Open(filePath, true);
            var sheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() > 0)
            {
                WorksheetPart worksheetPart = (WorksheetPart)book.WorkbookPart.GetPartById(sheets.First().Id.Value);
                sheet = worksheetPart.Worksheet;
            }
            else
                WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + templateFile);

            sheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == lineSheetName);
            if (sheets.Count() > 0)
            {
                WorksheetPart worksheetPart = (WorksheetPart)book.WorkbookPart.GetPartById(sheets.First().Id.Value);
                lineSheet = worksheetPart.Worksheet;
            }
            if (sheet != null)
            {
                var CurrencyCode = string.Empty;
                CategoryList = new List<CategoryInfo>();
                //Style Format
                var styleSheet = book.WorkbookPart.WorkbookStylesPart.Stylesheet;
                var fontId = CreateFontFormat(styleSheet);
                var borderAllId = CreateBorderFormat(styleSheet, true, true, true, true);
                var borderLRId = CreateBorderFormat(styleSheet, true, true, false, false);
                var borderLRTId = CreateBorderFormat(styleSheet, true, true, true, false);
                var borderLRBId = CreateBorderFormat(styleSheet, true, true, false, true);
                //var borderTBId = CreateBorderFormat(styleSheet, false, false, true, true);
                var doubleStyleId = CreateCellFormat(styleSheet, fontId, null, borderAllId, UInt32Value.FromUInt32(2));
                var intStyleId = CreateCellFormat(styleSheet, fontId, 4, borderAllId, UInt32Value.FromUInt32(1));
                var intLStyleId = CreateCellFormat(styleSheet, fontId, null, borderAllId, UInt32Value.FromUInt32(1), false, false, HorizontalAlignmentValues.General);
                var textAllStyleId = CreateCellFormat(styleSheet, fontId, null, borderAllId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                var textLRStyleId = CreateCellFormat(styleSheet, fontId, null, borderLRId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                var textLRTStyleId = CreateCellFormat(styleSheet, fontId, null, borderLRTId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                var textLRBStyleId = CreateCellFormat(styleSheet, fontId, null, borderLRBId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                var textNullStyleId = CreateCellFormat(styleSheet, fontId, null, null, UInt32Value.FromUInt32(49));
                var textDAllStyleId = CreateCellFormat(styleSheet, fontId, null, borderAllId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                var textDLRTStyleId = CreateCellFormat(styleSheet, fontId, null, borderLRTId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);

                var textAlternateAllStyleId = textAllStyleId;
                var textAlternateDAllStyleId = textDAllStyleId;
                var textAlternateDLRTStyleId = textDLRTStyleId;
                var textAlternateLRBStyleId = textLRBStyleId;
                var textAlternateLRStyleId = textLRStyleId;
                var textAlternateLRTStyleId = textLRTStyleId;
                var intAlternateLStyleId = intLStyleId;
                var doubleAlternateStyleId = doubleStyleId;
                if (TabularAlternateFillID != -1)
                {
                    textAlternateAllStyleId = CreateCellFormat(styleSheet, fontId, (uint)TabularAlternateFillID, borderAllId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                    textAlternateLRStyleId = CreateCellFormat(styleSheet, fontId, (uint)TabularAlternateFillID, borderLRId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                    textAlternateLRTStyleId = CreateCellFormat(styleSheet, fontId, (uint)TabularAlternateFillID, borderLRTId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                    textAlternateLRBStyleId = CreateCellFormat(styleSheet, fontId, (uint)TabularAlternateFillID, borderLRBId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                    textAlternateDAllStyleId = CreateCellFormat(styleSheet, fontId, (uint)TabularAlternateFillID, borderAllId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                    textAlternateDLRTStyleId = CreateCellFormat(styleSheet, fontId, (uint)TabularAlternateFillID, borderLRTId, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.General);
                    intAlternateLStyleId = CreateCellFormat(styleSheet, fontId, (uint)TabularAlternateFillID, borderAllId, UInt32Value.FromUInt32(1), false, false, HorizontalAlignmentValues.General);
                    doubleAlternateStyleId = CreateCellFormat(styleSheet, fontId, (uint)TabularAlternateFillID, borderAllId, UInt32Value.FromUInt32(2));
                }
                //Check previous departmentid, style, color
                var preDept = string.Empty;
                var preStyle = string.Empty;
                var preColor = string.Empty;
                var multipleList = new List<OrderMultipleDataValidation>();

                var comColumnList = TabularComColumnList.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                var TabularSKUOrderColumnLists = TabularSKUOrderColumnList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                var styleIndex = 0;
                for (int i = 0; i < formData.Rows.Count; i++)
                {
                    Cell cell = null;
                    var row = formData.Rows[i];
                    if (i == 0)
                    {
                        cell = GetCell(sheet, catalogColumn, catalogRow);
                        SetCellValue(cell, GetColumnValue(row, "CatalogCode"), 1, textAllStyleId);

                        //CurrencyCode
                        CurrencyCode = GetColumnValue(row, "CurrencyCode");
                        var cCell = GetCell(sheet, TabularTotalCurrencyColumn, (uint)TabularTotalCurrencyRow);
                        SetCellValue(cCell, CurrencyCode, 1, textNullStyleId);
                        cCell = GetCell(sheet, TabularSubCurrencyColumn, (uint)TabularSubCurrencyRow);
                        SetCellValue(cCell, CurrencyCode, 1, textNullStyleId);
                    }

                    var deptId = GetColumnValue(row, "DepartmentID");
                    var style = GetColumnValue(row, "Style");
                    var color = GetColumnValue(row, "AttributeValue2");

                    if ((style != preStyle) || (style == preStyle && deptId != preDept))
                        styleIndex++;
                    if (styleIndex % 2 == 1)
                    {
                        foreach (string column in comColumnList)
                        {
                            LoadTabularDataToCell_old(column, sheet, (uint)i + startRow, row, deptId, preDept, style, preStyle, color, preColor, i, formData.Rows.Count, skuColumn,
                                textAllStyleId, textDAllStyleId, textDLRTStyleId, textLRBStyleId, textLRStyleId, textLRTStyleId);
                        }
                        cell = GetCell(sheet, TabularPriceColumn, (uint)i + startRow);
                        SetCellValue(cell, GetColumnValue(row, "PriceWholesale"), 2, doubleStyleId);

                        foreach (string ordercol in TabularSKUOrderColumnLists)
                        {
                            cell = GetCell(sheet, ordercol, (uint)i + startRow);
                            SetCellValue(cell, string.Empty, 3, intLStyleId);
                        }
                    }
                    else
                    {
                        foreach (string column in comColumnList)
                        {
                            LoadTabularDataToCell_old(column, sheet, (uint)i + startRow, row, deptId, preDept, style, preStyle, color, preColor, i, formData.Rows.Count, skuColumn,
                                textAlternateAllStyleId, textAlternateDAllStyleId, textAlternateDLRTStyleId, textAlternateLRBStyleId, textAlternateLRStyleId, textAlternateLRTStyleId);
                        }
                        cell = GetCell(sheet, TabularPriceColumn, (uint)i + startRow);
                        SetCellValue(cell, GetColumnValue(row, "PriceWholesale"), 2, doubleAlternateStyleId);

                        foreach (string ordercol in TabularSKUOrderColumnLists)
                        {
                            cell = GetCell(sheet, ordercol, (uint)i + startRow);
                            SetCellValue(cell, string.Empty, 3, intAlternateLStyleId);
                        }
                    }

                    preDept = deptId;
                    preStyle = style;
                    preColor = color;

                    //total funcational
                    cell = GetCell(sheet, TabularRowTotalColumn, (uint)i + startRow);
                    CellFormula cellformula = new CellFormula();
                    cellformula.Text = "SUM(" + TabularSKUOrderColumnLists[0] + (i + startRow).ToString() + ":" + TabularSKUOrderColumnLists[TabularSKUOrderColumnLists.Length - 1] + (i + startRow).ToString() + ")";
                    cell.Append(cellformula);
                    cell.CellValue = new CellValue(string.Empty);
                    SetCellValue(cell, string.Empty, 3, intStyleId);

                    //ProductLine_Style
                    if (lineSheet != null)
                    {
                        //ProductLine
                        cell = GetCell(lineSheet, "A", (uint)i + startRow);
                        var productLine = GetColumnValue(row, "ProductLine");
                        SetCellValue(cell, productLine, 1, textNullStyleId);
                        //Style
                        cell = GetCell(lineSheet, "B", (uint)i + startRow);
                        SetCellValue(cell, style, 1, textNullStyleId);
                        //set category
                        if (CategoryList.Count < 12 && !string.IsNullOrEmpty(productLine))
                        {
                            var results = from c in CategoryList where c.Category == productLine select c;
                            if (results.Count() <= 0)
                                CategoryList.Add(new CategoryInfo() { Category = productLine });
                        }
                    }

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
                            multiple = new OrderMultipleDataValidation { Multiple = OrderMultiple, SequenceOfReferences = TabularSKUOrderColumnLists[0] + (i + startRow).ToString() + ":" + TabularSKUOrderColumnLists[TabularSKUOrderColumnLists.Length - 1] + (i + startRow).ToString() };
                            multipleList.Add(multiple);
                        }
                        else
                        {
                            multiple = multips.First();
                            multiple.SequenceOfReferences = multiple.SequenceOfReferences + " " + TabularSKUOrderColumnLists[0] + (i + startRow).ToString() + ":" + TabularSKUOrderColumnLists[TabularSKUOrderColumnLists.Length - 1] + (i + startRow).ToString();
                        }
                    }
                }
                //Set Department Column Width
                /*var maxDeptName = string.Empty;
                foreach (DataRow row in formData.Rows)
                {
                    var deptName = GetColumnValue(row, "NavigationBar");
                    if (maxDeptName.Length < deptName.Length)
                        maxDeptName = deptName;
                }
                var deptWidth = GetWidth(maxDeptName);
                deptWidth = maxDeptWidth > 0 ? (deptWidth > maxDeptWidth ? maxDeptWidth : deptWidth) : deptWidth;
                var columns = sheet.GetFirstChild<Columns>().Elements<Column>();
                var column = (from c in columns where c.Max == 4 && c.Min == 4 select c).First();
                column.BestFit = true;
                column.CustomWidth = true;
                column.Width = new DoubleValue(deptWidth);*/
                if (TabularCancelDate == "0")
                {
                    var rows = sheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == TabularCancelDateRow);
                    if (rows.Count() > 0)
                        rows.First().Hidden = true;
                }
                if (TabularShipMethods == "0")
                {
                    var rows = sheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == TabularShipMethodsRow);
                    if (rows.Count() > 0)
                        rows.First().Hidden = true;
                }

                //Generate Category cells
                if (lineSheet != null && CategoryList.Count > 0)
                {
                    var TabularCategoryStartRow = 2;
                    var TabularCategory1Column = ConfigurationManager.AppSettings["TabularCategory1Column"] == null ? "J" : ConfigurationManager.AppSettings["TabularCategory1Column"];
                    var TabularCategory2Column = ConfigurationManager.AppSettings["TabularCategory2Column"] == null ? "M" : ConfigurationManager.AppSettings["TabularCategory2Column"];
                    var CategoryDataFormat = string.IsNullOrEmpty(ConfigurationManager.AppSettings["TabularCategoryDoubleFormat"]) ? intStyleId :
                        (ConfigurationManager.AppSettings["TabularCategoryDoubleFormat"] == "1" ? (UInt32Value)TabularCommonDoubleFormat : intStyleId);
                    if (CategoryList.Count > 10)
                        TabularCategoryStartRow = 1;
                    var category1Number = CategoryList.Count / 2 + CategoryList.Count % 2;
                    for (int i = 0; i < CategoryList.Count; i++)
                    {
                        if (i < category1Number)
                        {
                            CategoryList[i].Row = TabularCategoryStartRow + i;
                            CategoryList[i].Column = TabularCategory1Column;
                        }
                        else
                        {
                            CategoryList[i].Row = TabularCategoryStartRow + i - category1Number;
                            CategoryList[i].Column = TabularCategory2Column;
                        }
                        var cell = GetCell(sheet, CategoryList[i].Column, (uint)CategoryList[i].Row);
                        SetCellValue(cell, CategoryList[i].Category, 1, textAllStyleId);

                        CategoryList[i].Column = ((char)(char.Parse(CategoryList[i].Column) + 1)).ToString();
                        cell = GetCell(sheet, CategoryList[i].Column, (uint)CategoryList[i].Row);
                        cell.CellValue = new CellValue(string.Empty);
                        SetCellValue(cell, string.Empty, 3, CategoryDataFormat);
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
                        validation.SequenceOfReferences.InnerText = TabularSKUOrderColumnLists[0] + startRow.ToString() + ":" + TabularSKUOrderColumnLists[TabularSKUOrderColumnLists.Length - 1] + (startRow + formData.Rows.Count - 1).ToString();
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

                if (windowDt != null && windowDt.Rows.Count > 0)
                {
                    var TabularWindowsList = ConfigurationManager.AppSettings["TabularWindowsList"] == null ? "E|F|G" : ConfigurationManager.AppSettings["TabularWindowsList"];
                    var winlist = TabularWindowsList.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    var currentIndex = 11;
                    var TabularWindownNameRightAlign = ConfigurationManager.AppSettings["TabularWindownNameRightAlign"] == null ? "0" : ConfigurationManager.AppSettings["TabularWindownNameRightAlign"];
                    var textNullRStyleId = textNullStyleId;
                    if(TabularWindownNameRightAlign== "1")
                        textNullRStyleId = CreateCellFormat(styleSheet, fontId, null, null, UInt32Value.FromUInt32(49), true, true, HorizontalAlignmentValues.Right);
                    foreach (DataRow row in windowDt.Rows)
                    {
                        var windowName = row.IsNull("WindowName") ? string.Empty : row["WindowName"].ToString();
                        if (!string.IsNullOrEmpty(windowName))
                        {
                            var cell = GetCell(sheet, winlist[0], (uint)currentIndex);
                            SetCellValue(cell, windowName, 1, textNullRStyleId);
                            cell = GetCell(sheet, winlist[1], (uint)currentIndex);
                            SetCellValue(cell, "0", 2, TabularCommonDoubleFormat);
                            cell = GetCell(sheet, winlist[2], (uint)currentIndex);
                            SetCellValue(cell, CurrencyCode, 1, textNullStyleId);
                            currentIndex--;
                        }
                    }
                }

                sheet.Save();
                if (lineSheet != null)
                    lineSheet.Save();
            }
            book.Close();
        }

        private void LoadTabularDataToCell_old(string columnInfo, Worksheet sheet, uint rowIndex, DataRow row,
            string deptId, string preDept, string style, string preStyle, string color, string preColor, int loopIndex, int totalCount, string skuColumn,
            UInt32Value textAllStyleId, UInt32Value textDAllStyleId, UInt32Value textDLRTStyleId, UInt32Value textLRBStyleId, UInt32Value textLRStyleId, UInt32Value textLRTStyleId)
        {
            Cell cell = null;
            var colInfo = columnInfo.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            cell = GetCell(sheet, colInfo[0], rowIndex);
            switch (colInfo[1])
            {
                case "Catalog":
                    SetCellValue(cell, GetColumnValue(row, "CatalogCode"), 1, textAllStyleId);
                    break;
                case "SKU":
                    SetCellValue(cell, GetColumnValue(row, "SKU"), 1, textAllStyleId);
                    break;
                case "UPC":
                    SetCellValue(cell, GetColumnValue(row, "UPC"), 1, textAllStyleId);
                    break;
                case "Department":
                    if (deptId != preDept)
                    {
                        if (loopIndex == totalCount - 1)
                            SetCellValue(cell, GetColumnValue(row, "NavigationBar"), 1, textDAllStyleId);
                        else
                            SetCellValue(cell, GetColumnValue(row, "NavigationBar"), 1, textDLRTStyleId);//DepartmentName
                    }
                    else
                    {
                        if (loopIndex == totalCount - 1)
                            SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                        else
                            SetCellValue(cell, string.Empty, 1, textLRStyleId);
                    }
                    break;
                case "Style":
                    if ((style != preStyle) || (style == preStyle && deptId != preDept))
                    {
                        if (loopIndex == totalCount - 1)
                            SetCellValue(cell, style, 1, textAllStyleId);
                        else
                            SetCellValue(cell, style, 1, textLRTStyleId);//DepartmentName
                    }
                    else
                    {
                        if (loopIndex == totalCount - 1)
                            SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                        else
                            SetCellValue(cell, string.Empty, 1, textLRStyleId);
                    }
                    break;
                case "ProductName":
                    if (skuColumn == "F")
                    {
                        cell = GetCell(sheet, skuColumn, rowIndex);
                        SetCellValue(cell, GetColumnValue(row, "SKU"), 1, textAllStyleId);
                    }
                    else
                    {
                        var productName = GetColumnValue(row, "ProductName");
                        if ((style != preStyle) || (style == preStyle && deptId != preDept))
                        {
                            if (loopIndex == totalCount - 1)
                                SetCellValue(cell, productName, 1, textDAllStyleId);//textAllStyleId);
                            else
                                SetCellValue(cell, productName, 1, textDLRTStyleId);//textLRTStyleId);
                        }
                        else
                        {
                            if (loopIndex == totalCount - 1)
                                SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                            else
                                SetCellValue(cell, string.Empty, 1, textLRStyleId);
                        }
                    }
                    break;
                case "AttributeValue2":
                    if ((color != preColor) || (color == preColor && deptId != preDept) || (color == preColor && style != preStyle))
                    {
                        if (loopIndex == totalCount - 1)
                            SetCellValue(cell, color, 1, textAllStyleId);
                        else
                            SetCellValue(cell, color, 1, textLRTStyleId);//DepartmentName
                    }
                    else
                    {
                        if (loopIndex == totalCount - 1)
                            SetCellValue(cell, string.Empty, 1, textLRBStyleId);
                        else
                            SetCellValue(cell, string.Empty, 1, textLRStyleId);
                    }
                    break;
                case "AttributeValue1":
                    SetCellValue(cell, GetColumnValue(row, "AttributeValue1"), 1, textAllStyleId);
                    break;
                default:
                    SetCellValue(cell, GetColumnValue(row, colInfo[1]), 1, textAllStyleId);
                    break;
            }
        }

        private void SaveTabularCustomerDataToSheet_old(DataTable customerData, string filePath)
        {
            var sheetName = ConfigurationManager.AppSettings["CustomerSheetName"] == null ? "SoldToShipTo" : ConfigurationManager.AppSettings["CustomerSheetName"];
            var startColumn = ConfigurationManager.AppSettings["CustomerStartColumn"] == null ? "A" : ConfigurationManager.AppSettings["CustomerStartColumn"];
            var columnLength = ConfigurationManager.AppSettings["CustomerColumnLength"] == null ? 16 : int.Parse(ConfigurationManager.AppSettings["CustomerColumnLength"]);

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
                var columnList = new List<string>();
                for (int i = 0; i < columnLength; i++)
                {
                    var col = ((char)(char.Parse(startColumn) + i)).ToString();
                    var cell = GetCell(sheet, col, 1);
                    if (cell != null && cell.CellValue != null && !string.IsNullOrEmpty(cell.CellValue.Text))
                        columnList.Add(GetCellValue(book, sheet, cell));
                    else
                        columnList.Add(string.Empty);
                }
                for (int i = 0; i < customerData.Rows.Count; i++)
                {
                    var row = customerData.Rows[i];
                    for (int j = 0; j < columnLength; j++)
                    {
                        var col = ((char)(char.Parse(startColumn) + j)).ToString();
                        var cell = GetCell(sheet, col, (uint)(i + 2));
                        //cell.CellValue = new CellValue(GetColumnValue(row, columnList[j]));
                        SetCellValue(cell, GetColumnValue(row, columnList[j]), 1, TabularCommonTextFormat);
                    }
                }
                sheet.Save();
            }
            book.Close();
        }

        private void SaveTabularShipMethodsDataToSheet_old(DataTable methodsData, string filePath)
        {
            var sheetName = ConfigurationManager.AppSettings["TabularShipMethodsSheetName"] == null ? "ShipMethods" : ConfigurationManager.AppSettings["TabularShipMethodsSheetName"];

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
                for (int i = 0; i < methodsData.Rows.Count; i++)
                {
                    var row = methodsData.Rows[i];
                    var cell = GetCell(sheet, "A", (uint)(i + 1));
                    SetCellValue(cell, GetColumnValue(row, "code"), 1, TabularCommonTextFormat);
                    cell = GetCell(sheet, "B", (uint)(i + 1));
                    SetCellValue(cell, GetColumnValue(row, "description"), 1, TabularCommonTextFormat);
                }
                sheet.Save();
            }
            book.Close();
        }

        private void SaveTabularDateValidationToSheet_old(DataRow catalogRow, string filePath, int SKURowNumber, DataTable windowDt)
        {
            var sheetName = ConfigurationManager.AppSettings["GlobalSheetName"] == null ? "GlobalVariables" : ConfigurationManager.AppSettings["GlobalSheetName"];
            var startRow = ConfigurationManager.AppSettings["TabularSKUStartRow"] == null ? 16 : int.Parse(ConfigurationManager.AppSettings["TabularSKUStartRow"]);

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
                SetCellValue(cell, (string.IsNullOrEmpty(bookingOrder) ? "0" : bookingOrder), 1, TabularCommonDateFormat);
                //DateShipStart
                if (!catalogRow.IsNull("DateShipStart"))
                {
                    cell = GetCell(sheet, "B", 4);
                    var date = ((DateTime)catalogRow["DateShipStart"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);//1033
                    SetCellValue(cell, date, 4, TabularCommonDateFormat);
                }
                //DateShipEnd
                if (!catalogRow.IsNull("DateShipEnd"))
                {
                    cell = GetCell(sheet, "B", 5);
                    var date = ((DateTime)catalogRow["DateShipEnd"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);//1033
                    SetCellValue(cell, date, 4, TabularCommonDateFormat);
                }
                //ReqDateDays
                var ReqDateDays = (int)catalogRow["ReqDateDays"];
                cell = GetCell(sheet, "B", 6);
                SetCellValue(cell, ReqDateDays.ToString(), 1, TabularCommonTextFormat);
                //ReqDate
                //if (ReqDateDays > -1)
                //{
                //    cell = GetCell(sheet, "B", 7);
                //    var reqDate = DateTime.Today.AddDays(ReqDateDays).ToString((new System.Globalization.CultureInfo(1033)).DateTimeFormat.ShortDatePattern);
                //    SetCellValue(cell, reqDate, 4, TabularCommonDateFormat);
                //}
                //ReqDefaultDays
                var ReqDefaultDays = (int)catalogRow["DefaultReqDays"];
                cell = GetCell(sheet, "B", 7);
                SetCellValue(cell, ReqDefaultDays.ToString(), 1, TabularCommonTextFormat);
                //ReqDefaultDate
                //if (ReqDefaultDays > -1)
                //{
                //    cell = GetCell(sheet, "B", 9);
                //    var reqDate = DateTime.Today.AddDays(ReqDefaultDays).ToString((new System.Globalization.CultureInfo(1033)).DateTimeFormat.ShortDatePattern);
                //    SetCellValue(cell, reqDate, 4, TabularCommonDateFormat);
                //}
                //CancelDateDays
                var CancelDateDays = (int)catalogRow["CancelDateDays"];
                cell = GetCell(sheet, "B", 8);
                SetCellValue(cell, CancelDateDays.ToString(), 1, TabularCommonTextFormat);
                //CancelDefaultDays
                var CancelDefaultDays = (int)catalogRow["CancelDefaultDays"];
                cell = GetCell(sheet, "B", 9);
                SetCellValue(cell, CancelDefaultDays.ToString(), 1, TabularCommonTextFormat);

                cell = GetCell(sheet, "B", 10);
                SetCellValue(cell, (SKURowNumber + startRow - 1).ToString(), 1, TabularCommonTextFormat);

                cell = GetCell(sheet, "A", 11);
                SetCellValue(cell, "Categories", 1, TabularCommonTextFormat);
                //set for category configuration
                for (int i = 0; i < CategoryList.Count; i++)
                {
                    cell = GetCell(sheet, "B", 11 + (uint)i);
                    SetCellValue(cell, CategoryList[i].Category, 1, TabularCommonTextFormat);
                    cell = GetCell(sheet, "C", 11 + (uint)i);
                    SetCellValue(cell, CategoryList[i].Column, 1, TabularCommonTextFormat);
                    cell = GetCell(sheet, "D", 11 + (uint)i);
                    SetCellValue(cell, CategoryList[i].Row.ToString(), 1, TabularCommonTextFormat);
                }

                if (windowDt != null && windowDt.Rows.Count > 0)
                {
                    var currentIndex = 3;
                    foreach (DataRow row in windowDt.Rows)
                    {
                        var windowName = row.IsNull("WindowName") ? string.Empty : row["WindowName"].ToString();
                        if (!string.IsNullOrEmpty(windowName))
                        {
                            cell = GetCell(sheet, "E", (uint)currentIndex);
                            SetCellValue(cell, windowName, 1, TabularCommonTextFormat);
                            var date = ((DateTime)row["DateBegin"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);
                            cell = GetCell(sheet, "F", (uint)currentIndex);
                            SetCellValue(cell, date, 4, TabularCommonDateFormat);
                            date = ((DateTime)row["DateEnd"]).ToString((new System.Globalization.CultureInfo(2052)).DateTimeFormat.ShortDatePattern);
                            cell = GetCell(sheet, "G", (uint)currentIndex);
                            SetCellValue(cell, date, 4, TabularCommonDateFormat);
                            currentIndex++;
                        }
                    }
                }

            }

            book.WorkbookPart.Workbook.Save();
            book.Close();
        }

        private void ProtectWorkbook_old(string filename, string savedirectory)
        {
            var pwd = ConfigurationManager.AppSettings["ProtectPassword"] == null ? "Plumriver" : ConfigurationManager.AppSettings["ProtectPassword"];
            var password = HashPassword(pwd);
            var filePath = Path.Combine(savedirectory, filename);
            var sheetName = ConfigurationManager.AppSettings["TabularSheetName"] == null ? "TabularOfflineOrderForm" : ConfigurationManager.AppSettings["TabularSheetName"];
            Worksheet sheet = null;
            SpreadsheetDocument book = SpreadsheetDocument.Open(filePath, true);
            var sheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
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
                    sheetProtection = new SheetProtection() { Sheet = true, Objects = false, Scenarios = false }; //, DeleteRows = false, DeleteColumns = false, FormatCells = false, FormatColumns = false, FormatRows = false, InsertColumns = false, InsertRows = false, SelectLockedCells = false};
                    addNew = true;
                }
                else
                {
                    sheetProtection.Sheet = true;
                    sheetProtection.Objects = false;
                    sheetProtection.Scenarios = false;
                }
                //{ Sheet = true, Objects = true, Scenarios = true };
                sheetProtection.Password = password;
                if (addNew)
                    sheet.InsertAfter(sheetProtection, sheet.Descendants<SheetData>().LastOrDefault());
            }

            book.WorkbookPart.Workbook.WorkbookProtection = new WorkbookProtection() { LockStructure = true };
            ////book.WorkbookPart.Workbook.WorkbookProtection.LockWindows = true;
            //book.WorkbookPart.Workbook.WorkbookProtection.WorkbookAlgorithmName = "SHA-1";
            //book.WorkbookPart.Workbook.WorkbookProtection.WorkbookPassword = password;//new HexBinaryValue() { Value = password };

            sheetName = ConfigurationManager.AppSettings["GlobalSheetName"] == null ? "GlobalVariables" : ConfigurationManager.AppSettings["GlobalSheetName"];
            var pwdColumn = ConfigurationManager.AppSettings["GlobalPWDColumn"] == null ? "B" : ConfigurationManager.AppSettings["GlobalPWDColumn"];
            var pwdRow = ConfigurationManager.AppSettings["GlobalPWDRow"] == null ? 200 : int.Parse(ConfigurationManager.AppSettings["GlobalPWDRow"]);

            sheet = null;
            sheets = book.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() > 0)
            {
                WorksheetPart worksheetPart = (WorksheetPart)book.WorkbookPart.GetPartById(sheets.First().Id.Value);
                sheet = worksheetPart.Worksheet;
            }
            else
                WriteToLog("Wrong template order form file, cannot find sheet " + sheetName + ", file path:" + filePath);
            if (sheet != null)
            {
                var cell = GetCell(sheet, pwdColumn, (uint)(pwdRow));
                SetCellValue(cell, pwd, 1, TabularCommonTextFormat);
            }

            book.WorkbookPart.Workbook.Save();
            book.Close();
        }

        


        /*************************SUPPORT METHODS**************************/

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
            DocumentFormat.OpenXml.Spreadsheet.FontSize newFontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 9 };
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
            var borderContent = new DocumentFormat.OpenXml.Spreadsheet.Border();
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
            return CreateCellFormat(styleSheet, fontIndex, fillIndex, borderIndex, numberFormatId, true, false, HorizontalAlignmentValues.General);
        }

        private UInt32Value CreateCellFormat(Stylesheet styleSheet, UInt32Value fontIndex, UInt32Value fillIndex, UInt32Value borderIndex, UInt32Value numberFormatId, bool locked, bool alignment, HorizontalAlignmentValues horizontal)
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

            if (alignment)
            {
                cellFormat.ApplyAlignment = true;
                cellFormat.Alignment = new Alignment();
                cellFormat.Alignment.WrapText = true;
                cellFormat.Alignment.Horizontal = horizontal;
            }
            styleSheet.CellFormats.Append(cellFormat);

            UInt32Value result = styleSheet.CellFormats.Count;
            styleSheet.CellFormats.Count++;
            return result;
        }

        private void SetCellValue(Cell cell, string value, int format, UInt32Value styleId)
        {
            switch (format)
            {
                case 1:
                    cell.StyleIndex = styleId;
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    break;
                case 2:
                    cell.StyleIndex = styleId;
                    cell.CellValue = new CellValue(string.Format("{0:N2}", value));
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case 3:
                    cell.StyleIndex = styleId;
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case 4:
                    cell.StyleIndex = styleId;
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    break;
            }
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
        }


        private void WriteToLog(string msg, Exception e, string savedirectory)
        {
            if(string.IsNullOrEmpty(savedirectory))
                savedirectory = ConfigurationManager.AppSettings["savefolder"].ToString();
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

    }

    public class CategoryInfo
    {
        public string Category { get; set; }
        public string Column { get; set; }
        public int Row { get; set; }
    }

    public class OrderMultipleDataValidation
    {
        public int Multiple { get; set; }
        public string SequenceOfReferences { get; set; }
    }
}
