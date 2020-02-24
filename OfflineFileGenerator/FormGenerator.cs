namespace OfflineFileGenerator
{
    using System;
    using System.IO;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;
    using System.Data;
    using System.Data.SqlClient;
    using System.Collections;
    using System.Collections.Specialized;
    using System.Configuration;
    using System.Collections.Generic;
    using System.Xml.Linq;
    using System.Linq;
    using System.Linq.Expressions;
    //using CarlosAg.ExcelXmlWriter;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.SS.Util;

    public partial class App
    {
        public bool IncludePricing = true;
        public bool OrderShipDate = false;

        DataTable ATPValueData;

        //Hashtable variants used to store catalog data
        ListDictionary departments = new ListDictionary();
        ListDictionary cols = new ListDictionary();
        OrderedDictionary multiples = new OrderedDictionary();
        OrderedDictionary locked = new OrderedDictionary();
        OrderedDictionary unlocked = new OrderedDictionary();
        int datacolumnwidth = Convert.ToInt32(ConfigurationManager.AppSettings["datacolumnwidth"].ToString());
        int dataColumnNumber = ConfigurationManager.AppSettings.Get("datacolumnnumber") == null ? 30 : Convert.ToInt32(ConfigurationManager.AppSettings["datacolumnnumber"].ToString());
        bool multicolumndeptlabel = (ConfigurationManager.AppSettings["multicolumndeptlabel"].ToString().ToUpper() == "TRUE" ? true : false);
        int maxcols = 0;    //The maxium depth of the department hierarchy
        string ExcludePriceValue = ConfigurationManager.AppSettings.Get("ExcludePriceValue") == null ? "0" : ConfigurationManager.AppSettings["ExcludePriceValue"];

        bool TopDeptGroupStyle = (ConfigurationManager.AppSettings["TopDeptGroupStyle"] ?? "0") == "1";

        bool allThinBorder = ConfigurationManager.AppSettings["ThinBorder"] == null ? false : ConfigurationManager.AppSettings["ThinBorder"].ToString() == "1" ? true : false;
        bool FontBold = ConfigurationManager.AppSettings["FontBold"] == null ? true : ConfigurationManager.AppSettings["FontBold"].ToString() == "1" ? true : false;
        string fontName = ConfigurationManager.AppSettings["FontName"] ?? string.Empty;
        short FontSize = short.Parse( ConfigurationManager.AppSettings["FontSize"] ?? "8");
        float RowHeight = float.Parse(ConfigurationManager.AppSettings["RowHeight"] ?? "12");
        short TitleBGColor = short.Parse(ConfigurationManager.AppSettings["TitleBGColor"] ?? "22");
        short LockedCellBGColor = short.Parse(ConfigurationManager.AppSettings["LockedCellBGColor"] ?? "22");
        bool DeptDesciptionBold = (ConfigurationManager.AppSettings["DeptDesciptionBold"] ?? "0") == "1";
        bool TitleBorder = (ConfigurationManager.AppSettings["TitleBorder"] ?? "0") == "1";
        bool TotalSection = (ConfigurationManager.AppSettings["TotalSection"] ?? "0") == "1";
        int TotalAmtCol = 0, TotalQtyCol = 0;
        StringBuilder TotalAmtfn = new StringBuilder();
        StringBuilder TotalQtyfn = new StringBuilder();

        public Hashtable StyleCollection = new Hashtable();

        public string CatalogName = string.Empty;

        public int GenerateData(string filename, string soldto, string catalog, string pricecode, string savedirectory)
        {
            //Make sure data doesn't carry across multiple catalogs
            departments.Clear();
            cols.Clear();
            multiples.Clear();
            locked.Clear();
            unlocked.Clear();

            HSSFWorkbook book = new HSSFWorkbook();
            // -----------------------------------------------
            //  Generate Styles
            // -----------------------------------------------
            //this.GenerateStyles(book);
            // -----------------------------------------------
            //  Generate Order Template Worksheet
            // -----------------------------------------------
            maxcols = this.GenerateWorksheetOrderTemplate(book, soldto, catalog, pricecode, savedirectory);
            if (maxcols > -1)
            {
                //This saves the XML version of the generated spreadsheet for use by the next
                //step in the process
                FileStream file = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                book.Write(file);
                file.Close();
                return maxcols;
            }
            else
                return -1;
            
        }

        private HSSFCellStyle SetCellStyle(HSSFWorkbook book, string cellStyleName, bool bFont, FontBoldWeight? fontWeight, short? fontHeight, VerticalAlignment? valign,
            HorizontalAlignment? align, CellBorderType? borderBottom, CellBorderType? borderLeft, CellBorderType? borderRight, CellBorderType? borderTop, 
            short? bgColor, string dataFormat)
        {
            if(StyleCollection.ContainsKey(cellStyleName))
                return (HSSFCellStyle)StyleCollection[cellStyleName];
            HSSFCellStyle cellStyle = (HSSFCellStyle)book.CreateCellStyle();
            if (bFont)
            {
                HSSFFont font = (HSSFFont)book.CreateFont();
                if (!string.IsNullOrEmpty(fontName))
                    font.FontName = fontName;
                font.Boldweight = fontWeight == null ? (short)FontBoldWeight.NORMAL : (short)fontWeight;
                font.FontHeightInPoints = fontHeight == null ? (short)8 : (short)fontHeight;
                cellStyle.SetFont(font);
            }
            if(align != null)
                cellStyle.Alignment = (HorizontalAlignment)align;
            if(valign != null)
                cellStyle.VerticalAlignment = (VerticalAlignment)valign;
            if(borderBottom.HasValue)
                cellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : borderBottom.Value;
            if (borderLeft.HasValue)
                cellStyle.BorderLeft = allThinBorder ? CellBorderType.THIN : borderLeft.Value;
            if (borderRight.HasValue)
                cellStyle.BorderRight = allThinBorder ? CellBorderType.THIN : borderRight.Value;
            if (borderTop.HasValue)
                cellStyle.BorderTop = allThinBorder ? CellBorderType.THIN : borderTop.Value;
            if (bgColor != null)
            {
                cellStyle.FillForegroundColor = (short)bgColor;
                cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
            }
            if (!string.IsNullOrEmpty(dataFormat))
            {
                HSSFDataFormat format = (HSSFDataFormat)book.CreateDataFormat();
                cellStyle.DataFormat = format.GetFormat(dataFormat);
            }
            if(!StyleCollection.ContainsKey(cellStyleName))
                StyleCollection.Add(cellStyleName, cellStyle);
            return cellStyle;
        }

        private HSSFCellStyle GenerateStyle(HSSFWorkbook book, string styleName)
        {
            HSSFCellStyle cellStyle = null;
            switch (styleName)
            {
                case "H1":
                    cellStyle = SetCellStyle(book, "H1", true, FontBoldWeight.BOLD, 18, VerticalAlignment.BOTTOM, HorizontalAlignment.LEFT,
                    null, null, null, null, null, null);
                    break;
                case "18pt":
                    cellStyle = SetCellStyle(book, "18pt", true, FontBoldWeight.NORMAL, 18, VerticalAlignment.BOTTOM, HorizontalAlignment.LEFT,
                    null, null, null, null, null, null);
                    break;
                case "12pt":
                    cellStyle = SetCellStyle(book, "12pt", true, FontBoldWeight.NORMAL, 12, VerticalAlignment.BOTTOM, HorizontalAlignment.LEFT,
                    null, null, null, null, null, null);
                    break;
                case "Bold":
                    cellStyle = SetCellStyle(book, "Bold", true, FontBoldWeight.BOLD, 12, VerticalAlignment.BOTTOM, HorizontalAlignment.LEFT,
                    null, null, null, null, null, null);
                    break;

                case "Underline":
                    cellStyle = SetCellStyle(book, "Underline", true, FontBoldWeight.NORMAL, 12, VerticalAlignment.BOTTOM, HorizontalAlignment.JUSTIFY,
                    CellBorderType.THIN, null, null, null, null, null);
                    break;
                
                case "m19009468":
                    cellStyle = SetCellStyle(book, "m19009468", true, FontBoldWeight.BOLD, FontSize, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                    CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, null);
                    break;

                case "m19009478":
                    cellStyle = SetCellStyle(book, "m19009478", true, FontBoldWeight.BOLD, FontSize, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                        CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, null, null);
                    break;
                case "s26":
                    cellStyle = SetCellStyle(book, "s26", true, FontBold ? FontBoldWeight.BOLD : FontBoldWeight.NORMAL, FontSize, null, null,
                        CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, TitleBGColor, null);
                    break;
                case "s26Bold":
                    cellStyle = SetCellStyle(book, "s26Bold", true, FontBoldWeight.BOLD, FontSize, null, null,
                        CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, TitleBGColor, null);
                    break;
                case "s27":
                    cellStyle = SetCellStyle(book, "s27", true, FontBold ? FontBoldWeight.BOLD : FontBoldWeight.NORMAL, FontSize, null, HorizontalAlignment.RIGHT,
                CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, TitleBGColor, "#,##0\\ [$kr-41D]");
                    break;
                case "s28":
                    cellStyle = SetCellStyle(book, "s28", true, FontBold ? FontBoldWeight.BOLD : FontBoldWeight.NORMAL, FontSize, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, TitleBGColor, null);
                    break;
                case "s29":
                    cellStyle = SetCellStyle(book, "s29", true, FontBold ? FontBoldWeight.BOLD : FontBoldWeight.NORMAL, FontSize, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                        CellBorderType.MEDIUM, null, TitleBorder?CellBorderType.MEDIUM: CellBorderType.NONE , CellBorderType.MEDIUM, TitleBGColor, null);
                    cellStyle.WrapText = true;
                    break;
                case "s43":
                    cellStyle = cellStyle = SetCellStyle(book, "s43", true, FontBoldWeight.NORMAL, FontSize, null, null,
                null, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.THIN, null, null); ;
                    break;
                case "s24":
                    cellStyle = SetCellStyle(book, "s24", true, FontBoldWeight.NORMAL, FontSize, null, null,
                null, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.THIN, null, null); ;
                    break;
                case "s38":
                    cellStyle = SetCellStyle(book, "s38", true, FontBoldWeight.NORMAL, FontSize, null, null,
                null, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, null); ;
                    break;
                case "s40":
                    cellStyle = cellStyle = SetCellStyle(book, "s40", true, FontBoldWeight.NORMAL, FontSize, null, HorizontalAlignment.RIGHT,
                CellBorderType.THIN, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, "0.00"); ;
                    break;
                case "s41":
                    cellStyle = cellStyle = SetCellStyle(book, "s41", true, FontBold ? FontBoldWeight.BOLD : FontBoldWeight.NORMAL, FontSize, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, null); ;
                    break;
                case "s81":
                    cellStyle = cellStyle = SetCellStyle(book, "s81", true, FontBoldWeight.NORMAL, FontSize, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                CellBorderType.THIN, CellBorderType.THIN, CellBorderType.THIN, CellBorderType.MEDIUM, LockedCellBGColor, null); ;
                    break;
                case "s62":
                    cellStyle = cellStyle = SetCellStyle(book, "s62", true, FontBoldWeight.NORMAL, FontSize, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                CellBorderType.THIN, CellBorderType.THIN, CellBorderType.THIN, CellBorderType.MEDIUM, null, null); ;
                    break;
            }
            return cellStyle;
        }

        private int GenerateWorksheetOrderTemplate(HSSFWorkbook book, string soldto, string catalog, string pricecode, string savedirectory)
        {
            //Master proc to gather and layout the data

            var colMultiple = 50;
            //Pull values out of the config file
            int column1width = Convert.ToInt32(ConfigurationManager.AppSettings["column1width"].ToString()) * colMultiple;
            int column2width = Convert.ToInt32(ConfigurationManager.AppSettings["column2width"].ToString()) * colMultiple;
            int column3width = Convert.ToInt32(ConfigurationManager.AppSettings["column3width"].ToString()) * colMultiple;
            int column4width = Convert.ToInt32(ConfigurationManager.AppSettings["column4width"].ToString()) * colMultiple;
            int column5width = Convert.ToInt32(ConfigurationManager.AppSettings["column5width"].ToString()) * colMultiple;
            int column6width = Convert.ToInt32(ConfigurationManager.AppSettings["column6width"].ToString()) * colMultiple;
            int column7width = Convert.ToInt32(ConfigurationManager.AppSettings["column7width"].ToString()) * colMultiple;
            int column8width = Convert.ToInt32(ConfigurationManager.AppSettings["column8width"].ToString()) * colMultiple;
            int column9width = Convert.ToInt32(ConfigurationManager.AppSettings["column9width"].ToString()) * colMultiple;
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            int gridatlevel = Convert.ToInt32(ConfigurationManager.AppSettings["gridatlevel"].ToString());

            //Get catalog data from the database
            //int iret = PrepareData(soldto, catalog, pricecode, savedirectory);
            //if (iret == -1)
            //    return iret;
            WriteToLog("PrepareAllProductData begin:" + soldto + "|" + catalog + "|" + pricecode + "|" + DateTime.Now.TimeOfDay);
            if (!PrepareAllProductData(soldto, catalog, pricecode, savedirectory)) return -1;
            WriteToLog("PrepareAllProductData end:" + DateTime.Now.TimeOfDay);

            XmlReader reader = GetDepartments2(catalog, pricecode);
            reader.Read();
            reader.MoveToNextAttribute();
            maxcols = Convert.ToInt32(reader.Value);

            //Main tab
            HSSFSheet sheet = (HSSFSheet)book.CreateSheet("Order Template");
            //SKU tab - hidden - used to create sku column in upload tab
            HSSFSheet skusheet = (HSSFSheet)book.CreateSheet("WebSKU");
            //Wholesale price tab - hidden - used by price calculation formulas
            HSSFSheet pricesheet = (HSSFSheet)book.CreateSheet("WholesalePrice");
            //Lists of cells to be locked and have order multiple restrictions placed
            HSSFSheet formatcellssheet = (HSSFSheet)book.CreateSheet("CellFormats");
            //ATPDATE tab
            HSSFSheet atpsheet = (HSSFSheet)book.CreateSheet("ATPDate");
            var upcSheetName = ConfigurationManager.AppSettings["CatalogUPCSheetName"];
            HSSFSheet upcsheet = string.IsNullOrEmpty(upcSheetName) ? null : (HSSFSheet)book.CreateSheet(upcSheetName);
            //

            string ExclusiveStyles = ConfigurationManager.AppSettings["ExclusiveStyles"]??"0";
            if (ExclusiveStyles == "1")
            {
                WriteToLog("Generate ExclusiveStyles begin: " + DateTime.Now.TimeOfDay);
                HSSFSheet exclusivestylessheet = (HSSFSheet)book.CreateSheet("ExclusiveStyles");
                AddExclusiveStylesHeader(exclusivestylessheet);
                WriteExclusiveStylesValues(exclusivestylessheet, catalog);
                WriteToLog("Generate ExclusiveStyles end: " + DateTime.Now.TimeOfDay);
            }

            AddFormatCellsHeader(formatcellssheet);
            AddCatalogUPCSheetHeader(upcsheet);
            //sheet.Protected = false;
            //sheet.Table.FullColumns = 1;
            //sheet.Table.FullRows = 1;
            int colctr = 0;
            if (multicolumndeptlabel)
            {
                for (colctr = 0; colctr < maxcols; colctr++)
                {
                    sheet.SetColumnWidth(colctr, column1width);
                    skusheet.SetColumnWidth(colctr, column1width);
                    pricesheet.SetColumnWidth(colctr, column1width);
                }
            }
            else
            {
                sheet.SetColumnWidth(colctr, column1width);
                skusheet.SetColumnWidth(colctr, column1width);
                pricesheet.SetColumnWidth(colctr, column1width);
            }
            colctr++;

            sheet.SetColumnWidth(colctr, column2width);
            skusheet.SetColumnWidth(colctr, column2width);
            pricesheet.SetColumnWidth(colctr, column2width);
            colctr++;

            sheet.SetColumnWidth(colctr, column3width);
            skusheet.SetColumnWidth(colctr, column3width);
            pricesheet.SetColumnWidth(colctr, column3width);
            colctr++;

            sheet.SetColumnWidth(colctr, column4width);
            skusheet.SetColumnWidth(colctr, column4width);
            pricesheet.SetColumnWidth(colctr, column4width);
            colctr++;

            sheet.SetColumnWidth(colctr, column5width);
            skusheet.SetColumnWidth(colctr, column5width);
            pricesheet.SetColumnWidth(colctr, column5width);
            colctr++;

            sheet.SetColumnWidth(colctr, column6width);
            skusheet.SetColumnWidth(colctr, column6width);
            pricesheet.SetColumnWidth(colctr, column6width);
            colctr++;

            sheet.SetColumnWidth(colctr, column7width);
            skusheet.SetColumnWidth(colctr, column7width);
            pricesheet.SetColumnWidth(colctr, column7width);
            colctr++;

            // -----------------------------------------------
            if (!TopDeptGroupStyle)
            {
                HSSFRow Row0 = (HSSFRow)sheet.CreateRow(0);
                Row0.HeightInPoints = RowHeight;
                HSSFCell cell = (HSSFCell)Row0.CreateCell(1);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "m19009468");
                cell.SetCellValue("Reqested Delivery date:");
                cell = (HSSFCell)Row0.CreateCell(2);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "m19009468");
                cell = (HSSFCell)Row0.CreateCell(3);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "m19009468");
                cell = (HSSFCell)Row0.CreateCell(4);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "m19009468");
                sheet.AddMergedRegion(new Region(0, 1, 0, 4));
                AddSecondaryRows(RowHeight, skusheet, pricesheet, atpsheet, 0);
                // -----------------------------------------------
                HSSFRow Row1 = (HSSFRow)sheet.CreateRow(1);
                Row1.HeightInPoints = RowHeight;
                cell = (HSSFCell)Row1.CreateCell(1);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "m19009478");
                cell.SetCellValue("Customer name:");
                cell = (HSSFCell)Row1.CreateCell(2);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "m19009478");
                cell = (HSSFCell)Row1.CreateCell(3);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "m19009478");
                cell = (HSSFCell)Row1.CreateCell(4);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "m19009478");
                sheet.AddMergedRegion(new Region(1, 1, 1, 4));
                AddSecondaryRows(RowHeight, skusheet, pricesheet, atpsheet, 1);
            }
            else
            {
                HSSFRow Row0 = (HSSFRow)sheet.CreateRow(0);
                Row0.HeightInPoints = RowHeight;
                HSSFCell cell = (HSSFCell)Row0.CreateCell(1);
                cell = (HSSFCell)Row0.CreateCell(2);
                cell = (HSSFCell)Row0.CreateCell(3);
                cell = (HSSFCell)Row0.CreateCell(4);
                AddSecondaryRows(RowHeight, skusheet, pricesheet, atpsheet, 0);
                // -----------------------------------------------
                HSSFRow Row1 = (HSSFRow)sheet.CreateRow(1);
                Row1.HeightInPoints = RowHeight;
                cell = (HSSFCell)Row1.CreateCell(1);
                cell = (HSSFCell)Row1.CreateCell(2);
                cell = (HSSFCell)Row1.CreateCell(3);
                cell = (HSSFCell)Row1.CreateCell(4);
                AddSecondaryRows(RowHeight, skusheet, pricesheet, atpsheet, 1);

                AddBreakRows(sheet, skusheet, pricesheet, atpsheet, 1);

                string[][] rows = new string[][]{ // ColIndex, text, Font, Locked
                    new string[]{"1", "26", "2,,H1,1"},
                    new string[]{"4", "20", "1,Sold-To *,Bold,1|2,,Underline,0"},
                    new string[]{"5", "20", "1,Ship-To *,Bold,1|2,,Underline,0"},
                    new string[]{"6", "20", "1,PO #,Bold,1|2,,Underline,0"},
                    new string[]{"7", "20", "1,Requested Del Date *,Bold,1|2,,Underline,0"},
                    new string[]{"8", "20", "1,Cancel Date *,Bold,1|2,,Underline,0"}
                };
                foreach (string[] arr in rows)
                {
                    int rownumber = int.Parse(arr[0]);
                    float rowHeight = float.Parse(arr[1]);
                    HSSFRow rw = (HSSFRow)sheet.CreateRow(rownumber);
                    rw.HeightInPoints = rowHeight;
                    string[] cellArr = arr[2].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string cl in cellArr)
                    {
                        string[] attrs = cl.Split(new char[] { ',' });
                        cell = (HSSFCell)rw.CreateCell(int.Parse(attrs[0]));
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, attrs[2]);
                        cell.CellStyle.Alignment = HorizontalAlignment.GENERAL;
                        if (attrs[3] == "0") cell.CellStyle.IsLocked = false;
                        if (attrs[1].Length > 0) cell.SetCellValue(attrs[1]);
                    }
                    AddSecondaryRows(RowHeight, skusheet, pricesheet, atpsheet, rownumber);
                }
            }
            // -----------------------------------------------
            #region newcode
            ListDictionary colpositions = null;
            var sheetRow = 2;
            WriteToLog("CreateGridRows2 begin:" + DateTime.Now.TimeOfDay);
            XElement xroot = XElement.Load(TemplateXMLFilePath);

            //if (!TopDeptGroupStyle)
            //    AddBreakRows(sheet, skusheet, pricesheet, atpsheet, 3);

            string lastRootDeptID = "";
            bool? groupdTitle = null;
            TotalQtyfn = new StringBuilder();
            TotalAmtfn = new StringBuilder();
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

                    if (rootDept.Count > 0 && TopDeptGroupStyle)
                    {
                        if (rootDept != null && lastRootDeptID != rootDept.FirstOrDefault().Key)
                        {
                            upcsheet = (HSSFSheet)book.CreateSheet(rootDept.FirstOrDefault().Value[0]);
                            AddCatalogUPCSheetHeader(upcsheet);

                            string[] arr = rootDept.FirstOrDefault().Value;
                            AddBreakRows(sheet, skusheet, pricesheet, atpsheet, 1);
                            AddTextRow(sheet, skusheet, pricesheet, atpsheet, 20, new string[][] { new string[] { arr[0], "18pt", "0", "2" } });
                            AddTextRow(sheet, skusheet, pricesheet, atpsheet, 15, new string[][] { new string[] { string.Format("{0}: {1}", arr[1], arr[2]),"12pt", "0", "2", "0" },
                                                                                               new string[] { "BLACKED OUT = NOT AVAILABLE","Bold", "3", "8" } });
                            PrepackList = arr[3];
                            groupdTitle = true;
                            lastRootDeptID = rootDept.FirstOrDefault().Key;
                        }
                        else
                        {
                            groupdTitle = false;
                        }
                    }

                    CreateGrid2(xroot, depts, deptid, sheet, skusheet, pricesheet, atpsheet, ref colpositions, soldto, catalog, groupdTitle, PrepackList);
                    if (colpositions == null)
                        WriteToLog("colpositions is null, will make generation failed.issue in p_DepartmentAttributeValue1");
                    int gId = sheet.LastRowNum + 2;
                    CreateGridRows2(depts, sheet, skusheet, pricesheet, atpsheet, upcsheet, deptid, deptname, ref colpositions);
                    TotalQtyfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalQtyCol.ToName(), gId, sheet.LastRowNum + 1);
                    TotalAmtfn.AppendFormat("+SUM({0}{1}:{0}{2})", TotalAmtCol.ToName(), gId, sheet.LastRowNum + 1);
                }
            }

            if (TotalSection)
            {
                var rowIndex = sheet.LastRowNum + 3;
                HSSFRow hdrRow = (HSSFRow)sheet.CreateRow(rowIndex);
                hdrRow.HeightInPoints = RowHeight;
                HSSFCell cell = (HSSFCell)hdrRow.CreateCell(0);
                cell = (HSSFCell)hdrRow.CreateCell(TotalAmtCol);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26Bold");
                cell.SetCellValue(ConfigurationManager.AppSettings["TotalQtyLabel"] ?? "TOTAL PAIRS");

                cell = (HSSFCell)hdrRow.CreateCell(TotalAmtCol + 1);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s41");
                cell.CellFormula = "0" + TotalQtyfn.ToString();
                cell = (HSSFCell)hdrRow.CreateCell(TotalAmtCol + 2);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s41");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, TotalAmtCol + 1, TotalAmtCol + 2));
                rowIndex++;

                hdrRow = (HSSFRow)sheet.CreateRow(rowIndex);
                hdrRow.HeightInPoints = RowHeight;
                cell = (HSSFCell)hdrRow.CreateCell(0);
                cell = (HSSFCell)hdrRow.CreateCell(TotalAmtCol);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26Bold");
                cell.SetCellValue(ConfigurationManager.AppSettings["TotalAmountLabel"] ?? "TOTAL COST");

                cell = (HSSFCell)hdrRow.CreateCell(TotalAmtCol + 1);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s40");
                cell.CellFormula = "0" + TotalAmtfn.ToString();
                cell = (HSSFCell)hdrRow.CreateCell(TotalAmtCol + 2);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s40");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, TotalAmtCol + 1, TotalAmtCol + 2));
            }

            WriteToLog("CreateGridRows2 end:" + DateTime.Now.TimeOfDay);
            sheetRow = sheet.LastRowNum + 1;
            HSSFRow row = (HSSFRow)sheet.CreateRow(sheetRow);
            HSSFCell eofCell = (HSSFCell)row.CreateCell(0); 
            eofCell.SetCellValue("EOF");
            sheetRow++;

            WriteCellFormatValues(formatcellssheet);
            for (int i = 0; i < dataColumnNumber; i++)
            {
                sheet.SetColumnWidth(colctr, datacolumnwidth * colMultiple);
                colctr++;
            }

            #endregion

            // -----------------------------------------------
            HSSFRow Row47 = (HSSFRow)sheet.CreateRow(sheetRow);
            // -----------------------------------------------
            //  Options
            // -----------------------------------------------
            //sheet.Options.Selected = true;
            //sheet.Options.FreezePanes = false;
            //sheet.Options.ProtectObjects = true;
            //sheet.Options.ProtectScenarios = true;
            //sheet.Options.Print.ValidPrinterInfo = true;

            return maxcols;
        }

        private void AddFormatCellsHeader(HSSFSheet sheet)
        {
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Unlocked Cells");
            row.CreateCell(1);
            row.CreateCell(2).SetCellValue("Multiple Cells");
            row.CreateCell(3).SetCellValue("Multiple Value");
        }

        private void AddCatalogUPCSheetHeader(HSSFSheet sheet)
        {
            if (sheet == null)
                return;
            var header = ConfigurationManager.AppSettings["CatalogUPCSheetHeader"];
            header = string.IsNullOrEmpty(header) ? "Style Name|Material #|Gender|Size|UPC Code|Purchase Price" : header;
            var headers = header.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);
            for (int i = 0; i < headers.Count(); i++)
            {
                row.CreateCell(i).SetCellValue(headers[i]);
            }
        }

        private void AddExclusiveStylesHeader(HSSFSheet sheet)
        {
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("SoldTo");
            row.CreateCell(1).SetCellValue("Styles");
        }

        private void AddSecondaryRows(float rowHeight, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, int rowIndex)
        {
            HSSFRow newrow = (HSSFRow)skusheet.CreateRow(rowIndex);
            newrow.HeightInPoints = rowHeight;
            HSSFRow newrowp = (HSSFRow)pricesheet.CreateRow(rowIndex);
            newrowp.HeightInPoints = rowHeight;
            if (OrderShipDate)
            {
                HSSFRow newrowa = (HSSFRow)atpsheet.CreateRow(rowIndex);
                newrowa.HeightInPoints = rowHeight;
            }
        }

        private void WriteCellFormatValues(HSSFSheet formatcellssheet)
        {
            int ctr = 1;
            var rowNum = formatcellssheet.LastRowNum + 1;
            //Write previously stored order multiple restrictions in the cellformat sheet
            if (multiples.Count > unlocked.Count)
            {
                foreach (DictionaryEntry de in multiples)
                {
                    HSSFRow row = (HSSFRow)formatcellssheet.CreateRow(rowNum);
                    row.CreateCell(0);
                    row.CreateCell(1);
                    row.CreateCell(2).SetCellValue(Convert.ToString(de.Key));
                    row.CreateCell(3).SetCellValue(Convert.ToString(de.Value));
                    rowNum++;
                }
                //Write previously stored unlocked cell inventory in the cellformat sheet
                foreach (DictionaryEntry de in unlocked)
                {
                    formatcellssheet.GetRow(ctr).GetCell(0).SetCellValue(Convert.ToString(de.Key));
                    ctr++;
                }
            }
            else
            {
                foreach (DictionaryEntry de in unlocked)
                {
                    HSSFRow row = (HSSFRow)formatcellssheet.CreateRow(rowNum);
                    row.CreateCell(0).SetCellValue(Convert.ToString(de.Key));
                    row.CreateCell(1);
                    row.CreateCell(2);
                    row.CreateCell(3);
                    rowNum++;
                }
                foreach (DictionaryEntry de in multiples)
                {
                    formatcellssheet.GetRow(ctr).GetCell(2).SetCellValue(Convert.ToString(de.Key));
                    formatcellssheet.GetRow(ctr).GetCell(3).SetCellValue(Convert.ToString(de.Value));
                    ctr++;
                }
            }
        }

        private void WriteExclusiveStylesValues(HSSFSheet exclusivestylessheet, string catalog)
        {
            DataTable dtExclusiveStyles = GetExclusiveStyles(catalog);
            var rowNum = exclusivestylessheet.LastRowNum + 1;
            if (dtExclusiveStyles != null)
            {
                foreach (DataRow dr in dtExclusiveStyles.Rows)
                {
                    HSSFRow row = (HSSFRow)exclusivestylessheet.CreateRow(rowNum);

                    row.CreateCell(0).SetCellValue(Convert.ToString(dr["SoldTo"]));
                    row.CreateCell(1).SetCellValue(Convert.ToString(dr["Styles"]));
                    rowNum++;
                }
            }
        }

        private void CreateGrid2(XElement xroot, OrderedDictionary depts, int deptid, HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, ref ListDictionary colpositions, string soldto, string catalog, 
            bool? groupedTitle, string PrepackList)
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
                        colpositions = AddHeader2(depts, deptcols, sheet, skusheet, pricesheet, atpsheet, true, PrepackList);
                    colpositions = AddHeader2(depts, deptcols, sheet, skusheet, pricesheet, atpsheet, false, PrepackList);
                }
                else
                {
                    colpositions = AddHeader2(depts, deptcols, sheet, skusheet, pricesheet, atpsheet, null, PrepackList);
                }
            }
        }

        private ListDictionary AddHeader2(OrderedDictionary depts, OrderedDictionary cols, HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet,
            bool? groupedTitle, string PrepackList)
        {
            bool topTitle = false; int PrepackAmt = 0; 
            string[] sizePrepack = PrepackList.Split(new char[]{','}, StringSplitOptions.RemoveEmptyEntries);
            if (this.TopDeptGroupStyle && groupedTitle.HasValue && groupedTitle.Value)
            {
                topTitle = true;
                if(sizePrepack.Length>0)
                {
                    string[] arr = sizePrepack[0].Split(new char[]{'|'}, StringSplitOptions.RemoveEmptyEntries);
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

            //Add a break between top-level departments
            var sheetRow = sheet.LastRowNum + 1;
            if (!TopDeptGroupStyle)
                sheetRow = AddBreakRows(sheet, skusheet, pricesheet, atpsheet, 3);

            HSSFRow hdrRow = (HSSFRow)sheet.CreateRow(sheetRow);
            hdrRow.HeightInPoints = RowHeight;
            AddSecondaryRows(RowHeight, skusheet, pricesheet, atpsheet, sheetRow);
            HSSFRow mrgRow = null;
            if (topTitle) 
            {
                mrgRow = (HSSFRow)sheet.CreateRow(sheetRow+1);
                mrgRow.HeightInPoints = RowHeight;
                AddSecondaryRows(RowHeight, skusheet, pricesheet, atpsheet, sheetRow);
            }
            sheetRow++;

            var cellIndex = 0;
            HSSFCell cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cellIndex++;
            if (!multicolumndeptlabel)   //Put the whole dept hierarchy in the first header cell
            {
                string deptlabel = "> ";
                foreach (DictionaryEntry de in depts)
                    deptlabel += de.Value + " > ";
                deptlabel = deptlabel.Substring(0, deptlabel.Length - 3);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, DeptDesciptionBold ? "s26Bold" : "s26");
                if (!groupedTitle.HasValue || groupedTitle == false)
                    cell.SetCellValue(deptlabel);
            }
            else   //Lay the dept hierarchy out w/each level in its own column
            {
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, DeptDesciptionBold ? "s26Bold" : "s26");
                cell.SetCellValue(column1heading);
                foreach (DictionaryEntry de in depts)
                {
                    if (!firstrow)
                    {
                        cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, DeptDesciptionBold ? "s26Bold" : "s26");
                        cell.SetCellValue("");
                        cellIndex++;
                    }
                    firstrow = false;
                }
                //Fill the rest of the dept column labels horizontally with blanks
                if (depts.Count < maxcols)
                {
                    for (int i = 0; i < maxcols - depts.Count; i++)
                    {
                        cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, DeptDesciptionBold ? "s26Bold" : "s26");
                        cell.SetCellValue("");
                        cellIndex++;
                    }
                }
            }
            if (this.TopDeptGroupStyle && !topTitle)
            {
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow - 1, 0, 2));
            }
            if (topTitle)
            {
                cell = (HSSFCell)mrgRow.CreateCell(cellIndex-1);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow, cellIndex-1, cellIndex-1));
            }
            else
                cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
            cell.SetCellValue(column4heading);
            if (topTitle)
            {
                cell = (HSSFCell)mrgRow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow, cellIndex, cellIndex));
            }
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
            cell.SetCellValue(column5heading);
            if (topTitle)
            {
                cell = (HSSFCell)mrgRow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow, cellIndex, cellIndex));
            }
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
            cell.SetCellValue(column6heading);
            if (topTitle)
            {
                cell = (HSSFCell)mrgRow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow, cellIndex, cellIndex));
            }
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s27");
            cell.SetCellValue(column7heading);
            if (topTitle)
            {
                cell = (HSSFCell)mrgRow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow, cellIndex, cellIndex));
            }
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s28");
            cell.SetCellValue(column8heading);
            if (topTitle)
            {
                cell = (HSSFCell)mrgRow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow, cellIndex, cellIndex));
            }
            if (TotalSection)
            {
                TotalQtyCol = cellIndex;
            }
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s28");
            cell.SetCellValue(column9heading);
            if (topTitle)
            {
                cell = (HSSFCell)mrgRow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow, cellIndex, cellIndex));
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
                    cell = (HSSFCell)hdrRow.CreateCell(cellIndex + (j++));
                    cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s29");
                    if(j==1)
                        cell.SetCellValue(ConfigurationManager.AppSettings["Attr1Label"] ?? "Size");
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow - 1, sheetRow - 1, cellIndex, cellIndex + cols.Count -1));

                hdrRow = mrgRow;
            }
            int preparkIdx = 0; HSSFRow[] preparkRows = new HSSFRow[PrepackAmt];
            if (topTitle) 
            {
                for(int i = 0;i < PrepackAmt;i++)
                {
                    sheetRow++;
                    preparkRows[i] = (HSSFRow)sheet.CreateRow(sheetRow);
                    preparkRows[i].HeightInPoints = RowHeight;
                    AddSecondaryRows(RowHeight, skusheet, pricesheet, atpsheet, sheetRow);

                    for (int j = 0; j < cellIndex; j++)
                    {
                        cell = (HSSFCell)preparkRows[i].CreateCell(j);
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                        if (j == 0)
                        {
                            cell.SetCellValue(string.Format("Prepack {0}  ", ((char)(65 + i)).ToString()));
                            cell.CellStyle.Alignment = HorizontalAlignment.RIGHT;
                        }
                        else if (j == cellIndex - 2)
                        {
                            cell.CellFormula = string.Format("SUM({0}{2}:{1}{2})", cellIndex.ToName(), (cellIndex + sizePrepack.Length -1).ToName(), sheetRow + 1);
                        }
                    }
                    sheet.AddMergedRegion(new CellRangeAddress(sheetRow, sheetRow, 0, cellIndex - 1 - 2));
                }
            }
            foreach (DictionaryEntry de in cols)
            {
                colcounter++;
                cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s29");
                if (!groupedTitle.HasValue || groupedTitle == true)
                {
                    cell.SetCellValue(string.IsNullOrEmpty(Convert.ToString(de.Value)) ? " " : StripHTML(Convert.ToString(de.Value)).Replace("\r", System.Environment.NewLine));
                    if (sizePrepack.Length > preparkIdx)
                    {
                        string[] preparkItem = sizePrepack[preparkIdx++].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < preparkRows.Length; i++)
                        {
                            cell = (HSSFCell)(preparkRows[i].CreateCell(cellIndex));
                            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s41");
                            int preparkQty = 0;
                            int.TryParse(preparkItem[i], out preparkQty);
                            cell.SetCellValue(preparkQty);
                        }
                    }
                }
                cellIndex++;
                colpositions.Add(StripHTML(Convert.ToString(de.Value)), colcounter);
            }
            return colpositions;
        }

        private int AddBreakRows(HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, int rownumber)
        {
            var sheetRow = sheet.LastRowNum + 1;
            HSSFRow newrow = null;
            for (int i = 0; i < rownumber; i++)
            {
                newrow = (HSSFRow)sheet.CreateRow(sheetRow);
                newrow.HeightInPoints = RowHeight;
                if (i == rownumber - 2) newrow.CreateCell(0).SetCellValue("pagebreak");
                //sheet.AddMergedRegion(new CellRangeAddress(rownumber, rownumber, 0, 6));
                AddSecondaryRows(12, skusheet, pricesheet, atpsheet, sheetRow);
                sheetRow++;
            }
            return sheetRow;
        }

        private int AddTextRow(HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, float rowHeight, string[][] texts)
        {
            var sheetRow = sheet.LastRowNum + 1;
            HSSFRow newrow = (HSSFRow)sheet.CreateRow(sheetRow);
            newrow.HeightInPoints = rowHeight;
            for (int i = 0; i < texts.Length; i++)
            {
                string[] arr = texts[i];
                HSSFCell cell = (HSSFCell)newrow.CreateCell(int.Parse(arr[2]));
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, arr[1]);
                cell.SetCellValue(arr[0]);
                sheet.AddMergedRegion(new CellRangeAddress(sheetRow, sheetRow, int.Parse(arr[2]), int.Parse(arr[3])));
                AddSecondaryRows(rowHeight, skusheet, pricesheet, atpsheet, sheetRow);
            }
            sheetRow++;
            return sheetRow;
        }

        public string StripHTML(string HTMLText)
        {
            var reg = new Regex("<[^>]+>", RegexOptions.IgnoreCase);
            return reg.Replace(HTMLText.Replace("<br>", "\r").Replace("<br/>", "\r").Replace("<p/>", "\r").Replace("</p>", System.Environment.NewLine).Replace("&nbsp;", " "), "");
        }

        private void CreateGridRows2(OrderedDictionary depts, HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, HSSFSheet upcsheet, int deptid, string deptname, ref ListDictionary colpositions)
        {
            bool firstrow = true, lastrow = false;
            int i = 0;
            OrderedDictionary rows = (OrderedDictionary)departments[deptid];

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
                    AddRow2(depts, rowattribs, skus, colpositions, sheet, skusheet, pricesheet, atpsheet, upcsheet, firstrow, lastrow);
                    firstrow = false;
                }
            }
        }

        private void AddRow2(OrderedDictionary DeptLevels, StringDictionary rowattribs, OrderedDictionary skus, ListDictionary colpositions, HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, HSSFSheet upcsheet, bool firstrow, bool lastrow)
        {
            string errMsg = "";
            var rowIndex = sheet.LastRowNum + 1;
            var upcRowIndex = 0;
            if (upcsheet != null) upcRowIndex = upcsheet.LastRowNum + 1;
            //Add the row to each sheet to keep them in synch
            HSSFRow newrow = (HSSFRow)sheet.CreateRow(rowIndex);
            HSSFRow newrowa = (HSSFRow)skusheet.CreateRow(rowIndex);
            HSSFRow newrowb = (HSSFRow)pricesheet.CreateRow(rowIndex);
            HSSFRow newrowc = (HSSFRow)atpsheet.CreateRow(rowIndex);
            //int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            //int offset = 8;
            int offset = 6;
            if (DeptLevels != null && DeptLevels.Count > 0 && multicolumndeptlabel)
                offset = DeptLevels.Count + 5;
            HSSFCell cell = null;
            string valueformula = "";
            string ttlformula = "";
            newrow.HeightInPoints = RowHeight;
            newrowa.HeightInPoints = RowHeight;
            newrowb.HeightInPoints = RowHeight;
            //Add spacer cells to the secondary sheets
            if (!multicolumndeptlabel)
                AddSecondaryCells(newrowa, newrowb, newrowc, maxcols + 7);
            else
                AddSecondaryCells(newrowa, newrowb, newrowc, maxcols + 6);
            //newrow.AutoFitHeight = false;
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
            int cellIndex = 0;
            if (!multicolumndeptlabel)
            {
                cell = (HSSFCell)newrow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s43");
                if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.THIN; }
                cell.SetCellValue("");
                cellIndex++;
            }
            else
            {
                foreach (DictionaryEntry de in DeptLevels)
                {
                    if (firstrow)
                    {
                        cell = (HSSFCell)newrow.CreateCell(cellIndex);
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s24");
                        if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
                        cell.SetCellValue(de.Value.ToString());
                        cellIndex++;
                    }
                    else
                    {
                        cell = (HSSFCell)newrow.CreateCell(cellIndex);
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s43");
                        if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
                        cell.SetCellValue("");
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
                            cell = (HSSFCell)newrow.CreateCell(cellIndex);
                            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s24");
                            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
                            cell.SetCellValue("");
                            cellIndex++;
                        }
                        else
                        {
                            cell = (HSSFCell)newrow.CreateCell(cellIndex);
                            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s43");
                            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
                            cell.SetCellValue("");
                            cellIndex++;
                        }
                    }
                }
            }
            //Mat.
            cell = (HSSFCell)newrow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s38");
            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
            cell.SetCellValue(rowattribs["Style"].ToString());
            cellIndex++;
            //newrow.Cells.Add(rowattribs["Style"].ToString(), DataType.String, "s38");
            //Mat Desc
            cell = (HSSFCell)newrow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s38");
            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
            if (rowattribs["ProductName"] == null)
            {
                cell.SetCellValue(rowattribs["Style"].ToString());
            }
            else
            {
                cell.SetCellValue(rowattribs["ProductName"].ToString());
            }
            cellIndex++;
            //cell = newrow.Cells.Add(rowattribs["ProductName"].ToString(), DataType.String, "s38");
            //Dim1
            cell = (HSSFCell)newrow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s38");
            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
            cell.SetCellValue(rowattribs["GridAttributeValues"].ToString());
            cellIndex++;
            //newrow.Cells.Add(rowattribs["GridAttributeValues"].ToString(), DataType.String, "s38");
            //WHS $
            cell = (HSSFCell)newrow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s38");
            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
            cell.CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            var price = this.IncludePricing ? double.Parse(rowattribs["RowPriceWholesale"].ToString()) : double.Parse(ExcludePriceValue); // 0;
            cell.SetCellValue(price);
            cellIndex++;

            //TTL
            var TTLIndex = cellIndex;
            HSSFCell ttlcell = (HSSFCell)newrow.CreateCell(cellIndex);
            ttlcell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s41");
            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
            cellIndex++;
            //ttlcell.Data.Type = DataType.Number;

            //TTL Value
            HSSFCell valuecell = (HSSFCell)newrow.CreateCell(cellIndex);
            valuecell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s40");
            if (lastrow) { valuecell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
            cellIndex++;
            //valuecell.Data.Type = DataType.Number;

            //Lay out the row with empty, greyed-out cells
            for (int i = 0; i < colpositions.Count; i++)
            {
                cell = (HSSFCell)newrow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s81");
                if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
                cellIndex++;
            }
            AddSecondaryCells(newrowa, newrowb, newrowc, offset + colpositions.Count);

            //Now enable valid cells
            try
            {
                foreach (DictionaryEntry de in skus)
                {
                    //SKU must be enabled
                    if (((StringDictionary)(de.Value))["Enabled"].ToString() == "1")
                    {
                        //This is the column header value
                        string dav = ((StringDictionary)(de.Value))["AttributeValue1"].ToString();
                        string cellposition = "";
                        //If the row contains a cell with the column header value, enable it
                        dav = StripHTML(Convert.ToString(dav));
                        if (colpositions.Contains(dav))
                        {
                            int colposition = Convert.ToInt32(colpositions[dav].ToString());
                            if (multicolumndeptlabel)
                                cell = (HSSFCell)newrow.GetCell(Convert.ToInt32(colpositions[dav].ToString()) + maxcols + 5);
                            else
                                cell = (HSSFCell)newrow.GetCell(Convert.ToInt32(colpositions[dav].ToString()) + 6);
                            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s62");
                            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
                            if (multicolumndeptlabel)
                                cellposition = ConvertToLetter(Convert.ToInt32(colpositions[dav].ToString()) + maxcols + 6) + Convert.ToString(sheet.LastRowNum + 1);
                            else
                                cellposition = ConvertToLetter(Convert.ToInt32(colpositions[dav].ToString()) + 7) + Convert.ToString(sheet.LastRowNum + 1);
                            unlocked.Add(cellposition, cellposition);
                            if (((StringDictionary)(de.Value)).ContainsKey("OrderMultiple"))
                                multiples.Add(cellposition, ((StringDictionary)(de.Value))["OrderMultiple"].ToString());
                            //Plant the QuickWebSKU on its sheet
                            cell = SetCellValue(newrowa, Convert.ToInt32(colpositions[dav].ToString()) + offset, ((StringDictionary)(de.Value))["QuickWebSKUValue"].ToString(), false);
                            //Plant the ATPDate on its sheet
                            if (OrderShipDate)
                                cell = SetCellValue(newrowc, Convert.ToInt32(colpositions[dav].ToString()) + offset, ((StringDictionary)(de.Value))["ATPDate"].ToString(), false);
                            else
                                cell = SaveATPValueDataToSheet(newrowc, Convert.ToInt32(colpositions[dav].ToString()) + offset, ((StringDictionary)(de.Value))["QuickWebSKUValue"].ToString());
                            //Plant the wholesale price on its sheet
                            cell = SetCellValue(newrowb, Convert.ToInt32(colpositions[dav].ToString()) + offset, ((StringDictionary)(de.Value))["PriceWholesale"].ToString(), true);
                            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s62");
                            if (lastrow) { cell.CellStyle.BorderBottom = allThinBorder ? CellBorderType.THIN : CellBorderType.MEDIUM; }
                            //Build the ttl cell formula
                            ttlformula += GetExcelColumnName(TTLIndex + 1 + colposition + 1) + (rowIndex + 1).ToString() + "+";
                            //"RC[" + Convert.ToString(colposition + 1) + "]+";
                            //Build the value cell formula
                            valueformula += "(" + GetExcelColumnName(TTLIndex + 1 + colposition + 1) + (rowIndex + 1).ToString() +
                                "*WholesalePrice!" + GetExcelColumnName(TTLIndex + 1 + colposition + 1) + (rowIndex + 1).ToString() + ")" + "+";
                            //"(RC[" + Convert.ToString(colposition) + "]*WholesalePrice!RC[" + Convert.ToString(colposition) + "])+";

                            if (upcsheet != null)
                            {
                                HSSFRow newrowd = (HSSFRow)upcsheet.CreateRow(upcRowIndex);
                                var skuattrvalue2 = ((StringDictionary)(de.Value)).ContainsKey("AttributeValue3") ? ((StringDictionary)(de.Value))["AttributeValue3"].ToString() : string.Empty;
                                cell = SetCellValue(newrowd, 0, rowattribs["ProductName"].ToString(), false);
                                cell = SetCellValue(newrowd, 1, rowattribs["Style"].ToString(), false);
                                cell = SetCellValue(newrowd, 2, skuattrvalue2, false);
                                cell = SetCellValue(newrowd, 3, dav, false);
                                cell = SetCellValue(newrowd, 4, ((StringDictionary)(de.Value))["UPC"].ToString(), false);
                                cell = SetCellValue(newrowd, 5, ((StringDictionary)(de.Value))["PriceWholesale"].ToString(), false);
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
                ttlcell.CellFormula = ttlformula.Substring(0, ttlformula.Length - 1);
            if (valueformula.Length > 0)
                valuecell.CellFormula = valueformula.Substring(0, valueformula.Length - 1);
        }

        private HSSFCell SetCellValue(HSSFRow row, int cellIndex, string value, bool bdouble)
        {
            var cell = row.GetCell(cellIndex) == null ? (HSSFCell)row.CreateCell(cellIndex) : (HSSFCell)row.GetCell(cellIndex);
            if(bdouble)
                cell.SetCellValue(double.Parse(value));
            else
                cell.SetCellValue(value);
            return cell;
        }

        private void AddSecondaryCells(HSSFRow row1, HSSFRow row2, HSSFRow row3, int cellsToAdd)
        {
            var cellIndex1 = row1.LastCellNum + 1;
            for (int i = 0; i < cellsToAdd; i++)
            {
                row1.CreateCell(cellIndex1);
                cellIndex1++;
            }
            var cellIndex2 = row2.LastCellNum + 1;
            for (int i = 0; i < cellsToAdd; i++)
            {
                row2.CreateCell(cellIndex2);
                cellIndex2++;
            }
            if (OrderShipDate)
            {
                var cellIndex3 = row3.LastCellNum + 1;
                for (int i = 0; i < cellsToAdd; i++)
                {
                    row3.CreateCell(cellIndex3);
                    cellIndex3++;
                }
            }
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        } 

        private OrderedDictionary GetDeptHierarchy(XmlReader reader, ref int deptid, ref string deptname, ref Dictionary<string, string[]> rootDept, string catalog)
        {
            OrderedDictionary depts = new OrderedDictionary();
            reader.Read();
            reader.MoveToNextAttribute();
            deptid = Convert.ToInt32(reader.Value);
            reader.MoveToNextAttribute();
            var deptLevel = reader.Value;
            reader.MoveToNextAttribute();
            deptname = reader.Value;

            string lbl = string.Empty, attr1Values = string.Empty, PrepackValues = string.Empty;
            if (deptLevel == "0")
                rootDept.Add(deptid.ToString(), new string[] { deptname, lbl, attr1Values, PrepackValues });
            else
            {

                string parentdeptname = "";
                while (reader.Read())
                {
                    if (reader.Name == "ParentLevel")
                    {
                        reader.MoveToNextAttribute();
                        parentdeptname = reader.Value;
                        reader.MoveToNextAttribute();
                        string parentdeptId = reader.Value;
                        depts.Add(parentdeptId, parentdeptname);

                        try
                        {
                            reader.MoveToNextAttribute();
                            if (reader.Name == "level" && reader.Value == "0")
                            {
                                rootDept = new Dictionary<string, string[]>();
                                reader.MoveToNextAttribute();
                                lbl = string.Empty;
                                attr1Values = string.Empty; 
                                PrepackValues = string.Empty;
                                if (reader.Name == "Attribute1Label")
                                {
                                    lbl = reader.Value;
                                }
                                reader.MoveToNextAttribute();
                                if (reader.Name == "Attribute1Values")
                                {
                                    attr1Values = reader.Value;
                                }
                                reader.MoveToNextAttribute();
                                if (reader.Name == "PrepackValues")
                                {
                                    PrepackValues = reader.Value;
                                    if (PrepackValues.Split(new char[] { ',', '|' }, StringSplitOptions.RemoveEmptyEntries).Length == 0)
                                        PrepackValues = string.Empty;
                                }

                                rootDept.Add(parentdeptId, new string[] { parentdeptname, lbl, attr1Values, PrepackValues });
                            }
                        }
                        catch { }
                    }
                }
            }
            if (rootDept.Count <= 0)
            {
                rootDept.Add("0", new string[] { catalog, string.Empty, string.Empty, string.Empty });
                depts.Add("0", catalog);
            }
            depts.Add(deptid, deptname);
            return depts;
        }

        private void AddBlankHeader(ref OrderedDictionary deptcols, int deptid)
        {
            //See if there are any products in this dept with "*" (blank) AttributeValue1
            //If so, add a blank header
            OrderedDictionary rows = (OrderedDictionary)departments[deptid];
            if (rows != null && rows.Count > 0)
            {
                foreach (DictionaryEntry de in rows)
                {
                    OrderedDictionary skus = (OrderedDictionary)de.Value;
                    foreach (DictionaryEntry de1 in skus)
                    {
                        if (de1.Key.ToString() == " " && !deptcols.Contains(" "))
                        {
                            deptcols.Add(" ", " ");
                            break;
                        }
                    }
                }
            }
        }

        public XmlReader GetDepartments2(string catalog, string pricecode)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand("GetOfflineDepartments", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 240;

            cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
            cmd.Parameters["@catcd"].Value = catalog;

            cmd.Parameters.Add(new SqlParameter("@pricecd", SqlDbType.VarChar, 80));
            cmd.Parameters["@pricecd"].Value = pricecode;

            if (!string.IsNullOrEmpty(OfflineDeptID))
            {
                cmd.Parameters.Add(new SqlParameter("@DeptID", SqlDbType.VarChar, 80));
                cmd.Parameters["@DeptID"].Value = OfflineDeptID;
            }

            //Add params 
            XmlReader reader = null;
            try
            {
                conn.Open();
                reader = cmd.ExecuteXmlReader();
            }

            catch
            {
                throw;
            }
            return reader;
        }

        public SqlDataReader GetFiles(string batchNo)
        {
            //Get the code rows describing which spreadsheets to create
            string errMsg = "";
            SqlDataReader reader = null;
            try
            {
                reader = GetOfflineRows(batchNo);
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                if (errMsg.IndexOf("Timeout expired") >= 0)
                {
                    for (int i = 0; i < 10; i++)
                    {
                        try
                        {
                            reader = GetOfflineRows(batchNo);
                            break;
                        }
                        catch (Exception ex1)
                        {
                            WriteToLog("GetFiles attempt:" + i.ToString());
                        }
                    }
                }
            }
            return reader;
        }

        private SqlDataReader GetOfflineRows(string batchNo)
        {
            //Get the code rows describing which spreadsheets to create
            string errMsg = "";
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlDataReader reader = null;
            try
            {
                SqlConnection conn = new SqlConnection(connString);
                SqlCommand cmd = new SqlCommand("p_GetOfflineOrderFormTemplateByBatchNo", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@BatchNo", batchNo));
                cmd.CommandTimeout = 180;

                conn.Open();
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                WriteToLog("p_GetOfflineOrderFormTemplateByBatchNo", ex, string.Empty);
                //if(ex.GetType() == SqlConnectioexcep Timeout expired
                throw ex;
            }
            return reader;
        }

        private string ConvertToLetter(int iCol)
        {
            int iAlpha = 0;
            int iRemainder = 0;
            string temp = "";
            iAlpha = Convert.ToInt32(iCol / 27);
            iRemainder = iCol - (iAlpha * 26);
            if (iAlpha > 0)
                temp = Convert.ToString(Convert.ToChar(iAlpha + 64));
            if (iRemainder > 0)
                temp += Convert.ToString(Convert.ToChar(iRemainder + 64));
            if (temp == "A[")
                temp = "BA";
            if (temp == "B[")
                temp = "CA";
            if (temp == @"B\")
                temp = "CB";
            if (temp == "B]")
                temp = "CC";
            if (temp == "C[")
                temp = "DA";
            if (temp == @"C\")
                temp = "DB";
            if (temp == "C]")
                temp = "DC";
            return temp;
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

        public void LogOfflineOrderFormTemplate(string batchNo, string codeId)
        {
            try
            {
                //Get the code rows describing which spreadsheets to create
                string errMsg = "";
                string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
                SqlConnection conn = new SqlConnection(connString);
                SqlCommand cmd = new SqlCommand("p_LogOfflineOrderFormTemplate", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                //Add params 
                cmd.Parameters.Add(new SqlParameter("@BatchNo", batchNo));
                cmd.Parameters.Add(new SqlParameter("@CodeId", codeId));

                conn.Open();

                cmd.ExecuteNonQuery();

                conn.Close();
            }
            catch (Exception ex)
            {
                WriteToLog("p_LogOfflineOrderFormTemplate", ex, string.Empty);
                //throw ex;
            }
        }

        public DataSet GetAllShipMethod()
        {
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
            return shipmethodds;
        }

        private HSSFCell SaveATPValueDataToSheet(HSSFRow row, int cellIndex, string SKU)
        {
            HSSFCell cell = null;
            if(ATPValueData == null)
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
                cell = SetCellValue(row, cellIndex, cellValue, false);
            }
            return cell;
        }


        #region "Modify for performance"
        private static string TempDeptXMLPath;
        private static string TemplateXMLFilePath;
        private static List<DeptInfo> DeptIDList;
        private static List<SKUAttrValue1Info> AttrRegList;

        private class DeptInfo
        {
            public int DeptId { get; set; }
            public int Sort { get; set; }
        }

        private class SKUAttrValue1Info
        {
            public string AttrValue1 { get; set; }
            public int Sort { get; set; }
        }

        //string filename = soldto + "_" + catalog + "_" + "Template.xml";
        //string filepath = TemplateXMLFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filename);

        private bool PrepareAllProductData(string soldto, string catalog, string pricecode, string savedirectory)
        {
            TempDeptXMLPath = "";
            TemplateXMLFilePath = "";
            if(!GetDepartmentList(catalog, pricecode)) return false;

            string savexmlpath = ConfigurationManager.AppSettings["SaveXMLPath"];
            CreateTemplateXMLFile(soldto, catalog, savexmlpath, savedirectory);//******************

            return PreparePorductDataFromXML(soldto, catalog, savedirectory);
        }
        
        #region "get all departments list in order to loop them to generate product xml"
        private bool GetDepartmentList(string catalog, string pricecode)
        {
            CreateDepTemplateXML(catalog, pricecode);

            return LoadDepartmentList();
        }

        private void CreateDepTemplateXML(string catalog, string pricecode)
        {
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand("GetOfflineDepartments", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 240;

            cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
            cmd.Parameters["@catcd"].Value = catalog;

            cmd.Parameters.Add(new SqlParameter("@pricecd", SqlDbType.VarChar, 80));
            cmd.Parameters["@pricecd"].Value = pricecode;

            if (!string.IsNullOrEmpty(OfflineDeptID))
            {
                cmd.Parameters.Add(new SqlParameter("@DeptID", SqlDbType.VarChar, 80));
                cmd.Parameters["@DeptID"].Value = OfflineDeptID;
            }

            try
            {
                string filename = catalog.Replace("\\", "-").Replace("/", "-") + "_" + pricecode.Replace("\\", "-").Replace("/", "-") + "_" + "Template_Dept.xml";
                string filepath = TempDeptXMLPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Program.CleanFileName(filename));
                if (File.Exists(filepath)) File.Delete(filepath);
                conn.Open();
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    StreamWriter sw = new StreamWriter(filepath, true);
                    while (reader.Read())
                    {
                        string xmlvalue = reader.GetValue(0).ToString();
                        sw.Write(xmlvalue);
                    }
                    sw.Close();
                    reader.Close();
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }

        private bool LoadDepartmentList()
        {
            if (!string.IsNullOrEmpty(TempDeptXMLPath) && File.Exists(TempDeptXMLPath))
            {
                FileInfo file = new FileInfo(TempDeptXMLPath);
                if (file.Length <= 0) return false;

                DeptIDList = new List<DeptInfo>();
                XElement xroot = XElement.Load(TempDeptXMLPath);
                var cdeptlist = from r in xroot.Descendants("ProductLevel") select r;
                foreach (XElement cdept in cdeptlist)
                {
                    int cid = Convert.ToInt32(cdept.Attribute("departmentID").Value);
                    var exitdept = from d in DeptIDList where d.DeptId == cid select d;
                    if (exitdept.Count() <= 0)
                    {
                        DeptIDList.Add(new DeptInfo { DeptId = cid, Sort = Convert.ToInt32(cdept.Attribute("SortOrder").Value) });
                    }
                }
            }
            return true;
        }
        #endregion

        #region "generate products xml"
        private void CreateTemplateXMLFile(string soldto, string catalog, string xmlsavepath, string savedirectory)
        {
            if (DeptIDList == null || DeptIDList.Count <= 0) return;
            TemplateXMLFilePath = Path.Combine(xmlsavepath, Program.CleanFileName(soldto.Replace("\\", "-").Replace("/", "-") + "_" + catalog.Replace("\\", "-").Replace("/", "-") +
                (string.IsNullOrEmpty(OfflineDeptID) ? string.Empty : "_" + OfflineDeptID) + "_" + "Template.xml"));
            if (File.Exists(TemplateXMLFilePath))
                File.Delete(TemplateXMLFilePath);//temp to do ****************** 
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connString);
            conn.Open();
            try
            {
                string filenamei = soldto + "_" + "Template.xml";
                string filepathi = Path.Combine(xmlsavepath, Program.CleanFileName(filenamei));
                //----------------------
                string servername = ConfigurationManager.AppSettings["ServerName"];
                string dbname = ConfigurationManager.AppSettings["DBName"];
                CreateTemplateXMLPart(soldto, catalog, conn, filepathi, servername, dbname, savedirectory);
                //--------------------
                if (File.Exists(filepathi))
                {
                    FileInfo file = new FileInfo(filepathi);
                    if (file.Length > 0)
                    {
                        File.Copy(filepathi, TemplateXMLFilePath);
                    }
                    File.Delete(filepathi);
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                if (conn.State == ConnectionState.Open) conn.Close();
                throw;
            }
        }

        private void CreateTemplateXMLPart(string soldto, string catalog, SqlConnection conn, string partfilepath, string servername, string dbname, string savedirectory)
        {
            SqlCommand cmd = new SqlCommand("p_Offline_GridViewRectangular", conn);
            cmd.CommandTimeout = 0;
            cmd.CommandType = CommandType.StoredProcedure;
            //cmd.CommandTimeout = 240;

            cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
            cmd.Parameters["@catcd"].Value = catalog;

            cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
            cmd.Parameters["@SoldTo"].Value = soldto;

            cmd.Parameters.Add(new SqlParameter("@ServerName", SqlDbType.VarChar, 100));
            cmd.Parameters["@ServerName"].Value = servername;

            cmd.Parameters.Add(new SqlParameter("@DBName", SqlDbType.VarChar, 100));
            cmd.Parameters["@DBName"].Value = dbname;

            cmd.Parameters.Add(new SqlParameter("@SavePath", SqlDbType.VarChar, 500));
            cmd.Parameters["@SavePath"].Value = partfilepath;

            if (!string.IsNullOrEmpty(PriceType))
            {
                cmd.Parameters.Add(new SqlParameter("@PriceType", SqlDbType.VarChar, 500));
                cmd.Parameters["@PriceType"].Value = PriceType;
            }

            if (!string.IsNullOrEmpty(OfflineDeptID))
            {
                cmd.Parameters.Add(new SqlParameter("@DeptID", SqlDbType.VarChar, 80));
                cmd.Parameters["@DeptID"].Value = OfflineDeptID;
            }

            //Add params 
            try
            {
                if (File.Exists(partfilepath)) File.Delete(partfilepath);
                DataSet outds = new DataSet();
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                adapt.Fill(outds);
                if (outds.Tables.Count > 1)
                {
                    XElement xroot = XElement.Load(new StringReader(outds.Tables[1].Rows[0][0].ToString()));
                    var value1collection = from r in xroot.Descendants("AttrV1") select r;
                    AttrRegList = new List<SKUAttrValue1Info>();
                    value1collection.ToList().ForEach(v => 
                        {
                            if(v.Attribute("AttrValue1") != null)
                            AttrRegList.Add(new SKUAttrValue1Info() { AttrValue1 = v.Attribute("AttrValue1").Value, Sort = Convert.ToInt32(v.Attribute("SortOrder").Value)});
                            else
                                AttrRegList.Add(new SKUAttrValue1Info(){AttrValue1=" ", Sort=999});
                        });

                    ATPValueData= null;
                    if (outds.Tables.Count > 2)
                        ATPValueData = outds.Tables[2];
                }
            }
            catch (Exception ex)
            {
                //throw;
                WriteToLog("p_Offline_GridViewRectangular|CreateTemplateXMLFile|" + soldto + "|" + catalog, ex, savedirectory);
            }
        }

        private bool PreparePorductDataFromXML(string soldto, string catalog, string savedirectory)
        {
            try
            {
                if (!string.IsNullOrEmpty(TemplateXMLFilePath) && File.Exists(TemplateXMLFilePath))
                {
                    FileInfo file = new FileInfo(TemplateXMLFilePath);
                    if (file.Length <= 0) return false;

                    XDocument xdoc = XDocument.Load(TemplateXMLFilePath);
                    XElement xroot = xdoc.Root;
                    List<XElement> deptlist = (from r in xroot.Descendants("DeptInfo") select r).ToList();
                    foreach (XElement dept in deptlist)
                    {
                        OrderedDictionary rows = new OrderedDictionary();
                        List<XElement> prdlist = (from d in dept.Descendants("GridRow") select d).ToList();
                        foreach (XElement product in prdlist)
                        {
                            OrderedDictionary skus = new OrderedDictionary();
                            StringDictionary rowattribs = new StringDictionary();
                            foreach (XAttribute patt in product.Attributes())
                            {
                                rowattribs.Add(patt.Name.ToString(), patt.Value);
                            }
                            List<XElement> skulist = (from p in product.Descendants("SKU") select p).ToList();
                            foreach (XElement sku in skulist)
                            {//GridViewRectangular
                                StringDictionary skuattribs = new StringDictionary();
                                String dav = sku.Attribute("AttributeValue1") == null ? " " : sku.Attribute("AttributeValue1").Value;
                                try
                                {
                                    //Store the SKU using its column header as the key
                                    if (sku.Attribute("Enabled") != null && sku.Attribute("Enabled").Value == "1")//skuattribs["Enabled"].ToString() == "1")
                                    {
                                        foreach (XAttribute satt in sku.Attributes())
                                        {
                                            if (!skuattribs.ContainsKey(satt.Name.ToString()))
                                            {
                                                if (!IncludePricing && (satt.Name.ToString().ToLower() == "pricewholesale" || satt.Name.ToString().ToLower() == "pricenet"))
                                                    satt.Value = ExcludePriceValue; // "0";
                                                skuattribs.Add(satt.Name.ToString(), satt.Value);
                                            }
                                        }
                                        if (dav.Equals(" "))
                                            if (!skuattribs.ContainsKey("AttributeValue1"))
                                                skuattribs.Add("AttributeValue1", dav);
                                        if(!skus.Contains(dav))
                                            skus.Add(dav, skuattribs);
                                    }
                                }
                                catch (Exception e)
                                {
                                    string msg = "SoldTo:" + soldto + ", CatCd:" + catalog + ", Product:" + rowattribs["ProductName"].ToString() + ", UPC:" + skuattribs["UPC"].ToString();
                                    WriteToLog(msg, e, savedirectory);
                                }
                            }
                            if(!rows.Contains(rowattribs))
                                rows.Add(rowattribs, skus);
                        }
                        if (rows.Count > 0) departments.Add(Convert.ToInt32(dept.Attribute("DepartmentID").Value), rows);
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                WriteToLog("PrepareData", ex, savedirectory);
            }
            return false;
        }
        #endregion

        public OrderedDictionary GetDeptColsFromXML(XElement xroot,int deptid)
        {
            List<SKUAttrValue1Info> skuattlist = new List<SKUAttrValue1Info>();
            OrderedDictionary cols = new OrderedDictionary();
            if (AttrRegList != null && AttrRegList.Count > 0)
            {
                var sortresult = from a in AttrRegList orderby a.Sort select a;
                sortresult.ToList().ForEach(s => { if (!cols.Contains(s.AttrValue1)) cols.Add(s.AttrValue1, s.AttrValue1); });
            }
            else
            {
                XElement deptinfo = (from r in xroot.Descendants("DeptInfo") where r.Attribute("DepartmentID").Value == deptid.ToString() select r).FirstOrDefault();
                if (deptinfo != null)
                {
                    var xpdtlist = from d in deptinfo.Descendants("GridRow") select d;
                    foreach (XElement xpdt in xpdtlist)
                    {
                        var skulist = from p in xpdt.Descendants("SKU") select p;
                        foreach (XElement sku in skulist)
                        {
                            if (sku.Attribute("AttributeValue1") != null)
                            {
                                var existvalue = from a in skuattlist where a.AttrValue1 == sku.Attribute("AttributeValue1").Value select a;
                                if (existvalue.Count() <= 0)
                                    skuattlist.Add(new SKUAttrValue1Info { AttrValue1 = sku.Attribute("AttributeValue1").Value, Sort = Convert.ToInt32(sku.Attribute("SortSKU") == null ? "0" : sku.Attribute("SortSKU").Value) });
                                else
                                {
                                    var sort = Convert.ToInt32(sku.Attribute("SortSKU") == null ? "0" : sku.Attribute("SortSKU").Value);
                                    if (sort < existvalue.First().Sort)
                                        existvalue.First().Sort = sort;
                                }
                            }
                        }
                    }
                    var sortresult = from a in skuattlist orderby a.Sort select a;
                    sortresult.ToList().ForEach(s => { if (!cols.Contains(s.AttrValue1)) cols.Add(s.AttrValue1, s.AttrValue1); });
                }
            }
            return cols;
        }

        /// <summary>
        /// Get B2B setting value.
        /// </summary>
        /// <param name="category"></param>
        /// <param name="setting"></param>
        /// <returns></returns>
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
                catch(Exception ex)
                {
                    WriteToLog("p_B2B_Setting", ex, string.Empty);
                }
            }
            return ret;
        }
        /*public void LogOfflineOrd33erFormTemplate(string batchNo, string codeId)
        {
            try
            {
                //Get the code rows describing which spreadsheets to create
                string errMsg = "";
                string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
                SqlConnection conn = new SqlConnection(connString);
                SqlCommand cmd = new SqlCommand("p_LogOfflineOrderFormTemplate", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                //Add params 
                cmd.Parameters.Add(new SqlParameter("@BatchNo", batchNo));
                cmd.Parameters.Add(new SqlParameter("@CodeId", codeId));

                conn.Open();

                cmd.ExecuteNonQuery();

                conn.Close();
            }
            catch (Exception ex)
            {
                WriteToLog("p_LogOfflineOrderFormTemplate", ex, string.Empty);
                throw ex;
            }
        }*/



        #region "old code"
        private int PrepareData(string soldto, string catalog, string pricecode, string savedirectory)
        {
            //Returns the max number of data columns (used to set column widths)
            int maxcols = 0;
            //Number of levels deep for the catalog hierarchy
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());

            XmlReader reader = GetDepartments2(catalog, pricecode);
            reader.Read();
            reader.MoveToNextAttribute();
            maxcols = Convert.ToInt32(reader.Value);
            try
            {
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.EndElement && reader.Name == "ProductLevel")
                    {
                        //Product rows for this department
                        OrderedDictionary rows = new OrderedDictionary();
                        reader.MoveToNextAttribute();
                        int deptid = Convert.ToInt32(reader.Value);
                        AddProducts(soldto, catalog, pricecode, savedirectory, rows, deptid);

                        //Store the department by ID with its product rows, 
                        //only if there are any
                        if (rows.Count > 0)
                            departments.Add(deptid, rows);
                    }
                }
                if (departments.Count > 0)
                    return maxcols;
                else
                    return -1;
            }
            catch (Exception ex)
            {
                WriteToLog("PrepareData", ex, savedirectory);
                throw ex;
            }
            finally
            {
                reader.Close();
            }
        }

        private void AddProducts(string soldto, string catalog, string pricecode, string savedirectory, OrderedDictionary rows, int deptid)
        {
            XmlReader prodreader = GetProducts(deptid, catalog, soldto);
            if (prodreader != null)
            {
                while (prodreader.Read())
                {
                    if (prodreader.Name == "GridRow")
                    {
                        //Each row element will have several row-related attributes
                        //and one or more sku elements
                        OrderedDictionary skus = new OrderedDictionary();
                        StringDictionary rowattribs = new StringDictionary();
                        //Store the row attributes
                        for (int i = 0; i < prodreader.AttributeCount; i++)
                        {
                            prodreader.MoveToNextAttribute();
                            rowattribs.Add(prodreader.Name, prodreader.Value);
                        }
                        prodreader.Read();
                        //Store the SKU information
                        while (prodreader.Name == "SKU")
                        {
                            StringDictionary skuattribs = new StringDictionary();
                            String dav = "";
                            for (int j = 0; j < prodreader.AttributeCount; j++)
                            {
                                prodreader.MoveToNextAttribute();
                                skuattribs.Add(prodreader.Name, prodreader.Value);
                                if (prodreader.Name == "AttributeValue1")
                                    dav = prodreader.Value;
                            }
                            try
                            {
                                //Store the SKU using its column header as the key
                                if (skuattribs["Enabled"].ToString() == "1")
                                {
                                    if (dav.Length > 0)
                                        skus.Add(dav, skuattribs);
                                    else
                                    {
                                        skuattribs.Add("AttributeValue1", " ");
                                        skus.Add(" ", skuattribs);
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                string msg = "SoldTo:" + soldto + ", CatCd:" + catalog + ", PriceCode:" + pricecode + ", Product:" + rowattribs["ProductName"].ToString() + ", UPC:" + skuattribs["UPC"].ToString();
                                WriteToLog(msg, e, savedirectory);
                            }
                            //if(!deptcols.Contains(dav))
                            //    deptcols.Add(dav, dav);
                            prodreader.Read();
                        }
                        if(!rows.Contains(rowattribs))
                            rows.Add(rowattribs, skus);
                    }
                }
            }
        }

        public static OrderedDictionary GetDeptCols(int deptid, string soldto, string catalog)
        {
            //Get the attributevalue1 columns (size, color etc.) for the specified department
            OrderedDictionary cols = new OrderedDictionary();
            string errMsg = "";
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand("p_departmentattributevalue1", conn);
            cmd.CommandTimeout = 0;
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.Add(new SqlParameter("@Catalog", SqlDbType.VarChar, 80));
            cmd.Parameters["@Catalog"].Value = catalog;

            cmd.Parameters.Add(new SqlParameter("@DeptId", SqlDbType.Int, 4));
            cmd.Parameters["@DeptId"].Value = deptid;

            cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
            cmd.Parameters["@SoldTo"].Value = soldto;

            //Add params 
            SqlDataReader reader = null;
            try
            {
                conn.Open();
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    if (!reader.IsDBNull(0) && !cols.Contains(reader[0].ToString()))// && reader[0].ToString().Trim().Length > 0)
                    {
                        cols.Add(reader[0].ToString(), reader[0].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                WriteToLog("Failed to Call p_departmentattributevalue1:" + deptid.ToString() + "|" + catalog + "|" + soldto + "|" + errMsg);
            }

            return cols;
        }

        public static XmlReader GetProducts(int deptid, string catalog, string soldto)
        {
            string errMsg = "";
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand("p_GridViewRectangular", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 240;

            cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
            cmd.Parameters["@catcd"].Value = catalog;

            cmd.Parameters.Add(new SqlParameter("@deptid", SqlDbType.Int, 4));
            cmd.Parameters["@deptid"].Value = deptid;

            cmd.Parameters.Add(new SqlParameter("@SoldTo", SqlDbType.VarChar, 80));
            cmd.Parameters["@SoldTo"].Value = soldto;

            //Add params 
            XmlReader reader = null;
            try
            {
                conn.Open();
                reader = cmd.ExecuteXmlReader();
            }

            catch (Exception ex)
            {
                //throw;
                errMsg = ex.Message;
            }
            return reader;
        }
        #endregion
        #endregion

        private static void WriteToLog(string msg)
        {
            //string templatefile = ConfigurationManager.AppSettings["templatefile"].ToString();
            var logFile = "OfflineLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            string savefolder = ConfigurationManager.AppSettings["savefolder"].ToString();
            StreamWriter sw = new StreamWriter(savefolder + logFile, true);
            sw.WriteLine(msg);
            sw.Close();
        }



        #region GENERATE IMAGE OFFLINE ORDER FORM

        private class SKUImageInfo
        {
            public string SKU { get; set; }
            public int RowIndex { get; set; }
            public string ImageName { get; set; }
        }

        public void GenerateImageOfflineOrderForm(string noneImageFilePath, string imageFilePath)
        {
            //copy file
            if (!string.IsNullOrEmpty(noneImageFilePath))
                File.Copy(noneImageFilePath, imageFilePath, true);
            if (!File.Exists(imageFilePath))
            {
                WriteToLog("can not file offline file " + imageFilePath);
                return;
            }
            FileStream sourceFileStream = new FileStream(imageFilePath, FileMode.Open, FileAccess.Read);
            NPOI.POIFS.FileSystem.POIFSFileSystem sourceFile = new NPOI.POIFS.FileSystem.POIFSFileSystem(sourceFileStream);
            //FileStream sourceFile = new FileStream(imageFilePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook book = new HSSFWorkbook(sourceFile);
            //get image list with first sku
            var skus = GetFirstSKUList(book);
            //generate images in offline
            var dt = GetSKUImageList(skus);
            if (dt != null)
            {
                if (ConfigurationManager.AppSettings["column1Imagewidth"] != null)
                    book.GetSheet("Order Template").SetColumnWidth(0, Convert.ToInt32(ConfigurationManager.AppSettings["column1Imagewidth"].ToString()) * 50);

                SaveImagesToOffline(dt, book, imageFilePath);
            }
        }

        private List<SKUImageInfo> GetFirstSKUList(HSSFWorkbook book)
        {
            var column4heading = System.Configuration.ConfigurationManager.AppSettings["column4heading"];
            List<SKUImageInfo> skus = new List<SKUImageInfo>();
            var index = book.GetSheetIndex("WebSKU");
            var sheet = (HSSFSheet)book.GetSheetAt(index);

            var tindex = book.GetSheetIndex("Order Template");
            var tsheet = (HSSFSheet)book.GetSheetAt(tindex);

            var lastRow = sheet.LastRowNum;
            var maxColNumber = 20;
            var firstColumn = 7; //H
            var firstRow = int.Parse(System.Configuration.ConfigurationManager.AppSettings["SheetPrdFirstRow"] ?? "6");
            for (int i = firstRow; i <= lastRow; i++)
            {
                var skuvalue = string.Empty;
                for (int j = firstColumn; j < maxColNumber + firstColumn; j++)
                {
                    skuvalue = GetCellValue(sheet, i, j);
                    if (!string.IsNullOrEmpty(skuvalue))
                    {
                        break;
                    }
                }
                if (string.IsNullOrEmpty(skuvalue))
                    skuvalue = GetCellValue(tsheet, i, 1);
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

        private string GetCellValue(HSSFSheet sheet, int rowIndex, int cellIndex)
        {
            string value = string.Empty;
            if (sheet.GetRow(rowIndex) != null)
            {
                var row = (HSSFRow)sheet.GetRow(rowIndex);
                if (row.GetCell(cellIndex) == null)
                {
                    value = string.Empty;
                }
                else
                {
                    try
                    {
                        value = row.GetCell(cellIndex).StringCellValue;
                    }
                    catch
                    {
                        value = row.GetCell(cellIndex).NumericCellValue.ToString(new System.Globalization.CultureInfo(1033));
                    }
                }
            }
            return value.Trim();
        }

        private DataTable GetSKUImageList(List<SKUImageInfo> skus)
        {
            var root = new XElement("SKUs");
            skus.ForEach(s =>
            {
                var sku = new XElement("SKU");
                sku.Add(new XAttribute("SKU", s.SKU));
                sku.Add(new XAttribute("ROW", s.RowIndex.ToString()));
                root.Add(sku);
            });
            DataSet ds = new DataSet();
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["connString"].ConnectionString))
            {

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = connection;
                cmd.CommandText = "p_Offline_GetS7SKUImageNames";
                cmd.CommandType = CommandType.StoredProcedure;
                SqlParameter paramID = new SqlParameter("@XML", root.ToString());
                cmd.Parameters.AddRange(new SqlParameter[] { paramID });
                cmd.CommandTimeout = 0;

                connection.Open();
                var adapter = new SqlDataAdapter(cmd);
                adapter.Fill(ds);
            }

            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                return ds.Tables[0];
            else
                return null;
        }

        private void SaveImagesToOffline(DataTable dt, HSSFWorkbook book, string imageFilePath)
        {
            var imageFolder = ConfigurationManager.AppSettings["ImageFolder"];
            var logoPath = ConfigurationManager.AppSettings["LogoPath"];

            List<ISheet> sheets = new List<ISheet>();
            var idx = book.GetSheetIndex("Order Template");
            var sht = (HSSFSheet)book.GetSheetAt(idx);
            sheets.Add(sht);

            if (TopDeptGroupStyle)
            {
                //sheets.Clear();

                var sheetNum = book.NumberOfSheets;
                DateTime t;
                for (int i = 0; i < sheetNum; i++)
                {
                    var st = (HSSFSheet)book.GetSheetAt(i);
                    if (DateTime.TryParse(st.SheetName, out t))
                    {
                        sheets.Add(st);
                        //break;
                    }
                }
            }

            Dictionary<int, int[]> collect = new Dictionary<int, int[]>();
            foreach (ISheet st in sheets)
            {
                byte[] bytes = System.IO.File.ReadAllBytes(logoPath);
                int pictureIdx = book.AddPicture(bytes, PictureType.JPEG);

                HSSFPatriarch patriarch = (HSSFPatriarch)st.CreateDrawingPatriarch();
                HSSFClientAnchor anchor = new HSSFClientAnchor(50 * (TopDeptGroupStyle ? 1 : 3), 50 * (TopDeptGroupStyle ? 1 : 3), (TopDeptGroupStyle ? 500 : 600), 100, 0, 0, 1, (TopDeptGroupStyle ? 2 : 3));
                HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);

                foreach (DataRow row in dt.Rows)
                {
                    var sku = row["SKU"].ToString();
                    var rowIndex = (int)row["ROW"];
                    var imageName = row["ImageName"].ToString();
                    var imagePath = Path.Combine(imageFolder, imageName.Replace(".jpg", ".png"));
                    SaveImageInSheet(book, st, imagePath, rowIndex, 0, collect);
                }
            }

            var index = book.GetSheetIndex("GlobalVariables");
            var sheet = (HSSFSheet)book.GetSheetAt(index);
            var srow = sheet.GetRow(24) == null ? (HSSFRow)sheet.CreateRow(24) : (HSSFRow)sheet.GetRow(24);
            var cell = srow.GetCell(1) == null ? (HSSFCell)srow.CreateCell(1) : (HSSFCell)srow.GetCell(1);
            cell.SetCellValue("1");

            var orderFile = new FileStream(imageFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            book.Write(orderFile);
            orderFile.Close();
            //book.Dispose();
        }

        public static void SaveImageInSheet(HSSFWorkbook book, ISheet sheet, string imagePath, int row, int column, Dictionary<int, int[]> collect)
        {
            if (collect.ContainsKey(row))
            {
                CreateImg(sheet, row, column, collect[row][0], collect[row][1]);
                return;
            }

            if (!File.Exists(imagePath))
                return;
            byte[] bytes = System.IO.File.ReadAllBytes(imagePath);
            int pictureIdx = book.AddPicture(bytes, PictureType.JPEG);

            int height = 20;
            using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                System.Drawing.Image image = System.Drawing.Image.FromStream(fs);
                height = image.Height * 7 / 9;
            }
            collect.Add(row, new int[] { height, pictureIdx });

            CreateImg(sheet, row, column, height, pictureIdx);
        }

        private static void CreateImg(ISheet sheet, int row, int column, int height, int pictureIdx)
        {

            var srow = (HSSFRow)sheet.GetRow(row);

            srow.HeightInPoints = (float)height;

            HSSFPatriarch patriarch = null;
            try
            {
                patriarch = (HSSFPatriarch)sheet.DrawingPatriarch;
            }
            catch { }
            if (patriarch == null)
                patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

            //add a picture
            HSSFClientAnchor anchor = new HSSFClientAnchor(10 * 10, 10, 1023, 255, column, row, column, row);
            HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
            //pict.Resize();
        }

        public static DataTable GetExclusiveStyles(string catalog)
        {
            DataSet ds = new DataSet();
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["connString"].ConnectionString))
            {

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = connection;
                cmd.CommandText = "p_Offline_ExclusiveStyles";
                cmd.CommandType = CommandType.StoredProcedure;
                SqlParameter paramID = new SqlParameter("@catcd", catalog);
                cmd.Parameters.AddRange(new SqlParameter[] { paramID });

                connection.Open();
                var adapter = new SqlDataAdapter(cmd);
                adapter.Fill(ds);
            }

            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                return ds.Tables[0];
            else
                return null;
        }

        private DataTable GetCustomerData(string soldto, string catalog)
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
                    WriteToLog(string.Format("p_Offline_TabularCustomerList | {0} {1} {2}", soldto, catalog, ex.Message));
                    return null;
                }
            }
        }
        #endregion
    }

    public static class ExcelConvert
    {
        #region - Convert ColumnIdx to Column Letter -

        public static int ToIndex(this string columnName)
        {
            if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+")) { throw new Exception("invalid parameter"); }

            int index = 0;
            char[] chars = columnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index - 1;
        }


        public static string ToName(this int index)
        {
            if (index < 0) { throw new Exception("invalid parameter"); }

            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)'A')).ToString());
                index = (int)((index - index % 26) / 26);
            } while (index > 0);

            return String.Join(string.Empty, chars.ToArray());
        }
        #endregion
    }
}