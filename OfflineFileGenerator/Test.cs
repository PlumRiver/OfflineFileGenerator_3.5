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

    public class App
    {
        public bool IncludePricing = true;
        public bool OrderShipDate = false;

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

        public Hashtable StyleCollection = new Hashtable();

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
                font.Boldweight = fontWeight == null ? (short)FontBoldWeight.NORMAL : (short)fontWeight;
                font.FontHeightInPoints = fontHeight == null ? (short)8 : (short)fontHeight;
                cellStyle.SetFont(font);
            }
            if(align != null)
                cellStyle.Alignment = (HorizontalAlignment)align;
            if(valign != null)
                cellStyle.VerticalAlignment = (VerticalAlignment)valign;
            if(borderBottom != null)
                cellStyle.BorderBottom = (CellBorderType)borderBottom;
            if(borderLeft != null)
                cellStyle.BorderLeft = (CellBorderType)borderLeft;
            if(borderRight != null)
                cellStyle.BorderRight = (CellBorderType)borderRight;
            if(borderTop != null)
                cellStyle.BorderTop = (CellBorderType)borderTop;
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
                case "m19009468":
                    cellStyle = SetCellStyle(book, "m19009468", true, FontBoldWeight.BOLD, 8, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                    CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, null);
                    break;

                case "m19009478":
                    cellStyle = SetCellStyle(book, "m19009478", true, FontBoldWeight.BOLD, 8, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                        CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, null, null);
                    break;
                case "s26":
                    cellStyle = SetCellStyle(book, "s26", true, FontBoldWeight.BOLD, 8, null, null,
                        CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, 22, null);
                    break;
                case "s27":
                    cellStyle = SetCellStyle(book, "s27", true, FontBoldWeight.BOLD, 8, null, HorizontalAlignment.RIGHT,
                CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, 22, "#,##0\\ [$kr-41D]");
                    break;
                case "s28":
                    cellStyle = SetCellStyle(book, "s28", true, FontBoldWeight.BOLD, 8, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, 22, null);
                    break;
                case "s29":
                    cellStyle = SetCellStyle(book, "s29", true, FontBoldWeight.BOLD, 8, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                CellBorderType.MEDIUM, null, null, CellBorderType.MEDIUM, 22, null);
                    cellStyle.WrapText = true;
                    break;
                case "s43":
                    cellStyle = cellStyle = SetCellStyle(book, "s43", true, FontBoldWeight.NORMAL, 8, null, null,
                null, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.THIN, null, null); ;
                    break;
                case "s24":
                    cellStyle = SetCellStyle(book, "s24", true, FontBoldWeight.NORMAL, 8, null, null,
                null, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.THIN, null, null); ;
                    break;
                case "s38":
                    cellStyle = SetCellStyle(book, "s38", true, FontBoldWeight.NORMAL, 8, null, null,
                null, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, null); ;
                    break;
                case "s40":
                    cellStyle = cellStyle = SetCellStyle(book, "s40", true, FontBoldWeight.NORMAL, 8, null, HorizontalAlignment.RIGHT,
                CellBorderType.THIN, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, "0.00"); ;
                    break;
                case "s41":
                    cellStyle = cellStyle = SetCellStyle(book, "s41", true, FontBoldWeight.BOLD, 8, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, CellBorderType.MEDIUM, null, null); ;
                    break;
                case "s81":
                    cellStyle = cellStyle = SetCellStyle(book, "s81", true, FontBoldWeight.NORMAL, 8, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
                CellBorderType.THIN, CellBorderType.THIN, CellBorderType.THIN, CellBorderType.MEDIUM, 22, null); ;
                    break;
                case "s62":
                    cellStyle = cellStyle = SetCellStyle(book, "s62", true, FontBoldWeight.NORMAL, 8, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER,
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
            //
            AddFormatCellsHeader(formatcellssheet);
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
            HSSFRow Row0 = (HSSFRow)sheet.CreateRow(0);
            Row0.HeightInPoints = 12;
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
            AddSecondaryRows(12, skusheet, pricesheet, atpsheet, 0);
            // -----------------------------------------------
            HSSFRow Row1 = (HSSFRow)sheet.CreateRow(1);
            Row1.HeightInPoints = 12;
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
            AddSecondaryRows(12, skusheet, pricesheet, atpsheet, 1);
            // -----------------------------------------------
            #region newcode
            ListDictionary colpositions = null;
            var sheetRow = 2;
            WriteToLog("CreateGridRows2 begin:" + DateTime.Now.TimeOfDay);
            XElement xroot = XElement.Load(TemplateXMLFilePath);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.EndElement && reader.Name == "ProductLevel")
                {
                    //Build a grid
                    XmlReader currentNode = reader.ReadSubtree();
                    int deptid = 0;
                    string deptname = "";
                    OrderedDictionary depts = GetDeptHierarchy(currentNode, ref deptid, ref deptname);

                    CreateGrid2(xroot, depts, deptid, sheet, skusheet, pricesheet, atpsheet, ref colpositions, soldto, catalog);
                    CreateGridRows2(depts, sheet, skusheet, pricesheet, atpsheet, deptid, deptname, ref colpositions);
                }
            }
            WriteToLog("CreateGridRows2 end:" + DateTime.Now.TimeOfDay);
            sheetRow = sheet.LastRowNum + 1;
            HSSFRow row = (HSSFRow)sheet.CreateRow(sheetRow);
            row.CreateCell(0).SetCellValue("EOF");
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

        private void AddSecondaryRows(short rowHeight, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, int rowIndex)
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

        private void CreateGrid2(XElement xroot, OrderedDictionary depts, int deptid, HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, ref ListDictionary colpositions, string soldto, string catalog)
        {
            //Create a new grid
            OrderedDictionary deptcols = null;
            bool bRectangular = ConfigurationManager.AppSettings["AttributeRectangular"] == null ? true : (ConfigurationManager.AppSettings["AttributeRectangular"] == "1" ? true : false);
            if (!bRectangular) deptcols = GetDeptCols(deptid, soldto, catalog);
            else deptcols = GetDeptColsFromXML(xroot, deptid);
            AddBlankHeader(ref deptcols, deptid);

            if (deptcols != null && deptcols.Count > 0)
            {
                colpositions = AddHeader2(depts, deptcols, sheet, skusheet, pricesheet, atpsheet);
            }
        }

        private ListDictionary AddHeader2(OrderedDictionary depts, OrderedDictionary cols, HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet)
        {
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
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            bool firstrow = true;
            //Add a break between top-level departments
            var sheetRow = sheet.LastRowNum + 1;
            HSSFRow newrow = (HSSFRow)sheet.CreateRow(sheetRow);
            newrow.HeightInPoints = 12;
            AddSecondaryRows(12, skusheet, pricesheet, atpsheet, sheetRow);
            sheetRow++;
            newrow = (HSSFRow)sheet.CreateRow(sheetRow);
            newrow.HeightInPoints = 12;
            newrow.CreateCell(0).SetCellValue("pagebreak");
            AddSecondaryRows(12, skusheet, pricesheet, atpsheet, sheetRow);
            sheetRow++;
            newrow = (HSSFRow)sheet.CreateRow(sheetRow);
            newrow.HeightInPoints = 12;
            AddSecondaryRows(12, skusheet, pricesheet, atpsheet, sheetRow);
            sheetRow++;

            HSSFRow hdrRow = (HSSFRow)sheet.CreateRow(sheetRow);
            hdrRow.HeightInPoints = 12;
            AddSecondaryRows(12, skusheet, pricesheet, atpsheet, sheetRow);
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
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                cell.SetCellValue(deptlabel);
            }
            else   //Lay the dept hierarchy out w/each level in its own column
            {
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                cell.SetCellValue(column1heading);
                foreach (DictionaryEntry de in depts)
                {
                    if (!firstrow)
                    {
                        cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
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
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
                        cell.SetCellValue("");
                        cellIndex++;
                    }
                }
            }
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
            cell.SetCellValue(column4heading);
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
            cell.SetCellValue(column5heading);
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s26");
            cell.SetCellValue(column6heading);
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s27");
            cell.SetCellValue(column7heading);
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s28");
            cell.SetCellValue(column8heading);
            cellIndex++;
            cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s28");
            cell.SetCellValue(column9heading);
            cellIndex++;
            foreach (DictionaryEntry de in cols)
            {
                colcounter++;
                cell = (HSSFCell)hdrRow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s29");
                cell.SetCellValue(string.IsNullOrEmpty(Convert.ToString(de.Value)) ? " " : StripHTML(Convert.ToString(de.Value)).Replace("\r", System.Environment.NewLine));
                cellIndex++;
                colpositions.Add(StripHTML(Convert.ToString(de.Value)), colcounter);
            }
            return colpositions;
        }

        public string StripHTML(string HTMLText)
        {
            var reg = new Regex("<[^>]+>", RegexOptions.IgnoreCase);
            return reg.Replace(HTMLText.Replace("<br>", "\r").Replace("<br/>", "\r").Replace("<p/>", "\r").Replace("</p>", System.Environment.NewLine).Replace("&nbsp;", " "), "");
        }

        private void CreateGridRows2(OrderedDictionary depts, HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, int deptid, string deptname, ref ListDictionary colpositions)
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
                    AddRow2(depts, rowattribs, skus, colpositions, sheet, skusheet, pricesheet, atpsheet, firstrow, lastrow);
                    firstrow = false;
                }
            }
        }

        private void AddRow2(OrderedDictionary DeptLevels, StringDictionary rowattribs, OrderedDictionary skus, ListDictionary colpositions, HSSFSheet sheet, HSSFSheet skusheet, HSSFSheet pricesheet, HSSFSheet atpsheet, bool firstrow, bool lastrow)
        {
            string errMsg = "";
            var rowIndex = sheet.LastRowNum + 1;
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
            newrow.HeightInPoints = 12;
            newrowa.HeightInPoints = 12;
            newrowb.HeightInPoints = 12;
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
                        if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
                        cell.SetCellValue(de.Value.ToString());
                        cellIndex++;
                    }
                    else
                    {
                        cell = (HSSFCell)newrow.CreateCell(cellIndex);
                        cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s43");
                        if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
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
                            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
                            cell.SetCellValue("");
                            cellIndex++;
                        }
                        else
                        {
                            cell = (HSSFCell)newrow.CreateCell(cellIndex);
                            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s43");
                            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
                            cell.SetCellValue("");
                            cellIndex++;
                        }
                    }
                }
            }
            //Mat.
            cell = (HSSFCell)newrow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s38");
            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
            cell.SetCellValue(rowattribs["Style"].ToString());
            cellIndex++;
            //newrow.Cells.Add(rowattribs["Style"].ToString(), DataType.String, "s38");
            //Mat Desc
            cell = (HSSFCell)newrow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s38");
            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
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
            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
            cell.SetCellValue(rowattribs["GridAttributeValues"].ToString());
            cellIndex++;
            //newrow.Cells.Add(rowattribs["GridAttributeValues"].ToString(), DataType.String, "s38");
            //WHS $
            cell = (HSSFCell)newrow.CreateCell(cellIndex);
            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s38");
            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
            cell.CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            var price = this.IncludePricing ? double.Parse(rowattribs["RowPriceWholesale"].ToString()) : double.Parse(ExcludePriceValue); // 0;
            cell.SetCellValue(price);
            cellIndex++;

            //TTL
            var TTLIndex = cellIndex;
            HSSFCell ttlcell = (HSSFCell)newrow.CreateCell(cellIndex);
            ttlcell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s41");
            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
            cellIndex++;
            //ttlcell.Data.Type = DataType.Number;

            //TTL Value
            HSSFCell valuecell = (HSSFCell)newrow.CreateCell(cellIndex);
            valuecell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s40");
            if (lastrow) { valuecell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
            cellIndex++;
            //valuecell.Data.Type = DataType.Number;

            //Lay out the row with empty, greyed-out cells
            for (int i = 0; i < colpositions.Count; i++)
            {
                cell = (HSSFCell)newrow.CreateCell(cellIndex);
                cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s81");
                if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
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
                        if (colpositions.Contains(dav))
                        {
                            int colposition = Convert.ToInt32(colpositions[dav].ToString());
                            if (!multicolumndeptlabel)
                                cell = (HSSFCell)newrow.GetCell(Convert.ToInt32(colpositions[dav].ToString()) + 6);
                            else
                                cell = (HSSFCell)newrow.GetCell(Convert.ToInt32(colpositions[dav].ToString()) + maxcols + 5);
                            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s62");
                            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
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
                            if(OrderShipDate)
                                cell = SetCellValue(newrowc, Convert.ToInt32(colpositions[dav].ToString()) + offset, ((StringDictionary)(de.Value))["ATPDate"].ToString(), false);
                            //Plant the wholesale price on its sheet
                            cell = SetCellValue(newrowb, Convert.ToInt32(colpositions[dav].ToString()) + offset, ((StringDictionary)(de.Value))["PriceWholesale"].ToString(), true);
                            cell.CellStyle = GenerateStyle((HSSFWorkbook)sheet.Workbook, "s62");
                            if (lastrow) { cell.CellStyle.BorderBottom = CellBorderType.MEDIUM; }
                            //Build the ttl cell formula
                            ttlformula += GetExcelColumnName(TTLIndex + 1 + colposition + 1) + (rowIndex + 1).ToString() + "+";
                            //"RC[" + Convert.ToString(colposition + 1) + "]+";
                            //Build the value cell formula
                            valueformula += "(" + GetExcelColumnName(TTLIndex + 1 + colposition + 1) + (rowIndex + 1).ToString() +
                                "*WholesalePrice!" + GetExcelColumnName(TTLIndex + 1 + colposition + 1) + (rowIndex + 1).ToString() + ")" + "+";
                            //"(RC[" + Convert.ToString(colposition) + "]*WholesalePrice!RC[" + Convert.ToString(colposition) + "])+";
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

        private OrderedDictionary GetDeptHierarchy(XmlReader reader, ref int deptid, ref string deptname)
        {
            OrderedDictionary depts = new OrderedDictionary();
            reader.Read();
            reader.MoveToNextAttribute();
            deptid = Convert.ToInt32(reader.Value);
            reader.MoveToNextAttribute();
            reader.MoveToNextAttribute();
            deptname = reader.Value;

            string parentdeptname = "";
            while (reader.Read())
            {
                if (reader.Name == "ParentLevel")
                {
                    reader.MoveToNextAttribute();
                    parentdeptname = reader.Value;
                    reader.MoveToNextAttribute();
                    depts.Add(reader.Value, parentdeptname);
                }
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
            string connString = ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand("p_GetOfflineOrderFormTemplateByBatchNo", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@BatchNo", batchNo));

            //Add params 
            SqlDataReader reader = null;
            try
            {
                conn.Open();
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }

            catch (Exception ex)
            {
                errMsg = ex.Message;
                WriteToLog("p_GetOfflineOrderFormTemplateByBatchNo", ex, string.Empty);
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
                throw ex;
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
            TemplateXMLFilePath = Path.Combine(xmlsavepath, Program.CleanFileName(soldto.Replace("\\", "-").Replace("/", "-") + "_" + catalog.Replace("\\", "-").Replace("/", "-") + "_" + "Template.xml"));
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
        public void LogOfflineOrd33erFormTemplate(string batchNo, string codeId)
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
        }



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

        public OrderedDictionary GetDeptCols(int deptid, string soldto, string catalog)
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

        public XmlReader GetProducts(int deptid, string catalog, string soldto)
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




        #region OLD CODE

        /*
        public int Generate(string filename, string soldto, string catalog, string pricecode, string savedirectory)
        {
            //Make sure data doesn't carry across multiple catalogs
            departments.Clear();
            cols.Clear();
            multiples.Clear();
            locked.Clear();
            unlocked.Clear();
            Workbook book = new Workbook();
            //Properties global to the spreadsheet
            // -----------------------------------------------
            //  Properties
            // -----------------------------------------------
            book.Properties.Author = "Tessy Smith";
            book.Properties.LastAuthor = "Eric Bradford";
            book.Properties.Created = DateTime.Now;
            book.Properties.LastSaved = DateTime.Now;
            book.Properties.Company = "PlumRiver Software";
            book.Properties.Version = "10.6811";
            book.ExcelWorkbook.WindowHeight = 8580;
            book.ExcelWorkbook.WindowWidth = 11340;
            book.ExcelWorkbook.WindowTopX = 240;
            book.ExcelWorkbook.WindowTopY = 45;
            book.ExcelWorkbook.ProtectWindows = false;
            book.ExcelWorkbook.ProtectStructure = false;
            // -----------------------------------------------
            //  Generate Styles
            // -----------------------------------------------
            this.GenerateStyles(book.Styles);
            // -----------------------------------------------
            //  Generate Order Template Worksheet
            // -----------------------------------------------
            maxcols = this.GenerateWorksheetOrderTemplate(book.Worksheets, soldto, catalog, pricecode, savedirectory);
            if (maxcols > -1)
            {
                //This saves the XML version of the generated spreadsheet for use by the next
                //step in the process
                book.Save(filename);
                return maxcols;
            }
            else
                return -1;
        }

        private void GenerateStyles(WorksheetStyleCollection styles)
        {
            //This registers all of the spreadsheet styles used
            // -----------------------------------------------
            //  Default
            // -----------------------------------------------
            WorksheetStyle Default = styles.Add("Default");
            Default.Name = "Normal";
            Default.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            // -----------------------------------------------
            //  m19009468
            // -----------------------------------------------
            WorksheetStyle m19009468 = styles.Add("m19009468");
            m19009468.Font.Bold = true;
            m19009468.Font.Size = 8;
            m19009468.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            m19009468.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19009468.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            m19009468.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            m19009468.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            m19009468.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  m19009478
            // -----------------------------------------------
            WorksheetStyle m19009478 = styles.Add("m19009478");
            m19009478.Font.Bold = true;
            m19009478.Font.Size = 8;
            m19009478.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            m19009478.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            m19009478.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            m19009478.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            m19009478.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            m19009478.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s23
            // -----------------------------------------------
            WorksheetStyle s23 = styles.Add("s23");
            // -----------------------------------------------
            //  s24
            // -----------------------------------------------
            WorksheetStyle s24 = styles.Add("s24");
            s24.Font.Size = 8;
            s24.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s24.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s24.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            // -----------------------------------------------
            //  s25
            // -----------------------------------------------
            WorksheetStyle s25 = styles.Add("s25");
            s25.Font.Size = 8;
            // -----------------------------------------------
            //  s26
            // -----------------------------------------------
            WorksheetStyle s26 = styles.Add("s26");
            s26.Font.Bold = true;
            s26.Font.Size = 8;
            s26.Interior.Color = "#C0C0C0";
            s26.Interior.Pattern = StyleInteriorPattern.Solid;
            s26.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s26.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s26.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s26.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s27
            // -----------------------------------------------
            WorksheetStyle s27 = styles.Add("s27");
            s27.Font.Bold = true;
            s27.Font.Size = 8;
            s27.Alignment.Horizontal = StyleHorizontalAlignment.Right;
            s27.Interior.Color = "#C0C0C0";
            s27.Interior.Pattern = StyleInteriorPattern.Solid;
            s27.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s27.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s27.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s27.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            s27.NumberFormat = "#,##0\\ [$kr-41D]";
            // -----------------------------------------------
            //  s28
            // -----------------------------------------------
            WorksheetStyle s28 = styles.Add("s28");
            s28.Font.Bold = true;
            s28.Font.Size = 8;
            s28.Interior.Color = "#C0C0C0";
            s28.Interior.Pattern = StyleInteriorPattern.Solid;
            s28.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s28.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s28.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s28.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s28.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s28.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s29
            // -----------------------------------------------
            WorksheetStyle s29 = styles.Add("s29");
            s29.Font.Bold = true;
            s29.Font.Size = 8;
            s29.Interior.Color = "#C0C0C0";
            s29.Interior.Pattern = StyleInteriorPattern.Solid;
            s29.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s29.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s29.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s29.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s30
            // -----------------------------------------------
            WorksheetStyle s30 = styles.Add("s30");
            s30.Font.Bold = true;
            s30.Font.Size = 8;
            s30.Interior.Color = "#C0C0C0";
            s30.Interior.Pattern = StyleInteriorPattern.Solid;
            s30.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s30.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s30.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s30.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s30.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s31
            // -----------------------------------------------
            WorksheetStyle s31 = styles.Add("s31");
            s31.Font.Bold = true;
            s31.Font.Size = 8;
            s31.Interior.Color = "#C0C0C0";
            s31.Interior.Pattern = StyleInteriorPattern.Solid;
            s31.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s31.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s31.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s31.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s31.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s32
            // -----------------------------------------------
            WorksheetStyle s32 = styles.Add("s32");
            s32.Font.Size = 8;
            // -----------------------------------------------
            //  s33
            // -----------------------------------------------
            WorksheetStyle s33 = styles.Add("s33");
            s33.Font.Size = 8;
            s33.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s34
            // -----------------------------------------------
            WorksheetStyle s34 = styles.Add("s34");
            s34.Font.Bold = true;
            s34.Font.Size = 8;
            s34.Interior.Color = "#C0C0C0";
            s34.Interior.Pattern = StyleInteriorPattern.Solid;
            s34.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s34.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s34.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s34.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s35
            // -----------------------------------------------
            WorksheetStyle s35 = styles.Add("s35");
            s35.Font.Bold = true;
            s35.Font.Size = 8;
            s35.Interior.Color = "#C0C0C0";
            s35.Interior.Pattern = StyleInteriorPattern.Solid;
            s35.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s35.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s35.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s36
            // -----------------------------------------------
            WorksheetStyle s36 = styles.Add("s36");
            s36.Font.Bold = true;
            s36.Font.Size = 8;
            s36.Interior.Color = "#C0C0C0";
            s36.Interior.Pattern = StyleInteriorPattern.Solid;
            s36.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s36.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s36.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s36.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s37
            // -----------------------------------------------
            WorksheetStyle s37 = styles.Add("s37");
            s37.Font.Bold = true;
            s37.Font.Size = 8;
            s37.Interior.Color = "#C0C0C0";
            s37.Interior.Pattern = StyleInteriorPattern.Solid;
            s37.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s37.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s37.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s37.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s37.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s38
            // -----------------------------------------------
            WorksheetStyle s38 = styles.Add("s38");
            s38.Font.Size = 8;
            s38.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s38.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s38.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s39
            // -----------------------------------------------
            WorksheetStyle s39 = styles.Add("s39");
            s39.Font.Size = 8;
            s39.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s39.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s40
            // -----------------------------------------------
            WorksheetStyle s40 = styles.Add("s40");
            s40.Font.Size = 8;
            s40.Alignment.Horizontal = StyleHorizontalAlignment.Right;
            s40.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1, "#000000");
            s40.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2, "#000000");
            s40.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2, "#000000");
            s40.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            //s40.NumberFormat = "#,##0\\ [$kr-41D]";
            s40.NumberFormat = "0.00";
            // -----------------------------------------------
            //  s41
            // -----------------------------------------------
            WorksheetStyle s41 = styles.Add("s41");
            s41.Font.Bold = true;
            s41.Font.Size = 8;
            s41.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s41.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s41.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s41.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s41.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s41.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s42
            // -----------------------------------------------
            WorksheetStyle s42 = styles.Add("s42");
            s42.Font.Bold = true;
            s42.Font.Size = 8;
            s42.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s42.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s42.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s42.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            s42.NumberFormat = "#,##0";
            // -----------------------------------------------
            //  s43
            // -----------------------------------------------
            WorksheetStyle s43 = styles.Add("s43");
            s43.Font.Size = 8;
            s43.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s43.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s43.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "Background");
            // -----------------------------------------------
            //  s44
            // -----------------------------------------------
            WorksheetStyle s44 = styles.Add("s44");
            s44.Font.Size = 8;
            s44.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s44.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            // -----------------------------------------------
            //  s45
            // -----------------------------------------------
            WorksheetStyle s45 = styles.Add("s45");
            s45.Font.Size = 8;
            s45.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s45.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            s45.NumberFormat = "#,##0\\ [$kr-41D]";
            // -----------------------------------------------
            //  s46
            // -----------------------------------------------
            WorksheetStyle s46 = styles.Add("s46");
            s46.Font.Size = 8;
            s46.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s46.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            s46.NumberFormat = "#,##0\\ [$kr-41D]";
            // -----------------------------------------------
            //  s47
            // -----------------------------------------------
            WorksheetStyle s47 = styles.Add("s47");
            s47.Font.Size = 8;
            s47.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s47.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s47.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s47.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "Background");
            // -----------------------------------------------
            //  s48
            // -----------------------------------------------
            WorksheetStyle s48 = styles.Add("s48");
            s48.Font.Size = 8;
            s48.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s48.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s48.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s48.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            // -----------------------------------------------
            //  s49
            // -----------------------------------------------
            WorksheetStyle s49 = styles.Add("s49");
            s49.Font.Size = 8;
            s49.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s49.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s49.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            // -----------------------------------------------
            //  s50
            // -----------------------------------------------
            WorksheetStyle s50 = styles.Add("s50");
            s50.Font.Size = 8;
            s50.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s50.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s50.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            s50.NumberFormat = "#,##0\\ [$kr-41D]";
            // -----------------------------------------------
            //  s51
            // -----------------------------------------------
            WorksheetStyle s51 = styles.Add("s51");
            s51.Font.Size = 8;
            s51.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s51.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s52
            // -----------------------------------------------
            WorksheetStyle s52 = styles.Add("s52");
            s52.Font.Size = 8;
            s52.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "#000000");
            // -----------------------------------------------
            //  s53
            // -----------------------------------------------
            WorksheetStyle s53 = styles.Add("s53");
            s53.Font.Size = 8;
            s53.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s54
            // -----------------------------------------------
            WorksheetStyle s54 = styles.Add("s54");
            s54.Font.Size = 8;
            s54.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s54.NumberFormat = "#,##0\\ [$kr-41D]";
            // -----------------------------------------------
            //  s55
            // -----------------------------------------------
            WorksheetStyle s55 = styles.Add("s55");
            s55.Font.Size = 8;
            s55.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s55.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "#000000");
            s55.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1, "#000000");
            // -----------------------------------------------
            //  s56
            // -----------------------------------------------
            WorksheetStyle s56 = styles.Add("s56");
            s56.Font.Bold = true;
            s56.Font.Size = 8;
            s56.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s56.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s56.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s57
            // -----------------------------------------------
            WorksheetStyle s57 = styles.Add("s57");
            s57.Font.Bold = true;
            s57.Font.Size = 8;
            s57.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s57.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "Background");
            s57.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s58
            // -----------------------------------------------
            WorksheetStyle s58 = styles.Add("s58");
            s58.Font.Bold = true;
            s58.Font.Size = 8;
            s58.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s58.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "Background");
            s58.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s58.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s59
            // -----------------------------------------------
            WorksheetStyle s59 = styles.Add("s59");
            s59.Font.Bold = true;
            s59.Font.Size = 8;
            s59.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s59.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1, "Background");
            s59.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s59.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            s59.NumberFormat = "#,##0\\ [$kr-41D]";
            // -----------------------------------------------
            //  s60
            // -----------------------------------------------
            WorksheetStyle s60 = styles.Add("s60");
            s60.Font.Bold = true;
            s60.Font.Size = 8;
            s60.Interior.Color = "#FFFF00";
            s60.Interior.Pattern = StyleInteriorPattern.Solid;
            s60.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s60.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s60.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s60.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s60.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s60.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s61
            // -----------------------------------------------
            WorksheetStyle s61 = styles.Add("s61");
            s61.Font.Bold = true;
            s61.Font.Size = 8;
            s61.Interior.Color = "#FFFF00";
            s61.Interior.Pattern = StyleInteriorPattern.Solid;
            s61.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s61.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s61.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s61.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            s61.NumberFormat = "#,##0\\ [$kr-41D]";
            // -----------------------------------------------
            //  s62
            // -----------------------------------------------
            WorksheetStyle s62 = styles.Add("s62");
            s62.Font.Size = 8;
            s62.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s62.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s62.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s62.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s62.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s62.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s63
            // -----------------------------------------------
            WorksheetStyle s63 = styles.Add("s63");
            s63.Font.Size = 8;
            s63.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s63.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s63.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s63.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s63.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s63.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s64
            // -----------------------------------------------
            WorksheetStyle s64 = styles.Add("s64");
            s64.Font.Size = 8;
            s64.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s64.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s64.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s64.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s64.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s65
            // -----------------------------------------------
            WorksheetStyle s65 = styles.Add("s65");
            s65.Font.Size = 8;
            s65.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s65.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s65.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s65.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s65.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s65.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s66
            // -----------------------------------------------
            WorksheetStyle s66 = styles.Add("s66");
            s66.Font.Size = 8;
            s66.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s66.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s66.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s66.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s66.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s66.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s67
            // -----------------------------------------------
            WorksheetStyle s67 = styles.Add("s67");
            s67.Font.Size = 8;
            s67.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s67.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s67.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s67.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s67.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s67.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s68
            // -----------------------------------------------
            WorksheetStyle s68 = styles.Add("s68");
            s68.Font.Size = 8;
            s68.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s68.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s68.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s68.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s68.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s68.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s69
            // -----------------------------------------------
            WorksheetStyle s69 = styles.Add("s69");
            s69.Font.Size = 8;
            s69.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s69.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s69.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s69.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s69.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s70
            // -----------------------------------------------
            WorksheetStyle s70 = styles.Add("s70");
            s70.Font.Size = 8;
            s70.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s70.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s70.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s70.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s70.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s71
            // -----------------------------------------------
            WorksheetStyle s71 = styles.Add("s71");
            s71.Font.Size = 8;
            s71.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s71.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s71.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s71.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s71.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s71.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s72
            // -----------------------------------------------
            WorksheetStyle s72 = styles.Add("s72");
            s72.Font.Size = 8;
            s72.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s72.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s72.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s72.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s73
            // -----------------------------------------------
            WorksheetStyle s73 = styles.Add("s73");
            s73.Font.Size = 8;
            s73.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s73.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s73.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s73.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s74
            // -----------------------------------------------
            WorksheetStyle s74 = styles.Add("s74");
            s74.Font.Size = 8;
            s74.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s74.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s74.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s74.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s74.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s75
            // -----------------------------------------------
            WorksheetStyle s75 = styles.Add("s75");
            s75.Font.Size = 8;
            s75.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s75.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s75.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s75.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s75.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s75.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s76
            // -----------------------------------------------
            WorksheetStyle s76 = styles.Add("s76");
            s76.Font.Size = 8;
            s76.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s76.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s76.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s76.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s76.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s77
            // -----------------------------------------------
            WorksheetStyle s77 = styles.Add("s77");
            s77.Font.Size = 8;
            s77.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s77.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s77.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s77.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s77.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s78
            // -----------------------------------------------
            WorksheetStyle s78 = styles.Add("s78");
            s78.Font.Size = 8;
            s78.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s78.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s78.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s78.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s78.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 2);
            s78.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s79
            // -----------------------------------------------
            WorksheetStyle s79 = styles.Add("s79");
            s79.Font.Size = 8;
            s79.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s79.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s79.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 2);
            s79.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s80
            // -----------------------------------------------
            WorksheetStyle s80 = styles.Add("s80");
            s80.Font.Size = 8;
            s80.Interior.Color = "#C0C0C0";
            s80.Interior.Pattern = StyleInteriorPattern.Solid;
            s80.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s80.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s80.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s80.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s80.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s80.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s81
            // -----------------------------------------------
            WorksheetStyle s81 = styles.Add("s81");
            s81.Font.Size = 8;
            s81.Interior.Color = "#C0C0C0";
            s81.Interior.Pattern = StyleInteriorPattern.Solid;
            s81.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s81.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s81.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s81.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s81.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s81.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 2);
            // -----------------------------------------------
            //  s82
            // -----------------------------------------------
            WorksheetStyle s82 = styles.Add("s82");
            s82.Font.Size = 8;
            s82.Interior.Color = "#C0C0C0";
            s82.Interior.Pattern = StyleInteriorPattern.Solid;
            s82.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s82.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s82.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s82.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s82.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s83
            // -----------------------------------------------
            WorksheetStyle s83 = styles.Add("s83");
            s83.Font.Size = 8;
            s83.Interior.Color = "#C0C0C0";
            s83.Interior.Pattern = StyleInteriorPattern.Solid;
            s83.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s83.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s83.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 2);
            s83.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s83.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s83.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s84
            // -----------------------------------------------
            WorksheetStyle s84 = styles.Add("s84");
            s84.Font.Size = 8;
            s84.Interior.Color = "#C0C0C0";
            s84.Interior.Pattern = StyleInteriorPattern.Solid;
            s84.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s84.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s84.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s84.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s84.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            // -----------------------------------------------
            //  s85
            // -----------------------------------------------
            WorksheetStyle s85 = styles.Add("s85");
            s85.Font.Size = 8;
            s85.Interior.Color = "#C0C0C0";
            s85.Interior.Pattern = StyleInteriorPattern.Solid;
            s85.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s85.Alignment.Vertical = StyleVerticalAlignment.Bottom;
            s85.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s85.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
        }

        private int GenerateWorksheetOrderTemplate(WorksheetCollection sheets, string soldto, string catalog, string pricecode, string savedirectory)
        {
            //Master proc to gather and layout the data

            //Pull values out of the config file
            int column1width = Convert.ToInt32(ConfigurationManager.AppSettings["column1width"].ToString());
            int column2width = Convert.ToInt32(ConfigurationManager.AppSettings["column2width"].ToString());
            int column3width = Convert.ToInt32(ConfigurationManager.AppSettings["column3width"].ToString());
            int column4width = Convert.ToInt32(ConfigurationManager.AppSettings["column4width"].ToString());
            int column5width = Convert.ToInt32(ConfigurationManager.AppSettings["column5width"].ToString());
            int column6width = Convert.ToInt32(ConfigurationManager.AppSettings["column6width"].ToString());
            int column7width = Convert.ToInt32(ConfigurationManager.AppSettings["column7width"].ToString());
            int column8width = Convert.ToInt32(ConfigurationManager.AppSettings["column8width"].ToString());
            int column9width = Convert.ToInt32(ConfigurationManager.AppSettings["column9width"].ToString());
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
            Worksheet sheet = sheets.Add("Order Template");
            //SKU tab - hidden - used to create sku column in upload tab
            Worksheet skusheet = sheets.Add("WebSKU");
            //Wholesale price tab - hidden - used by price calculation formulas
            Worksheet pricesheet = sheets.Add("WholesalePrice");
            //Lists of cells to be locked and have order multiple restrictions placed
            Worksheet formatcellssheet = sheets.Add("CellFormats");
            //
            AddFormatCellsHeader(formatcellssheet);
            sheet.Protected = false;
            sheet.Table.FullColumns = 1;
            sheet.Table.FullRows = 1;
            if (multicolumndeptlabel)
            {
                for (int colctr = 0; colctr < maxcols; colctr++)
                {
                    sheet.Table.Columns.Add(column1width);
                    skusheet.Table.Columns.Add(column1width);
                    pricesheet.Table.Columns.Add(column1width);
                }
            }
            else
            {
                sheet.Table.Columns.Add(column1width);
                skusheet.Table.Columns.Add(column1width);
                pricesheet.Table.Columns.Add(column1width);
            }

            sheet.Table.Columns.Add(column2width);
            skusheet.Table.Columns.Add(column2width);
            pricesheet.Table.Columns.Add(column2width);

            sheet.Table.Columns.Add(column3width);
            skusheet.Table.Columns.Add(column3width);
            pricesheet.Table.Columns.Add(column3width);

            sheet.Table.Columns.Add(column4width);
            skusheet.Table.Columns.Add(column4width);
            pricesheet.Table.Columns.Add(column4width);

            sheet.Table.Columns.Add(column5width);
            skusheet.Table.Columns.Add(column5width);
            pricesheet.Table.Columns.Add(column5width);

            sheet.Table.Columns.Add(column6width);
            skusheet.Table.Columns.Add(column6width);
            pricesheet.Table.Columns.Add(column6width);

            sheet.Table.Columns.Add(column7width);
            skusheet.Table.Columns.Add(column7width);
            pricesheet.Table.Columns.Add(column7width);

            // -----------------------------------------------
            WorksheetRow Row0 = sheet.Table.Rows.Add();
            Row0.Height = 12;
            WorksheetCell cell;
            cell = Row0.Cells.Add();
            cell.StyleID = "m19009468";
            cell.Data.Type = DataType.String;
            cell.Data.Text = "Reqested Delivery date:";
            cell.Index = 2;
            cell.MergeAcross = 3;
            AddSecondaryRows(12, skusheet, pricesheet);
            // -----------------------------------------------
            WorksheetRow Row1 = sheet.Table.Rows.Add();
            Row1.Height = 12;
            cell = Row1.Cells.Add();
            cell.StyleID = "m19009478";
            cell.Data.Type = DataType.String;
            cell.Data.Text = "Customer name:";
            cell.Index = 2;
            cell.MergeAcross = 3;
            AddSecondaryRows(12, skusheet, pricesheet);
            // -----------------------------------------------
            #region newcode
            ListDictionary colpositions = null;

            WriteToLog("CreateGridRows2 begin:" + DateTime.Now.TimeOfDay);
            XElement xroot = XElement.Load(TemplateXMLFilePath);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.EndElement && reader.Name == "ProductLevel")
                {
                    //Build a grid
                    XmlReader currentNode = reader.ReadSubtree();
                    int deptid = 0;
                    string deptname = "";
                    OrderedDictionary depts = GetDeptHierarchy(currentNode, ref deptid, ref deptname);

                    CreateGrid2(xroot, depts, deptid, sheet, skusheet, pricesheet, ref colpositions, soldto, catalog);
                    CreateGridRows2(depts, sheet, skusheet, pricesheet, deptid, deptname, ref colpositions);
                }
            }
            WriteToLog("CreateGridRows2 end:" + DateTime.Now.TimeOfDay);
            WorksheetRow row = sheet.Table.Rows.Add();
            row.Cells.Add("EOF");

            WriteCellFormatValues(formatcellssheet);
            for (int i = 0; i < dataColumnNumber; i++)
            {
                sheet.Table.Columns.Add(datacolumnwidth);
            }

            #endregion

            // -----------------------------------------------
            WorksheetRow Row47 = sheet.Table.Rows.Add();
            // -----------------------------------------------
            //  Options
            // -----------------------------------------------
            sheet.Options.Selected = true;
            sheet.Options.FreezePanes = false;
            //sheet.Options.SplitHorizontal = 2;
            //sheet.Options.TopRowBottomPane = 2;
            //sheet.Options.ActivePane = 2;
            sheet.Options.ProtectObjects = true;
            sheet.Options.ProtectScenarios = true;
            sheet.Options.Print.ValidPrinterInfo = true;

            return maxcols;
        }

        private void CreateGrid(Worksheet sheet, Worksheet skusheet, Worksheet pricesheet, ref ListDictionary colpositions, ref int maxcols, SqlDataReader reader, ref bool newLevel1, ref bool newLevel2, string soldto, string catalog)
        {
            //Create a new grid
            int deptid = Convert.ToInt32(reader["dept_id"].ToString());
            OrderedDictionary deptcols = GetDeptCols(deptid, soldto, catalog);
            AddBlankHeader(ref deptcols, deptid);
            if (deptcols.Count > maxcols)
                maxcols = deptcols.Count;

            if (deptcols != null && deptcols.Count > 0)
            {
                colpositions = AddHeader(deptcols, sheet, skusheet, pricesheet);
                //Set newlevel1 flag even though it isn't, because we're just under a header
                newLevel1 = true;
                newLevel2 = true;
            }
        }

        private void CreateGrid2(XElement xroot, OrderedDictionary depts, int deptid, Worksheet sheet, Worksheet skusheet, Worksheet pricesheet, ref ListDictionary colpositions, string soldto, string catalog)
        {
            //Create a new grid
            OrderedDictionary deptcols = null;
            bool bRectangular = ConfigurationManager.AppSettings["AttributeRectangular"] == null ? true : (ConfigurationManager.AppSettings["AttributeRectangular"] == "1" ? true : false);
            if (!bRectangular) deptcols = GetDeptCols(deptid, soldto, catalog);
            else deptcols = GetDeptColsFromXML(xroot, deptid);
            AddBlankHeader(ref deptcols, deptid);

            if (deptcols != null && deptcols.Count > 0)
            {
                colpositions = AddHeader2(depts, deptcols, sheet, skusheet, pricesheet);
            }
        }

        private void CreateGridRows(Worksheet sheet, Worksheet skusheet, Worksheet pricesheet, SqlDataReader reader, string levelOne, string levelTwo, ref ListDictionary colpositions, ref bool newLevel1, ref bool newLevel2, ref bool newLevel3)
        {
            //Header has been created, now iterate through the product rows
            //and create one grid row per
            int deptid = Convert.ToInt32(reader["dept_id"].ToString());
            OrderedDictionary rows = (OrderedDictionary)departments[deptid];
            foreach (DictionaryEntry de in rows)
            {
                StringDictionary rowattribs = (StringDictionary)de.Key;
                OrderedDictionary skus = (OrderedDictionary)de.Value;
                AddRow(rowattribs, skus, colpositions, newLevel1, newLevel2, newLevel3, levelOne, levelTwo, reader["dept_description"].ToString(), sheet, skusheet, pricesheet);
                if (newLevel1)
                    newLevel1 = false;
                if (newLevel2)
                    newLevel2 = false;
                if (newLevel3)
                    newLevel3 = false;
            }
        }

        private void CreateGridRows2(OrderedDictionary depts, Worksheet sheet, Worksheet skusheet, Worksheet pricesheet, int deptid, string deptname, ref ListDictionary colpositions)
        {
            bool firstrow = true;
            OrderedDictionary rows = (OrderedDictionary)departments[deptid];

            if (rows != null)
            {
                foreach (DictionaryEntry de in rows)
                {
                    StringDictionary rowattribs = (StringDictionary)de.Key;
                    OrderedDictionary skus = (OrderedDictionary)de.Value;
                    AddRow2(depts, rowattribs, skus, colpositions, sheet, skusheet, pricesheet, firstrow);
                    firstrow = false;
                }
            }
        }

        private void WriteCellFormatValues(Worksheet formatcellssheet)
        {
            int ctr = 1;
            //Write previously stored order multiple restrictions in the cellformat sheet
            if (multiples.Count > unlocked.Count)
            {
                foreach (DictionaryEntry de in multiples)
                {
                    WorksheetRow row = formatcellssheet.Table.Rows.Add();
                    row.Cells.Add();
                    row.Cells.Add();
                    WorksheetCell cell = row.Cells.Add(Convert.ToString(de.Key));
                    cell = row.Cells.Add(Convert.ToString(de.Value));
                }
                //Write previously stored unlocked cell inventory in the cellformat sheet
                foreach (DictionaryEntry de in unlocked)
                {
                    formatcellssheet.Table.Rows[ctr].Cells[0].Data.Text = Convert.ToString(de.Key);
                    ctr++;
                }
            }
            else
            {
                foreach (DictionaryEntry de in unlocked)
                {
                    WorksheetRow row = formatcellssheet.Table.Rows.Add();
                    row.Cells.Add(Convert.ToString(de.Key));
                    row.Cells.Add();
                    row.Cells.Add();
                    row.Cells.Add();
                }
                foreach (DictionaryEntry de in multiples)
                {
                    formatcellssheet.Table.Rows[ctr].Cells[2].Data.Text = Convert.ToString(de.Key);
                    formatcellssheet.Table.Rows[ctr].Cells[3].Data.Text = Convert.ToString(de.Value);
                    ctr++;
                }
            }
        }

        private void AddRow(StringDictionary rowattribs, OrderedDictionary skus, ListDictionary colpositions, bool newlevel1, bool newlevel2, bool newlevel3, string levelone, string leveltwo, string levelthree, Worksheet sheet, Worksheet skusheet, Worksheet pricesheet)
        {
            //Add the row to each sheet to keep them in synch
            WorksheetRow newrow = sheet.Table.Rows.Add();
            WorksheetRow newrowa = skusheet.Table.Rows.Add();
            WorksheetRow newrowb = pricesheet.Table.Rows.Add();
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            //int offset = 5 + hierarchylevels;
            int offset = 8;
            WorksheetCell cell = null;
            string valueformula = "";
            string ttlformula = "";
            newrow.Height = 12;
            newrowa.Height = 12;
            newrowb.Height = 12;
            //Add spacer cells to the secondary sheets
            AddSecondaryCells(newrowa, newrowb, 9);
            newrow.AutoFitHeight = false;
            string wholesaleprice = "";
            if (this.IncludePricing)
            {
                //Strip a star from the end of wholesale price
                if (rowattribs["RowPriceWholesale"].ToString().Length == 6 && rowattribs["RowPriceWholesale"].ToString().Substring(5, 1) == "*")
                    wholesaleprice = rowattribs["RowPriceWholesale"].ToString().Substring(0, 5);
                else
                    wholesaleprice = rowattribs["RowPriceWholesale"].ToString();
                //If it's not a new top level, merge vertically - same idea for subsequent levels
            }
            else
            {
                wholesaleprice = "0";
            }
            if (newlevel1)
            {
                newrow.Cells.Add(levelone, DataType.String, "s24");
            }
            else
            {
                cell = newrow.Cells.Add();
                cell.StyleID = "s43";
            }
            if (newlevel2)
            {
                newrow.Cells.Add(leveltwo, DataType.String, "s38");
            }
            else
            {
                cell = newrow.Cells.Add();
                cell.StyleID = "s43";
            }
            if (newlevel3)
            {
                newrow.Cells.Add(levelthree, DataType.String, "s38");
            }
            else
            {
                cell = newrow.Cells.Add();
                cell.StyleID = "s43";
            }
            //Mat.
            newrow.Cells.Add(rowattribs["Style"].ToString(), DataType.String, "s38");
            //Mat Desc
            cell = newrow.Cells.Add(rowattribs["ProductName"].ToString(), DataType.String, "s38");
            //Dim1
            newrow.Cells.Add(rowattribs["GridAttributeValues"].ToString(), DataType.String, "s38");
            //WHS $
            newrow.Cells.Add(this.IncludePricing ? rowattribs["RowPriceWholesale"].ToString() : "0", DataType.Number, "s40");
            //TTL
            WorksheetCell ttlcell = newrow.Cells.Add();
            ttlcell.StyleID = "s41";
            ttlcell.Data.Type = DataType.Number;
            //cell.Formula = "=SUM(RC[2]:RC[" + Convert.ToString(colpositions.Count + 2) + "])";

            //TTL Value
            WorksheetCell valuecell = newrow.Cells.Add();
            valuecell.StyleID = "s40";
            valuecell.Data.Type = DataType.Number;
            //cell.Formula = "=RC[-1]*(1-R1C10)";

            //Lay out the row with empty, greyed-out cells
            for (int i = 0; i < colpositions.Count; i++)
            {
                cell = newrow.Cells.Add();
                cell.StyleID = "s81";
                //locked.Add(ConvertToLetter(i + 10) + Convert.ToString(sheet.Table.Rows.Count), "");
            }
            AddSecondaryCells(newrowa, newrowb, colpositions.Count);

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
                        if (colpositions.Contains(dav))
                        {
                            int colposition = Convert.ToInt32(colpositions[dav].ToString());
                            cell = newrow.Cells[Convert.ToInt32(colpositions[dav].ToString()) + offset];
                            cell.StyleID = "s62";
                            cellposition = ConvertToLetter(Convert.ToInt32(colpositions[dav].ToString()) + offset + 1) + Convert.ToString(sheet.Table.Rows.Count);
                            unlocked.Add(cellposition, cellposition);
                            multiples.Add(cellposition, ((StringDictionary)(de.Value))["OrderMultiple"].ToString());
                            //Plant the QuickWebSKU on its sheet
                            cell = newrowa.Cells[Convert.ToInt32(colpositions[dav].ToString()) + offset];
                            cell.Data.Type = DataType.String;
                            cell.Data.Text = ((StringDictionary)(de.Value))["QuickWebSKUValue"].ToString();
                            //cell.StyleID = "s62";
                            //Plant the wholesale price on its sheet
                            cell = newrowb.Cells[Convert.ToInt32(colpositions[dav].ToString()) + offset];
                            cell.Data.Text = ((StringDictionary)(de.Value))["PriceWholesale"].ToString();
                            cell.StyleID = "s62";
                            //Build the ttl cell formula
                            ttlformula += "RC[" + Convert.ToString(colposition + 1) + "]+";
                            //Build the value cell formula
                            valueformula += "(RC[" + Convert.ToString(colposition) + "]*WholesalePrice!RC[" + Convert.ToString(colposition) + "])+";
                        }
                    }
                }
            }
            catch (Exception ex)
            { throw; }
            //Assign the formulas
            if (ttlformula.Length > 0)
                ttlcell.Formula = ttlformula.Substring(0, ttlformula.Length - 1);
            if (valueformula.Length > 0)
                valuecell.Formula = valueformula.Substring(0, valueformula.Length - 1);
        }

        private void AddRow2(OrderedDictionary DeptLevels, StringDictionary rowattribs, OrderedDictionary skus, ListDictionary colpositions, Worksheet sheet, Worksheet skusheet, Worksheet pricesheet, bool firstrow)
        {
            string errMsg = "";
            //Add the row to each sheet to keep them in synch
            WorksheetRow newrow = sheet.Table.Rows.Add();
            WorksheetRow newrowa = skusheet.Table.Rows.Add();
            WorksheetRow newrowb = pricesheet.Table.Rows.Add();
            //int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            //int offset = 8;
            int offset = 6;
            if (DeptLevels != null && DeptLevels.Count > 0 && multicolumndeptlabel)
                offset = DeptLevels.Count + 5;
            WorksheetCell cell = null;
            string valueformula = "";
            string ttlformula = "";
            newrow.Height = 12;
            newrowa.Height = 12;
            newrowb.Height = 12;
            //Add spacer cells to the secondary sheets
            if (!multicolumndeptlabel)
                AddSecondaryCells(newrowa, newrowb, maxcols + 7);
            else
                AddSecondaryCells(newrowa, newrowb, maxcols + 6);
            newrow.AutoFitHeight = false;
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
                    wholesaleprice = "0";
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
            }
            //If it's not a new top level, merge vertically - same idea for subsequent levels
            int ictr = 0;
            if (!multicolumndeptlabel)
                newrow.Cells.Add("", DataType.String, "s43");
            else
            {
                foreach (DictionaryEntry de in DeptLevels)
                {
                    if (firstrow)
                    {
                        newrow.Cells.Add(de.Value.ToString(), DataType.String, "s24");
                    }
                    else
                    {
                        cell = newrow.Cells.Add();
                        cell.StyleID = "s43";
                    }
                    ictr++;
                }
                //Fill the rest of the dept column labels horizontally with blanks
                if (ictr < maxcols)
                {
                    for (int i = 0; i < maxcols - ictr; i++)
                    {
                        if (firstrow)
                            newrow.Cells.Add("", DataType.String, "s24");
                        else
                            newrow.Cells.Add("", DataType.String, "s43");
                    }
                }
            }
            //Mat.
            newrow.Cells.Add(rowattribs["Style"].ToString(), DataType.String, "s38");
            //Mat Desc
            cell = newrow.Cells.Add(rowattribs["ProductName"].ToString(), DataType.String, "s38");
            //Dim1
            newrow.Cells.Add(rowattribs["GridAttributeValues"].ToString(), DataType.String, "s38");
            //WHS $
            newrow.Cells.Add(this.IncludePricing ? rowattribs["RowPriceWholesale"].ToString() : "0", DataType.Number, "s40");
            //TTL
            WorksheetCell ttlcell = newrow.Cells.Add();
            ttlcell.StyleID = "s41";
            ttlcell.Data.Type = DataType.Number;

            //TTL Value
            WorksheetCell valuecell = newrow.Cells.Add();
            valuecell.StyleID = "s40";
            valuecell.Data.Type = DataType.Number;

            //Lay out the row with empty, greyed-out cells
            for (int i = 0; i < colpositions.Count; i++)
            {
                cell = newrow.Cells.Add();
                cell.StyleID = "s81";
            }
            AddSecondaryCells(newrowa, newrowb, offset + colpositions.Count);

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
                        if (colpositions.Contains(dav))
                        {
                            int colposition = Convert.ToInt32(colpositions[dav].ToString());
                            if (!multicolumndeptlabel)
                                cell = newrow.Cells[Convert.ToInt32(colpositions[dav].ToString()) + 6];
                            else
                                cell = newrow.Cells[Convert.ToInt32(colpositions[dav].ToString()) + maxcols + 5];
                            cell.StyleID = "s62";
                            if (multicolumndeptlabel)
                                cellposition = ConvertToLetter(Convert.ToInt32(colpositions[dav].ToString()) + maxcols + 6) + Convert.ToString(sheet.Table.Rows.Count);
                            else
                                cellposition = ConvertToLetter(Convert.ToInt32(colpositions[dav].ToString()) + 7) + Convert.ToString(sheet.Table.Rows.Count);
                            unlocked.Add(cellposition, cellposition);
                            if (((StringDictionary)(de.Value)).ContainsKey("OrderMultiple"))
                                multiples.Add(cellposition, ((StringDictionary)(de.Value))["OrderMultiple"].ToString());
                            //Plant the QuickWebSKU on its sheet
                            cell = newrowa.Cells[Convert.ToInt32(colpositions[dav].ToString()) + offset];
                            cell.Data.Type = DataType.String;
                            cell.Data.Text = ((StringDictionary)(de.Value))["QuickWebSKUValue"].ToString();
                            //Plant the wholesale price on its sheet
                            cell = newrowb.Cells[Convert.ToInt32(colpositions[dav].ToString()) + offset];
                            cell.Data.Type = DataType.Number;
                            cell.Data.Text = ((StringDictionary)(de.Value))["PriceWholesale"].ToString();
                            cell.StyleID = "s62";
                            //Build the ttl cell formula
                            ttlformula += "RC[" + Convert.ToString(colposition + 1) + "]+";
                            //Build the value cell formula
                            valueformula += "(RC[" + Convert.ToString(colposition) + "]*WholesalePrice!RC[" + Convert.ToString(colposition) + "])+";
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
        }

        private ListDictionary AddHeader(OrderedDictionary cols, Worksheet sheet, Worksheet skusheet, Worksheet pricesheet)
        {
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
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            //Add a break between top-level departments
            WorksheetRow newrow = sheet.Table.Rows.Add();
            newrow.Height = 12;
            newrow = sheet.Table.Rows.Add();
            newrow.Height = 12;
            newrow.Cells.Add("pagebreak");
            newrow = sheet.Table.Rows.Add();
            newrow.Height = 12;
            AddSecondaryRows(12, skusheet, pricesheet);
            AddSecondaryRows(12, skusheet, pricesheet);
            AddSecondaryRows(12, skusheet, pricesheet);

            WorksheetRow hdrRow = sheet.Table.Rows.Add();
            hdrRow.Height = 12;
            AddSecondaryRows(12, skusheet, pricesheet);
            hdrRow.Cells.Add(column1heading, DataType.String, "s26");
            hdrRow.Cells.Add(column2heading, DataType.String, "s26");
            hdrRow.Cells.Add(column3heading, DataType.String, "s26");
            hdrRow.Cells.Add(column4heading, DataType.String, "s26");
            hdrRow.Cells.Add(column5heading, DataType.String, "s26");
            hdrRow.Cells.Add(column6heading, DataType.String, "s26");
            hdrRow.Cells.Add(column7heading, DataType.String, "s27");
            hdrRow.Cells.Add(column8heading, DataType.String, "s28");
            hdrRow.Cells.Add(column9heading, DataType.String, "s28");
            foreach (DictionaryEntry de in cols)
            {
                colcounter++;
                WorksheetCell cell = hdrRow.Cells.Add(Convert.ToString(de.Value), DataType.String, "s29");
                //sheet.Table.Columns[cell.Index].Width = datacolumnwidth;
                colpositions.Add(Convert.ToString(de.Value), colcounter);
            }
            return colpositions;
        }

        private ListDictionary AddHeader2(OrderedDictionary depts, OrderedDictionary cols, Worksheet sheet, Worksheet skusheet, Worksheet pricesheet)
        {
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
            int hierarchylevels = Convert.ToInt32(ConfigurationManager.AppSettings["hierarchylevels"].ToString());
            bool firstrow = true;
            //Add a break between top-level departments
            WorksheetRow newrow = sheet.Table.Rows.Add();
            newrow.Height = 12;
            newrow = sheet.Table.Rows.Add();
            newrow.Height = 12;
            newrow.Cells.Add("pagebreak");
            newrow = sheet.Table.Rows.Add();
            newrow.Height = 12;
            AddSecondaryRows(12, skusheet, pricesheet);
            AddSecondaryRows(12, skusheet, pricesheet);
            AddSecondaryRows(12, skusheet, pricesheet);

            WorksheetRow hdrRow = sheet.Table.Rows.Add();
            hdrRow.Height = 12;
            AddSecondaryRows(12, skusheet, pricesheet);
            if (!multicolumndeptlabel)   //Put the whole dept hierarchy in the first header cell
            {
                string deptlabel = "> ";
                foreach (DictionaryEntry de in depts)
                    deptlabel += de.Value + " > ";
                deptlabel = deptlabel.Substring(0, deptlabel.Length - 3);
                hdrRow.Cells.Add(deptlabel, DataType.String, "s26");
            }
            else   //Lay the dept hierarchy out w/each level in its own column
            {
                hdrRow.Cells.Add(column1heading, DataType.String, "s26");
                foreach (DictionaryEntry de in depts)
                {
                    if (!firstrow)
                        hdrRow.Cells.Add("", DataType.String, "s26");
                    firstrow = false;
                }
                //Fill the rest of the dept column labels horizontally with blanks
                if (depts.Count < maxcols)
                {
                    for (int i = 0; i < maxcols - depts.Count; i++)
                    {
                        hdrRow.Cells.Add("", DataType.String, "s26");
                    }
                }
            }
            hdrRow.Cells.Add(column4heading, DataType.String, "s26");
            hdrRow.Cells.Add(column5heading, DataType.String, "s26");
            hdrRow.Cells.Add(column6heading, DataType.String, "s26");
            hdrRow.Cells.Add(column7heading, DataType.String, "s27");
            hdrRow.Cells.Add(column8heading, DataType.String, "s28");
            hdrRow.Cells.Add(column9heading, DataType.String, "s28");
            foreach (DictionaryEntry de in cols)
            {
                colcounter++;
                WorksheetCell cell = hdrRow.Cells.Add(Convert.ToString(de.Value), DataType.String, "s29");
                colpositions.Add(Convert.ToString(de.Value), colcounter);
            }
            return colpositions;
        }

        public SqlDataReader GetDepartments(string catalog, string pricecode)
        {
            string connString = System.Configuration.ConfigurationManager.ConnectionStrings["connString"].ConnectionString;
            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand("GetValidDeptsforCatPriceCd", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.Add(new SqlParameter("@catcd", SqlDbType.VarChar, 80));
            cmd.Parameters["@catcd"].Value = catalog;

            cmd.Parameters.Add(new SqlParameter("@pricecd", SqlDbType.VarChar, 80));
            cmd.Parameters["@pricecd"].Value = pricecode;

            //Add params 
            SqlDataReader reader = null;
            try
            {
                conn.Open();
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }

            catch
            {
                throw;
            }
            return reader;
        }

        private void AddSecondaryRows(int rowHeight, Worksheet skusheet, Worksheet pricesheet)
        {
            WorksheetRow newrow = skusheet.Table.Rows.Add();
            newrow.Height = rowHeight;
            newrow = pricesheet.Table.Rows.Add();
            newrow.Height = rowHeight;
        }

        private void AddSecondaryCells(WorksheetRow row1, WorksheetRow row2, int cellsToAdd)
        {
            for (int i = 0; i < cellsToAdd; i++)
            {
                row1.Cells.Add();
                row2.Cells.Add();
            }
        }

        private void AddFormatCellsHeader(Worksheet sheet)
        {
            WorksheetRow row = sheet.Table.Rows.Add();
            row.Cells.Add("Unlocked Cells");
            row.Cells.Add();
            row.Cells.Add("Multiple Cells");
            row.Cells.Add("Multiple Value");
        }

        private string FormulaString(int colcount)
        {
            string temp = "(RC[1]*WholesalePrice!RC[1])";
            for (int i = 0; i < colcount; i++)
            {
                temp += "+(RC[" + Convert.ToString(i + 2) + "]*WholesalePrice!RC[" + Convert.ToString(i + 2) + "])";
            }
            return temp;
        }

        public OrderedDictionary DeepCopy(OrderedDictionary sl)
        {
            OrderedDictionary newsl = new OrderedDictionary();
            foreach (DictionaryEntry de in sl)
            {
                newsl.Add(de.Key, de.Value);
            }
            return newsl;
        }
        */

        #endregion


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
            FileStream sourceFile = new FileStream(imageFilePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook book = new HSSFWorkbook(sourceFile);
            //get image list with first sku
            var skus = GetFirstSKUList(book);
            //generate images in offline
            var dt = GetSKUImageList(skus);
            if (dt != null)
            {
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
            var firstRow = 6;
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
            foreach (DataRow row in dt.Rows)
            {
                var sku = row["SKU"].ToString();
                var rowIndex = (int)row["ROW"];
                var imageName = row["ImageName"].ToString();
                var imagePath = Path.Combine(imageFolder, imageName.Replace(".jpg", ".png"));
                SaveImageInSheet(book, imagePath, rowIndex, 0);
            }
            var logoPath = ConfigurationManager.AppSettings["LogoPath"];
            byte[] bytes = System.IO.File.ReadAllBytes(logoPath);
            int pictureIdx = book.AddPicture(bytes, PictureType.JPEG);

            var index = book.GetSheetIndex("Order Creation");
            var sheet = (HSSFSheet)book.GetSheetAt(index);
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor(50 * 3, 50 * 3, 600, 100, 0, 0, 1, 3);
            HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);

            index = book.GetSheetIndex("GlobalVariables");
            sheet = (HSSFSheet)book.GetSheetAt(index);
            var srow = sheet.GetRow(24) == null ? (HSSFRow)sheet.CreateRow(24) : (HSSFRow)sheet.GetRow(24);
            var cell = srow.GetCell(1) == null ? (HSSFCell)srow.CreateCell(1) : (HSSFCell)srow.GetCell(1);
            cell.SetCellValue("1");

            var orderFile = new FileStream(imageFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            book.Write(orderFile);
            orderFile.Close();
            book.Dispose();
        }

        private void SaveImageInSheet(HSSFWorkbook book, string imagePath, int row, int column)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(imagePath);
            int pictureIdx = book.AddPicture(bytes, PictureType.JPEG);

            var index = book.GetSheetIndex("Order Template");
            var sheet = (HSSFSheet)book.GetSheetAt(index);

            var srow = (HSSFRow)sheet.GetRow(row);
            srow.HeightInPoints = 40;

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

        #endregion
    }

}