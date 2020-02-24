using System;
using System.IO;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace OfflineFileGenerator
{

    class Program
    {
        private static void WriteToLog(string msg)
        {
            //string templatefile = ConfigurationManager.AppSettings["templatefile"].ToString();
            var logFile = "OfflineLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            string savefolder = ConfigurationManager.AppSettings["savefolder"].ToString();
            StreamWriter sw = new StreamWriter(savefolder + logFile, true);
            sw.WriteLine(msg);
            sw.Close();
            Console.WriteLine(msg);
        }

        static void Main(string[] args)
        {
            //(new EPPlusTest()).GetCellValue();
            //2014/EQUIPMENT/BOOKINGCAD //2014/PRO-GOALIE/BOOKING //2014/LICENSED/REPEAR //2014/CLEAR-OUT/CL
//            new CategoryTabularFormGenerator().GenerateCategoryTabularOrderForm(
//@"D:\PRWork\Plumriver-Customers\Reebok\Offline\TabularOfflineOrderForm.xlsm", "testcategory.xlsm", "CB3", "2014/LICENSED/REPEAR", "2014/LICENSED/REPEARCAD", @"D:\PRWork\Plumriver-Customers\Reebok\Offline\");
//            new CategoryTabularFormGenerator().GenerateCategoryTabularOrderForm(
//ConfigurationManager.AppSettings["TabularTemplateFile"], "testcategory.xlsm", "CB3", "2014/LICENSED/REPEAR", "2014/LICENSED/REPEARCAD", ConfigurationManager.AppSettings["savefolder"]);
//            return; 
            //using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"D:\APPSetUp\NS_BKG17_002.xlsm")))//NS_BKG17.xlsm
            //{
            //    var a = "aa";
            //}


            //If the argument 0 is "1", then CMSGenerate is true.
            //If there're any argument with suffix ".xls" or ".xlsm", it will be regonized as Test Mode, and the file name will be assigned to variant testFormName.
            var CMSGenerate = false;
            string testFormName = string.Empty; // for test purpose
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].EndsWith(".xls") || args[i].EndsWith(".xlsm"))
                        testFormName = args[i];
                    WriteToLog(args[i].Trim());
                }
                CMSGenerate = args[0].Trim() == "1" ? true : false;
            }

            //Global variants of folders.
            string templatefile = ConfigurationManager.AppSettings["templatefile"].ToString();
            string savefolder = ConfigurationManager.AppSettings["savefolder"].ToString();
            var saveExcelFolder = savefolder;

            //Create a log file.
            string batchNo = DateTime.Now.ToString("yyyyMMddhhmmss");
            batchNo = (string.IsNullOrEmpty(ConfigurationManager.AppSettings["BatchFlag"]) ? string.Empty : ConfigurationManager.AppSettings["BatchFlag"]) + batchNo;
            var logFile = "OfflineLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            StreamWriter sw = new StreamWriter(savefolder + logFile, true);
            sw.WriteLine("Started at " + DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToLongTimeString() + ":");
            sw.Close();

            try
            {
                //Make sure use 1033 as lang id.
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                //Style Mapping, call an outside service to load mapping information to database. NOT execute always, only execute when this function is enabled in app.config.
                StyleMapping(sw, savefolder, logFile);

                OfflineFileGenerator.App app = new OfflineFileGenerator.App();
                WriteToLog("Step 1:" + DateTime.Now.TimeOfDay);

                //Retrieve the code rows describing which spreadsheets to create
                SqlDataReader reader = app.GetFiles(batchNo);
                WriteToLog("Step 2:" + DateTime.Now.TimeOfDay);

                OfflineGenerator og = new OfflineGenerator();

                FileInfo newFile = new FileInfo(templatefile);
                if (newFile.Extension.ToLower() == ".xls")
                    app.IncludePricing = og.isPricingInclude(templatefile, "IncludePricing");

                var OfflineGenerateSuccessful = true;
                while (reader.Read())
                {
                    if (!string.IsNullOrEmpty(testFormName) && reader["OfflineOrderFormFileName"].ToString() != testFormName)
                        continue;

                    DateTime startTime = DateTime.Now;
                    try
                    {
                        WriteToLog("==============================================");

                        //CREATE DIRECTORY BY OfflineOrderFormDistribution
                        var OfflineOrderFormDistribution = "";
                        saveExcelFolder = savefolder;

                        reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'OfflineOrderFormDistribution'");
                        if (reader.GetSchemaTable().DefaultView.Count > 0)
                            OfflineOrderFormDistribution = reader["OfflineOrderFormDistribution"] == null ? "" : reader["OfflineOrderFormDistribution"].ToString();
                     
                        reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'Currency'");
                        if (reader.GetSchemaTable().DefaultView.Count > 0)
                            app.Currency = reader["Currency"] == null ? "" : reader["Currency"].ToString();
                        
                        if(!string.IsNullOrEmpty(OfflineOrderFormDistribution.Trim()))
                        {
                            var subFolder = OfflineOrderFormDistribution.Trim();
                            switch (subFolder.ToUpper())
                            {
                                case "ALL":
                                    break;
                                case "COUNTRY":
                                    subFolder = reader["Country"].ToString();
                                    break;
                                case "CATEGORY":
                                    subFolder = reader["EntityType"].ToString();
                                    break;
                                case "SOLDTO":
                                    subFolder = reader["SoldTo"].ToString();
                                    break;
                                case "CATALOGCATEGORY":
                                    subFolder = reader["CatalogCategory"].ToString();
                                    break;
                            }
                            saveExcelFolder = Path.Combine(savefolder, subFolder) + "\\";
                            if(!Directory.Exists(saveExcelFolder))
                                Directory.CreateDirectory(saveExcelFolder);

                            var OfflineOrderFormSubDistribution = string.Empty;
                            reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'OfflineOrderFormSubDistribution'");
                            if (reader.GetSchemaTable().DefaultView.Count > 0)
                                OfflineOrderFormSubDistribution = reader["OfflineOrderFormSubDistribution"] == null ? "" : reader["OfflineOrderFormSubDistribution"].ToString();
                            if (!string.IsNullOrEmpty(OfflineOrderFormSubDistribution))
                            {
                                var ssubFolder = OfflineOrderFormSubDistribution.Trim(); //CAN BE CatalogCategory OR OTHER VALUES
                                saveExcelFolder = Path.Combine(saveExcelFolder, ssubFolder) + "\\";
                                if (!Directory.Exists(saveExcelFolder))
                                    Directory.CreateDirectory(saveExcelFolder);
                            }
                        }
                        WriteToLog(OfflineOrderFormDistribution + " => " + saveExcelFolder);

                        if (CMSGenerate)
                        {
                            var cfileName = CleanFileName(reader["OfflineOrderFormFileName"].ToString());
                            var cfilePath = Path.Combine(saveExcelFolder, cfileName);
                            WriteToLog("CMSGenerate check file:" + cfilePath);
                            if (File.Exists(cfilePath))
                                continue;
                        }

                        var formType = string.Empty;
                        reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'FormType'";
                        if(reader.GetSchemaTable().DefaultView.Count > 0)
                            formType = reader["FormType"] == null ? string.Empty : reader["FormType"].ToString();
                        var catalog = reader["Catalog"].ToString();

                        if (ConfigurationManager.AppSettings[catalog] != null)
                            templatefile = ConfigurationManager.AppSettings[catalog].ToString();
                        //Log.
                        app.StyleCollection = new Hashtable();
                        app.LogOfflineOrderFormTemplate(batchNo, reader["CodeId"].ToString());
                        //Create the XML version of the file - see FormGenerator.cs
                        WriteToLog("Generate Template begin:" + DateTime.Now.TimeOfDay);//Generate

                        if (!string.IsNullOrEmpty(formType) && formType == ConfigurationManager.AppSettings["TabularFormType"])
                        {
                            if (string.IsNullOrEmpty(ConfigurationManager.AppSettings[catalog]))
                                templatefile = ConfigurationManager.AppSettings["TabularTemplateFile"];
                            switch (formType)
                            {
                                case "Category":
                                    var priceType = "W";
                                    var programcode = string.Empty;
                                    reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'PriceType'";
                                    if (reader.GetSchemaTable().DefaultView.Count > 0)
                                        priceType = reader["PriceType"] == null ? string.Empty : reader["PriceType"].ToString();
                                    reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'ProgramCode'";
                                    if (reader.GetSchemaTable().DefaultView.Count > 0)
                                        programcode = reader["ProgramCode"] == null ? string.Empty : reader["ProgramCode"].ToString();
                                    var cgenerator = new CategoryTabularFormGenerator();
                                    cgenerator.PriceType = priceType;
                                    cgenerator.ProgramCode = programcode;
                                    cgenerator.GenerateCategoryTabularOrderForm(templatefile, CleanFileName(reader["OfflineOrderFormFileName"].ToString()), reader["SoldTo"].ToString(), reader["Catalog"].ToString(), reader["PriceCode"].ToString(), saveExcelFolder);
                                    break;
                                default:
                                    var tgenerator = new TabularFormGenerator();
                                    tgenerator.GenerateTabularOrderForm(templatefile, CleanFileName(reader["OfflineOrderFormFileName"].ToString()), reader["SoldTo"].ToString(), reader["Catalog"].ToString(), reader["PriceCode"].ToString(), saveExcelFolder);
                                    break;
                            }
                        }
                        else
                        {
                            var priceType = string.Empty;
                            reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'PriceType'";
                            if (reader.GetSchemaTable().DefaultView.Count > 0)
                                priceType = reader["PriceType"] == null ? string.Empty : reader["PriceType"].ToString();
                            app.PriceType = priceType;

                            var deptID = string.Empty;
                            reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'TopDeptID'";
                            if (reader.GetSchemaTable().DefaultView.Count > 0)
                                deptID = reader["TopDeptID"] == null ? string.Empty : reader["TopDeptID"].ToString();
                            app.OfflineDeptID = deptID;

                            var ImageOffline = string.Empty;
                            var noneImageFilePath = string.Empty;
                            reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'ShortForm'");
                            if (reader.GetSchemaTable().DefaultView.Count > 0)
                                ImageOffline = reader["ShortForm"] == null ? "" : reader["ShortForm"].ToString();

                            if (ImageOffline == "Images")
                            {
                                var noneImageFileName = reader["NoneImageFileName"].ToString();
                                if (!string.IsNullOrEmpty(noneImageFileName))
                                    noneImageFilePath = Path.Combine(saveExcelFolder, CleanFileName(noneImageFileName));
                            }

                            bool newXLSMForm = false; // for Toms
                            reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'Skip'");
                            if (reader.GetSchemaTable().DefaultView.Count > 0)
                            {
                                newXLSMForm = true;
                                if (reader["Skip"] != null && reader["Skip"].ToString() == "1")
                                {
                                    WriteToLog("Skipped");
                                    continue;
                                }
                            }

                            if (ImageOffline != "Images" || !File.Exists(noneImageFilePath))
                            {
                                //continue;
                                reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'OrderShipDate'");
                                if (reader.GetSchemaTable().DefaultView.Count > 0)
                                {
                                    app.OrderShipDate = false;
                                    var OrderShipDate = reader["OrderShipDate"] == null ? "" : reader["OrderShipDate"].ToString();
                                    if (OrderShipDate == "1")
                                        app.OrderShipDate = true;                                       
                                }
                                
                                int maxcols = newXLSMForm ?
                                    app.GenerateStandardXLSMData(
                                        Path.Combine(saveExcelFolder, CleanFileName(reader["OfflineOrderFormFileName"].ToString())), 
                                        saveExcelFolder,
                                        reader) :
                                    app.GenerateData(
                                        Path.Combine(saveExcelFolder, "test.xls"), 
                                        reader["SoldTo"].ToString(), 
                                        reader["Catalog"].ToString(), 
                                        reader["PriceCode"].ToString(), 
                                        saveExcelFolder);

                                WriteToLog("Generate Template end:" + DateTime.Now.TimeOfDay);

                                if (!app.OfflineGenerateSuccessful)
                                    OfflineGenerateSuccessful = false;

                                if (maxcols > -1)
                                {
                                    WriteToLog("Generate Excel begin:" + DateTime.Now.TimeOfDay);
                                    og = new OfflineGenerator();
                                    og.OrderShipDate = app.OrderShipDate;
                                    og.CopyForm(Path.Combine(saveExcelFolder, "test.xls"), templatefile, Path.Combine(saveExcelFolder, CleanFileName(reader["OfflineOrderFormFileName"].ToString())), maxcols, reader);
                                    WriteToLog("Generate Excel end:" + DateTime.Now.TimeOfDay);
                                }
                            }

                            if (ImageOffline == "Images" && File.Exists(noneImageFilePath))
                            {
                                string patchPath = ConfigurationManager.AppSettings["ImagePatchPath"]??"";
                                if (File.Exists(patchPath)) // Obsolete!!!! Use EPPlus instead. first implement for BauerUS, to solve the issue that product image could not be shared in pre-created date sheets for PNOI 1.2.4.
                                {
                                    Process p = new Process();
                                    p.StartInfo.FileName = "cmd.exe";
                                    p.StartInfo.UseShellExecute = false;
                                    p.StartInfo.RedirectStandardInput = true;
                                    p.StartInfo.RedirectStandardOutput = true;
                                    p.StartInfo.CreateNoWindow = true;
                                    p.Start();
                                    string Cmdstring = string.Format("\"{0}\" \"{1}\" \"{2}\"", patchPath, noneImageFilePath, Path.Combine(saveExcelFolder, CleanFileName(reader["OfflineOrderFormFileName"].ToString())));
                                    p.StandardInput.WriteLine(Cmdstring);
                                    p.StandardInput.WriteLine("exit");
                                    string s = p.StandardOutput.ReadToEnd();
                                    p.Close();
                                    Console.WriteLine(s);
                                }
                                else
                                {
                                    WriteToLog("Generate Image Excel gegin:" + DateTime.Now.TimeOfDay);
                                    app.GenerateImageOfflineOrderForm(noneImageFilePath, Path.Combine(saveExcelFolder, CleanFileName(reader["OfflineOrderFormFileName"].ToString())));
                                    WriteToLog("Generate Image Excel end:" + DateTime.Now.TimeOfDay);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sw = new StreamWriter(savefolder + logFile, true);
                        sw.WriteLine("At " + DateTime.Now.Date + ", " + DateTime.Now.TimeOfDay + ":");
                        sw.WriteLine(ex.Message);
                        sw.WriteLine(ex.StackTrace);
                        if (ex.InnerException != null)
                            sw.WriteLine(ex.InnerException.Message);
                        sw.WriteLine();
                        sw.Close();

                        //Go on with next form.
                        reader = app.GetFiles(batchNo);
                    }
                    finally
                    {
                        WriteToLog(string.Format("{0} Total time: {1}minute(s) {2}second(s)", CleanFileName(reader["OfflineOrderFormFileName"].ToString()), (DateTime.Now - startTime).Minutes, (DateTime.Now - startTime).Seconds));
                    }
                }

                if(OfflineGenerateSuccessful)
                    System.Environment.ExitCode = 100;
            }
            catch(Exception e)
            {
                sw = new StreamWriter(savefolder + logFile, true);
                sw.WriteLine("At " + DateTime.Now.Date + ", " + DateTime.Now.TimeOfDay + ":");
                sw.WriteLine(e.Message);
                sw.WriteLine(e.StackTrace);
                if(e.InnerException!=null)
                    sw.WriteLine(e.InnerException.Message);
                sw.WriteLine();
                sw.Close();
            }
            finally {
                sw = new StreamWriter(savefolder + logFile, true);
                sw.WriteLine("Finished at " + DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToLongTimeString() + ":");
                sw.Close();
            }
        }

        public static string CleanFileName(string fileName)
        {
            return Path.GetInvalidFileNameChars().Aggregate(fileName, (current, c) => current.Replace(c.ToString(), ""));
        }

        protected static void StyleMapping(StreamWriter sw, string savefolder, string logFile)
        {
            try
            {
                var styleMapping = string.IsNullOrEmpty(ConfigurationManager.AppSettings["StyleMapping"]) ? "0" : ConfigurationManager.AppSettings["StyleMapping"];
                var styleMappingService = ConfigurationManager.AppSettings["StyleMappingService"];

                if (styleMapping == "1" && !string.IsNullOrEmpty(styleMappingService))
                {
                    OfflineFileGenerator.StyleMapping.StyleMapping mappingService = new OfflineFileGenerator.StyleMapping.StyleMapping();
                    mappingService.Url = styleMappingService;
                    mappingService.LoadMapping();
                }
            }
            catch (Exception e)
            {
                sw = new StreamWriter(savefolder + logFile, true);
                sw.WriteLine("At " + DateTime.Now.Date + ", " + DateTime.Now.TimeOfDay + ":");
                sw.WriteLine(e.Message);
                sw.WriteLine(e.StackTrace);
                if (e.InnerException != null)
                    sw.WriteLine(e.InnerException.Message);
                sw.WriteLine();
                sw.Close();
            }
        }

    }

    public static class MiscUtility
    {
        public static Dictionary<string, string> ToDic(this string list, char[] splitChars)
        {
            string[] columnListArr = list.Split(splitChars, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, string> columnList = new Dictionary<string, string>();
            for (int j = 0; j < columnListArr.Length; j = j + 2)
            {
                columnList.Add(columnListArr[j], columnListArr[j + 1]);
            }
            return columnList;
        }
    }
}
