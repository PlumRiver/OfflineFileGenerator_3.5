using System;
using System.IO;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        }

        static void Main(string[] args)
        {
            //var app1 = new App();
            //app1.GenerateImageOfflineOrderForm(@"D:\PRWork\Plumriver-Customers\Toms\offline\Eyewear Collection_US.xls", @"D:\PRWork\Plumriver-Customers\Toms\offline\Eyewear Collection_US_Image.xls");
            //app1.InserImagesToOffline();


            string templatefile = ConfigurationManager.AppSettings["templatefile"].ToString();
            string savefolder = ConfigurationManager.AppSettings["savefolder"].ToString();
            var saveExcelFolder = savefolder;
            string batchNo = DateTime.Now.ToString("yyyyMMddhhmmss");
            var logFile = "OfflineLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            StreamWriter sw = new StreamWriter(savefolder + logFile, true);
            sw.WriteLine("Started at " + DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToLongTimeString() + ":");
            sw.Close();

            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                //Style Mapping
                StyleMapping(sw, savefolder, logFile);

                OfflineFileGenerator.App app = new OfflineFileGenerator.App();
                WriteToLog("Step 1:" + DateTime.Now.TimeOfDay);

                //Retrieve the code rows describing which spreadsheets to create
                SqlDataReader reader = app.GetFiles(batchNo);
                WriteToLog("Step 2:" + DateTime.Now.TimeOfDay);

                OfflineGenerator og = new OfflineGenerator();

                app.IncludePricing =  og.isPricingInclude(templatefile, "IncludePricing");

                while (reader.Read())
                {
                    try
                    {
                        //CREATE DIRECTORY BY OfflineOrderFormDistribution
                        var OfflineOrderFormDistribution = "";
                        saveExcelFolder = savefolder;

                        reader.GetSchemaTable().DefaultView.RowFilter = string.Format("ColumnName= 'OfflineOrderFormDistribution'");
                        if (reader.GetSchemaTable().DefaultView.Count > 0)
                            OfflineOrderFormDistribution = reader["OfflineOrderFormDistribution"] == null ? "" : reader["OfflineOrderFormDistribution"].ToString();
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
                            }
                            saveExcelFolder = Path.Combine(savefolder, subFolder) + "\\";
                            if(!Directory.Exists(saveExcelFolder))
                                Directory.CreateDirectory(saveExcelFolder);
                        }
                        WriteToLog(OfflineOrderFormDistribution);
                        WriteToLog(saveExcelFolder);

                        var formType = string.Empty;
                        reader.GetSchemaTable().DefaultView.RowFilter = "ColumnName= 'FormType'";
                        if(reader.GetSchemaTable().DefaultView.Count > 0)
                            formType = reader["FormType"] == null ? string.Empty : reader["FormType"].ToString();
                        var catalog = reader["Catalog"].ToString();

                        if (ConfigurationManager.AppSettings[catalog] != null)
                            templatefile = ConfigurationManager.AppSettings[catalog].ToString();
                        WriteToLog(templatefile);
                        //Log.
                        app.StyleCollection = new Hashtable();
                        app.LogOfflineOrderFormTemplate(batchNo, reader["CodeId"].ToString());
                        //Create the XML version of the file - see FormGenerator.cs
                        WriteToLog("Generate Template begin:" + DateTime.Now.TimeOfDay);//Generate

                        if (!string.IsNullOrEmpty(formType) && formType == ConfigurationManager.AppSettings["TabularFormType"])
                        {
                            if (string.IsNullOrEmpty(ConfigurationManager.AppSettings[catalog]))
                                templatefile = ConfigurationManager.AppSettings["TabularTemplateFile"];
                            var generator = new TabularFormGenerator();
                            generator.GenerateTabularOrderForm(templatefile, CleanFileName(reader["OfflineOrderFormFileName"].ToString()), reader["SoldTo"].ToString(), reader["Catalog"].ToString(), reader["PriceCode"].ToString(), saveExcelFolder);
                        }
                        else
                        {
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
                            if (ImageOffline != "Images" || !File.Exists(noneImageFilePath))
                            {
                                //continue;
                                int maxcols = app.GenerateData(Path.Combine(saveExcelFolder, "test.xls"), reader["SoldTo"].ToString(), reader["Catalog"].ToString(), reader["PriceCode"].ToString(), saveExcelFolder);
                                WriteToLog("Generate Template end:" + DateTime.Now.TimeOfDay);
                                if (maxcols > -1)
                                {
                                    WriteToLog("Generate Excel begin:" + DateTime.Now.TimeOfDay);
                                    og = new OfflineGenerator();
                                    og.CopyForm(Path.Combine(saveExcelFolder, "test.xls"), templatefile, Path.Combine(saveExcelFolder, CleanFileName(reader["OfflineOrderFormFileName"].ToString())), maxcols, reader);
                                    WriteToLog("Generate Excel begin:" + DateTime.Now.TimeOfDay);
                                }
                            }
                            if (ImageOffline == "Images" && File.Exists(noneImageFilePath))
                            {
                                app.GenerateImageOfflineOrderForm(noneImageFilePath, Path.Combine(saveExcelFolder, CleanFileName(reader["OfflineOrderFormFileName"].ToString())));
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
                    }
                }
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
}
