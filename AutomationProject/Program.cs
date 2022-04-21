using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.OleDb;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using ExcelDataReader;

namespace AutomationProject
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string poolMappingFile = Convert.ToString(args[0]);
                string excelFile = Convert.ToString(args[1]);
                string tsvFile = Convert.ToString(args[2]);

                DataTable dtResults = new DataTable();
                
                // Read the Excel file first
                dtResults = ReadExcel(excelFile);
                
                // Read the TSV file next
                dtResults.Merge(ReadTSV(tsvFile, poolMappingFile));


                // Sort results in ascending order and and order by language code
                DataView dvResults = new DataView(dtResults);
                dvResults.Sort = "LanguageCode, Retailers ASC";
                dtResults = dvResults.ToTable();

                DataColumn dataColumn = new DataColumn("Processed");
                dataColumn.DefaultValue = "N";
                dtResults.Columns.Add(dataColumn);

                List<JsonOutput> outputList = new List<JsonOutput>();

                for (int i = 0; i < dtResults.Rows.Count; i++)
                {
                    string retailer = dtResults.Rows[i][2].ToString();
                    string isProcessed = dtResults.Rows[i][4].ToString();
                    if (!String.IsNullOrEmpty(retailer) && isProcessed.Equals("N"))
                    {
                        DataRow[] rows = dtResults.Select($"Retailers = '{retailer}'");
                        JsonOutput jsonOutput = new JsonOutput();
                        jsonOutput.Domain = rows[0]["Retailers"].ToString();
                        jsonOutput.Market = rows[0]["LanguageCode"].ToString();
                        if (rows.Length > 1)
                        {
                            for (int n = 0; n < rows.Length; n++)
                            {
                                // if the codes are delimited using the "<SEP>" delimiter
                                if (rows[n]["Code"].ToString().Contains("<SEP>"))
                                {
                                    string[] stringSeparators = new string[] { "<SEP>" };
                                    string[] couponCodes = rows[n]["Code"].ToString().Split(stringSeparators, StringSplitOptions.None);
                                    string[] descriptions = rows[n]["Description"].ToString().Split(stringSeparators, StringSplitOptions.None);
                                    for (int j = 0; j < couponCodes.Length; j++)
                                    {
                                        try
                                        {
                                            jsonOutput.Coupons.Add(new Coupon { CouponCode = couponCodes[j], CouponDescription = descriptions[j] });
                                        }
                                        catch (Exception)
                                        {
                                            jsonOutput.Coupons.Add(new Coupon { CouponCode = couponCodes[j], CouponDescription = "" });
                                        }

                                    }
                                }
                                else
                                {
                                    jsonOutput.Coupons.Add(new Coupon { CouponCode = rows[n]["Code"].ToString(), CouponDescription = rows[n]["Description"].ToString() });
                                }
                                rows[n]["Processed"] = "Y";
                            }
                        }
                        else
                        {
                            foreach (DataRow row in rows)
                            {
                                jsonOutput.Coupons.Add(new Coupon { CouponCode = row["Code"].ToString(), CouponDescription = row["Description"].ToString() });
                                row["Processed"] = "Y";
                            }
                        }
                        outputList.Add(jsonOutput);
                    }
                }

                string jsonString = JsonConvert.SerializeObject(outputList);
                File.WriteAllText(@"JsonOutput.json", jsonString);
            }
            catch (Exception)
            {
                throw;
            }
        }

        // Method to read data from the TSV file
        public static DataTable ReadTSV(string tsvFile, string poolMappingFile)
        {
            int count = 0;
            DataTable dt = new DataTable();
            if (File.Exists(tsvFile))
            {
                string[] filename = tsvFile.Split('_');
                string LanguageCode = string.Empty;
                int poolId = Convert.ToInt32(filename[filename.Length - 1].Split('.')[0]);
                FileStream fs = new FileStream(poolMappingFile, FileMode.Open, FileAccess.Read);
                Dictionary<int, string> poolMapping = new Dictionary<int, string>();
                using (StreamReader sr = new StreamReader(fs))
                {
                    string mappingFileContent = string.Empty;
                    while ((mappingFileContent = sr.ReadLine()) != null)
                    {
                        string[] mappingString = mappingFileContent.Split('\t');
                        if (poolId == Convert.ToInt32(mappingString[0]))
                        {
                            LanguageCode = mappingString[1];
                            break;
                        }
                    }
                }

                using(TextReader tr = File.OpenText(tsvFile))
                {
                    string line;
                    while((line= tr.ReadLine()) != null)
                    {
                        string[] items = line.Split('\t');
                        if (dt.Columns.Count == 0 && count == 0)
                        {
                            // Create the data columns for the data table based on the number of items
                            // on the first line of the file
                            for (int i = 0; i < items.Length; i++)
                            {
                                if (i == 1)
                                {
                                    dt.Columns.Add(new DataColumn("Description", typeof(string)));
                                }
                                if (i == 2)
                                {
                                    dt.Columns.Add(new DataColumn("Code", typeof(string)));
                                }
                                if (i == 6)
                                {
                                    dt.Columns.Add(new DataColumn("Retailers", typeof(string)));
                                }
                            }
                            DataColumn dataColumn = new DataColumn("LanguageCode");
                            dataColumn.DefaultValue = LanguageCode;
                            dt.Columns.Add(dataColumn);
                        }
                        else
                        {
                            int columnsLength = line.Split('\t').Length;
                            dt.Rows.Add();
                            if (columnsLength < 18)
                            {
                                string completeLine = line + " " + tr.ReadLine();
                                // handle the line break as seen in the input file where the description breaks into the next line
                                if (!String.IsNullOrEmpty(completeLine.Split('\t')[2]))
                                {
                                    dt.Rows[count][0] = completeLine.Split('\t')[1];    // Description
                                    dt.Rows[count][1] = completeLine.Split('\t')[2];    // Code
                                    if (completeLine.Split('\t')[6].StartsWith("https"))
                                    {
                                        dt.Rows[count][2] = completeLine.Split('\t')[6].Substring("https://www.".Length, completeLine.Split('\t')[6].Length - "https://www.".Length);    // Retailer
                                    }
                                    else
                                    {
                                        dt.Rows[count][2] = completeLine.Split('\t')[6];    // Retailer
                                    }

                                }
                            }
                            else
                            {
                                if (!String.IsNullOrEmpty(line.Split('\t')[2]))
                                {
                                    dt.Rows[count][0] = line.Split('\t')[1];    // Description
                                    dt.Rows[count][1] = line.Split('\t')[2];    // Code
                                    if (line.Split('\t')[6].StartsWith("https"))
                                    {
                                        dt.Rows[count][2] = line.Split('\t')[6].Substring("https://www.".Length, line.Split('\t')[6].Length - "https://www.".Length);    // Retailer
                                    }
                                    else
                                    {
                                        dt.Rows[count][2] = line.Split('\t')[6];    // Retailer
                                    }
                                }
                            }
                            count++;
                        }
                    }

                        DataView dv = new DataView(dt);
                        DataTable dt2 = dv.ToTable(true, "Description", "Code", "Retailers");
                        dt = dt2.DefaultView.ToTable();
                }
            }
            return dt;
        }

        // Method to read the data from the Excel file
        public static DataTable ReadExcel(string excelFile)
        {
            Excel.Application _Excel = new Excel.Application();

            string filepath = excelFile;

            Excel.Workbook workBook = _Excel.Workbooks.Open(filepath);
                       
            String[] excelSheets = new String[workBook.Worksheets.Count];

            DataTable dtMarkets = new DataTable();
            foreach (Excel.Worksheet worksheet in workBook.Worksheets)
            {
                string worksheetName = worksheet.Name;
                int totalRows = worksheet.UsedRange.Rows.Count;

                DataTable dt = new DataTable();
                dt.Columns.Add(new DataColumn("Description", typeof(string)));
                dt.Columns.Add(new DataColumn("Code", typeof(string)));
                dt.Columns.Add(new DataColumn("Retailers", typeof(string)));
                DataColumn dataColumn = new DataColumn("LanguageCode");
                dataColumn.DefaultValue = worksheetName;
                dt.Columns.Add(dataColumn);
                int count = 0;
                // start adding records from the worksheet to the Data Table
                for (int i = 2; i < totalRows; i++)
                {
                    if (!String.IsNullOrEmpty(((Excel.Range)worksheet.Cells[i, 2]).Value) && !String.IsNullOrEmpty(((Excel.Range)worksheet.Cells[i, 3]).Value))
                    {
                        dt.Rows.Add();
                        string retailer = ((Excel.Range)worksheet.Cells[i, 2]).Value;
                        string code = ((Excel.Range)worksheet.Cells[i, 3]).Value;
                        string description = ((Excel.Range)worksheet.Cells[i, 4]).Value;
                        dt.Rows[count][0] = description;
                        dt.Rows[count][1] = code;
                        dt.Rows[count][2] = retailer;
                        count++;
                    }
                }

                if (dtMarkets.Rows.Count == 0)
                {
                    dtMarkets = dt;
                }
                else
                {
                    dtMarkets.Merge(dt);
                }
            }

            DataView dv = new DataView(dtMarkets);
            dtMarkets = dv.ToTable(true, "Description", "Code", "Retailers", "LanguageCode");
            return dtMarkets;
        }
    }
}
