using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using System.Net;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPA_Automation
{
    class Program
    {
        private static readonly HttpClient client = new HttpClient();
        private static List<FacilityDetails> APB_Report_Basin;
        private static List<FacilityDetails> STX_Report_Basin;
        private static List<FacilityDetails> WIB_Report_Basin;

        private static List<FacilityDetails> APB_Report_BasinGeo;
        private static List<FacilityDetails> STX_Report_BasinGeo;
        private static List<FacilityDetails> WIB_Report_BasinGeo;
        

        class APICallResult
        {
            public string data { get; set; }
            public string message { get; set; }
        }

        class FacilityDetails
        {
            public string CarbonDioxide { get; set; }
            public string Methane { get; set; }
            public string GasSales { get; set; }
            public string OilSales { get; set; }
            public string ParentCompanyLegalName { get; set; }
            public string QuantityGasReceived { get; set; }
            public string QuantityOfHydrocarbonLiquidsReceived { get; set; }

        }

        public class Col
        {
            public string id { get; set; }
            public string name { get; set; }
            public string field { get; set; }
            public bool sortable { get; set; }
            public string type { get; set; }
            public string cssClass { get; set; }
        }

        public class Row
        {
            public string id { get; set; }
            public string icons { get; set; }
            public string facility { get; set; }
            public string city { get; set; }
            public string state { get; set; }
            public string total { get; set; }
            public string sectors { get; set; }
        }

        public class Data
        {
            public List<Col> cols { get; set; }
            public List<Row> rows { get; set; }
        }

        public class RootObject
        {
            public Data data { get; set; }
            public string tableHeader { get; set; }
            public string domain { get; set; }
            public string mode { get; set; }
            public string unit { get; set; }
            public int year { get; set; }
        }

        public class Col_BasinGeo
        {
            public string id { get; set; }
            public string name { get; set; }
            public string field { get; set; }
            public bool sortable { get; set; }
            public string type { get; set; }
            public string cssClass { get; set; }
        }

        public class Row_BasinGeo
        {
            public string id { get; set; }
            public string facility { get; set; }
            public string petroleum { get; set; }
        }

        public class Data_BasinGeo
        {
            public List<Col_BasinGeo> Col_BasinGeos { get; set; }
            public List<Row_BasinGeo> Row_BasinGeos { get; set; }
        }

        public class RootObject_BasinGeo
        {
            public Data_BasinGeo data { get; set; }
            public string tableHeader { get; set; }
            public string domain { get; set; }
            public string mode { get; set; }
            public string unit { get; set; }
            public int year { get; set; }
        }

        private static async Task RunAsync()
        {
            //var values = new Dictionary<string, string>
            //        {
            //            { "basin", "395" },
            //            { "countyFips", "" },
            //            { "currentYear", "2017" },
            //            { "dataSource", "0" },
            //            { "emissionsType", "" },
            //            { "gases", [true, true, false, false, false, false, false, false, false, false, false, false]},
            //            { "highE", "23000000" },
            //            { "injectionSelection", 11},
            //            { "lowE", "-20000" },
            //            { "msaCode", "" },
            //            { "overlayLevel", "hello" },
            //            { "query", "world" },
            //            { "reportingStatus", "hello" },
            //            { "reportingYear", "world" },
            //            { "searchOptions", "hello" },
            //            { "sectors", "world" },
            //            { "sortOrder", "hello" },
            //            { "state", "world" },
            //            { "stateLevel", "hello" },
            //            { "supplierSector", "world" },
            //            { "trend", "hello" },
            //            { "tribalLandId", "world" }
            //        };

            //var content = new FormUrlEncodedContent(new[]
            //{
            //    new KeyValuePair<string, string>("basin", "395"),
            //    new KeyValuePair<string, string>("countyFips", ""),
            //    new KeyValuePair<string, string>("currentYear", "2017"),
            //    new KeyValuePair<string, string>("dataSource", ""),
            //    new KeyValuePair<string, string>("emissionsType", ""),
            //    new KeyValuePair<string, string>("gases", "login"),
            //    new KeyValuePair<string, string>("highE", "login"),
            //    new KeyValuePair<string, string>("injectionSelection", "login"),
            //    new KeyValuePair<string, string>("lowE", "login"),
            //    new KeyValuePair<string, string>("msaCode", "login"),
            //    new KeyValuePair<string, string>("overlayLevel", "login"),
            //    new KeyValuePair<string, string>("query", "login"),
            //    new KeyValuePair<string, string>("reportingStatus", "login"),
            //    new KeyValuePair<string, string>("reportingYear", "login"),
            //    new KeyValuePair<string, string>("searchOptions", "login"),
            //    new KeyValuePair<string, string>("sectors", "login"),
            //    new KeyValuePair<string, string>("sortOrder", "login"),
            //    new KeyValuePair<string, string>("state", "login"),
            //    new KeyValuePair<string, string>("stateLevel", "login"),
            //    new KeyValuePair<string, string>("supplierSector", "login"),
            //    new KeyValuePair<string, string>("trend", "login"),
            //    new KeyValuePair<string, string>("tribalLandId", "login")
            //});

            var myData = new {
                basin = "395",
                countyFips = "",
                currentYear = "2017",
                dataSource = "0",
                emissionsType = "",
                gases = new[] { true, true, false, false, false, false, false, false, false, false, false, false },
                highE = "23000000",
                injectionSelection = 11,
                lowE = "-20000",
                msaCode = "",
                overlayLevel = 0,
                query = "",
                reportingStatus = "ALL",
                reportingYear = "2017",
                searchOptions = "11001000",
                sectors = new[] {
                    new[] {false},
                    new[] {false, false, false, false, false},
                    new[] {false, false, false, false, false, false, false, false},
                    new[] {false, false, false, false, false, false},
                    new[] {false},
                    new[] {false, false, false},
                    new[] {false, false, false, false, false, false, false, false, false, false, false, false},
                    new[] {false, false, false, false, false, false, false, false, false, false, false},
                    new[] {true, false, true, false, false, false, false, false, false, false, false, false}
                },
                sortOrder = "0",
                state = "",
                stateLevel = "0",
                supplierSector = 0,
                trend = "current",
                tribalLandId = ""
            };

            var jsonObject = JsonConvert.SerializeObject(myData);
            var stringContent = new StringContent(jsonObject.ToString(), System.Text.Encoding.UTF8, "application/json");
            //client.Timeout = new TimeSpan(0, 5, 0);
            try
            {
                HttpResponseMessage response = await client.PostAsync("https://ghgdata.epa.gov/ghgp/service/listFacilityForBasin/", stringContent);
                response.EnsureSuccessStatusCode();

                string data = await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
            Console.WriteLine("Testing Done!");

        }

        static void GetRequest()
        {
            try
            {
                string html = string.Empty;
                string url = @"https://ghgdata.epa.gov/ghgp/service/export?q=&tr=current&ds=O&ryr=2017&cyr=2017&lowE=-20000&highE=23000000&st=&fc=&mc=&rs=ALL&sc=0&is=11&et=&tl=&pn=undefined&ol=0&sl=0&bs=160A&g1=1&g2=1&g3=0&g4=0&g5=0&g6=0&g7=0&g8=0&g9=0&g10=0&g11=0&g12=0&s1=0&s2=0&s3=0&s4=0&s5=0&s6=0&s7=0&s8=0&s9=1&s10=0&s201=0&s202=0&s203=0&s204=0&s301=0&s302=0&s303=0&s304=0&s305=0&s306=0&s307=0&s401=0&s402=0&s403=0&s404=0&s405=0&s601=0&s602=0&s701=0&s702=0&s703=0&s704=0&s705=0&s706=0&s707=0&s708=0&s709=0&s710=0&s711=0&s801=0&s802=0&s803=0&s804=0&s805=0&s806=0&s807=0&s808=0&s809=0&s810=0&s901=0&s902=1&s903=0&s904=0&s905=0&s906=0&s907=0&s908=0&s909=0&s910=0&s911=0&sf=11001000&listExport=false"; ;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                request.Method = "GET";
                request.Headers["cache-control"] = "no-store, no-cache, must-revalidate";
                //request.Headers["connection"] = "Keep-Alive";
                //request.Connection = "Keep-Alive";
                request.Headers["content-disposition"] = "attachment; filename=flight.xls";
                //request.Headers["content-type"] = "application/ms-excel;charset=UTF-8";
                request.ContentType = "application/ms-excel;charset=UTF-8";
                request.Headers["keep-alive"] = "timeout=5, max=100";
                request.Headers["pragma"] = "no-cache";
                request.Headers["server"] = "Apache";
                request.Headers["strict-transport-security"] = "max-age=31536000; includeSubDomainsâ";
                //request.Headers["transfer-encoding"] = "chunked";
                //request.SendChunked = true;
                request.Headers["x-frame-options"] = "SAMEORIGIN";
                request.Host = "ghgdata.epa.gov";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36";
                request.Accept = "*/*";
                request.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-us,en;q=0.5");




                request.AutomaticDecompression = DecompressionMethods.GZip;

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream))
                {
                    html = reader.ReadToEnd();
                }

                Console.WriteLine(html);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void doTesting()
        {
            //string strGHGRPID = "1009238";

            //Task<string> callDetailTask = Task.Run(() => GetFacilityDetails(strGHGRPID, "2017"));
            //callDetailTask.Wait();
            //string strDetails = callDetailTask.Result;

            //FacilityDetails fd = extractFacilityDetails(strDetails);

            using (ExcelPackage excel = new ExcelPackage())
            {
                var worksheet = excel.Workbook.Worksheets.Add("Worksheet1");
                worksheet.Cells["B1:E1"].Merge = true;
                worksheet.Cells["F1:I1"].Merge = true;

                // Popular header row data
                worksheet.Cells["A1"].Value = "Bakken";
                worksheet.Cells["B1"].Value = "Onshore Production";
                worksheet.Cells["F1"].Value = "Onshore Gathering and Boosting";

                List<string[]> headerRow = new List<string[]>()
                {
                  new string[] { "Parent Company", "CO2 Tonnes", "CH4 Tonnes", "Gas Sales (mcf)", "Oil Sales (bbls)", "CO2 Tonnes", "CH4 Tonnes", "Gas Sales (mcf)", "Oil Sales (bbls)" }
                };
                string headerRange = "A2:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "2";
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                FileInfo excelFile = new FileInfo(@"test.xlsx");
                excel.SaveAs(excelFile);
            }

            Console.WriteLine("Done Testing!");
        }

        /// <summary>
        /// Extract facility details from XML
        /// </summary>
        static FacilityDetails extractFacilityDetails(string strXML)
        {
            string strCarbonDioxide = "";
            string strMethane = "";
            string strGasSales = "";
            string strOilSales = "";
            string strParentCompanyLegalName = "";
            string strQuantityGasReceived = "";
            string strQuantityOfHydrocarbonLiquidsReceived = "";

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(strXML);

            XmlNodeList nodes = xml.GetElementsByTagName("GHGasInfoDetails");
            foreach (XmlNode xn in nodes)
            {
                if (xn["GHGasName"] != null)
                {
                    switch (xn["GHGasName"].InnerText.ToLower())
                    {
                        case "carbon dioxide":
                            strCarbonDioxide = xn["GHGasQuantity"]["CalculatedValue"].InnerText;
                            break;
                        case "methane":
                            strMethane = xn["GHGasQuantity"]["CalculatedValue"].InnerText;
                            break;
                        default:

                            break;
                    }
                }


            }

            nodes = xml.GetElementsByTagName("GasProducedCalendarYearForSales");
            if (nodes.Count>0)
            {
                strGasSales = nodes[0].InnerText;
            }
            

            nodes = xml.GetElementsByTagName("OilProducedCalendarYearForSales");
            if (nodes.Count > 0)
            {
                strOilSales = nodes[0].InnerText;
            }
            

            nodes = xml.GetElementsByTagName("ParentCompanyLegalName");
            if (nodes.Count > 0)
            {
                strParentCompanyLegalName = nodes[0].InnerText;
            }

            nodes = xml.GetElementsByTagName("QuantityGasReceived");
            if (nodes.Count > 0)
            {
                strQuantityGasReceived = nodes[0].InnerText;
            }

            nodes = xml.GetElementsByTagName("QuantityOfHydrocarbonLiquidsReceived");
            if (nodes.Count > 0)
            {
                strQuantityOfHydrocarbonLiquidsReceived = nodes[0].InnerText;
            }


            return new FacilityDetails
            {
                CarbonDioxide = strCarbonDioxide,
                Methane = strMethane,
                GasSales = strGasSales,
                OilSales = strOilSales,
                ParentCompanyLegalName = strParentCompanyLegalName,
                QuantityGasReceived = strQuantityGasReceived,
                QuantityOfHydrocarbonLiquidsReceived = strQuantityOfHydrocarbonLiquidsReceived
            };
        }

        private static async Task GetAsync(string uri)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);

            request.Headers["cache-control"] = "no-store, no-cache, must-revalidate";
            request.Headers["connection"] = "Keep-Alive";
            request.Headers["content-disposition"] = "attachment; filename=flight.xls";
            request.Headers["content-type"] = "application/ms-excel;charset=UTF-8";
            request.Headers["keep-alive"] = "timeout=5, max=100";
            request.Headers["pragma"] = "no-cache";
            request.Headers["server"] = "Apache";
            request.Headers["strict-transport-security"] = "max-age=31536000; includeSubDomainsâ";
            request.Headers["transfer-encoding"] = "chunked";
            request.Headers["x-frame-options"] = "SAMEORIGIN";



            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

            using (HttpWebResponse response = (HttpWebResponse)await request.GetResponseAsync())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                var xx= await reader.ReadToEndAsync();
            }
        }


        /// <summary>
        /// Request onshore gathering and boosting facility list for basin 
        /// </summary>
        private static async Task<APICallResult> GetBasinIDList(string strBasinID,string strYear)
        {
            string strError = "";
            string strReturnData = "";
            var myData = new
            {
                basin = strBasinID,
                countyFips = "",
                currentYear = strYear,
                dataSource = "0",
                emissionsType = "",
                gases = new[] { true, true, false, false, false, false, false, false, false, false, false, false },
                highE = "23000000",
                injectionSelection = 11,
                lowE = "-20000",
                msaCode = "",
                overlayLevel = 0,
                query = "",
                reportingStatus = "ALL",
                reportingYear = strYear,
                searchOptions = "11001000",
                sectors = new[] {
                    new[] {false},
                    new[] {false, false, false, false, false},
                    new[] {false, false, false, false, false, false, false, false},
                    new[] {false, false, false, false, false, false},
                    new[] {false},
                    new[] {false, false, false},
                    new[] {false, false, false, false, false, false, false, false, false, false, false, false},
                    new[] {false, false, false, false, false, false, false, false, false, false, false},
                    new[] {true, false, true, false, false, false, false, false, false, false, false, false}
                },
                sortOrder = "0",
                state = "",
                stateLevel = "0",
                supplierSector = 0,
                trend = "current",
                tribalLandId = ""
            };

            var jsonObject = JsonConvert.SerializeObject(myData);
            var stringContent = new StringContent(jsonObject.ToString(), System.Text.Encoding.UTF8, "application/json");
            //client.Timeout = new TimeSpan(0, 5, 0);
            try
            {
                HttpResponseMessage response = await client.PostAsync("https://ghgdata.epa.gov/ghgp/service/listFacilityForBasin/", stringContent);
                response.EnsureSuccessStatusCode();

                strReturnData = await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                strError = ex.Message;
                //Console.WriteLine(ex.Message);
            }

            return new APICallResult { data = strReturnData, message = strError };


        }

        /// <summary>
        /// Request onshore gathering and boosting facility list for basin geo 
        /// </summary>
        private static async Task<APICallResult> GetBasinGeoIDList(string strBasinID, string strYear)
        {
            string strError = "";
            string strReturnData = "";
            var myData = new
            {
                basin = strBasinID,
                countyFips = "",
                currentYear = strYear,
                dataSource = "B",
                emissionsType = "",
                gases = new[] { true, true, false, false, false, false, false, false, false, false, true, true },
                highE = "23000000",
                injectionSelection = 11,
                lowE = "-20000",
                msaCode = "",
                overlayLevel = 0,
                query = "",
                reportingStatus = "ALL",
                reportingYear = strYear,
                searchOptions = "11001000",
                sectors = new[] {
                    new[] {false},
                    new[] {false, false, false, false, false},
                    new[] {false, false, false, false, false, false, false, false},
                    new[] {false, false, false, false, false, false},
                    new[] {false},
                    new[] {false, false, false},
                    new[] {false, false, false, false, false, false, false, false, false, false, false, false},
                    new[] {false, false, false, false, false, false, false, false, false, false, false},
                    new[] { true, false, false, false, false, false, false, false, false, false, true, false }
                },
                sortOrder = "0",
                state = "",
                stateLevel = "0",
                supplierSector = 0,
                trend = "current",
                tribalLandId = ""
            };

            var jsonObject = JsonConvert.SerializeObject(myData);
            var stringContent = new StringContent(jsonObject.ToString(), System.Text.Encoding.UTF8, "application/json");
            //client.Timeout = new TimeSpan(0, 5, 0);
            try
            {
                HttpResponseMessage response = await client.PostAsync("https://ghgdata.epa.gov/ghgp/service/listFacilityForBasinGeo/", stringContent);
                response.EnsureSuccessStatusCode();

                strReturnData = await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                strError = ex.Message;
                //Console.WriteLine(ex.Message);
            }

            return new APICallResult { data = strReturnData, message = strError };


        }

        /// <summary>
        /// Request onshore production facility list 
        /// </summary>
        private static async Task<APICallResult> GetFacilityList(string strBasinID, string strYear)
        {
            string strError = "";
            string strReturnData = "";
            var myData = new
            {
                basin = strBasinID,
                countyFips = "",
                currentYear = strYear,
                dataSource = "B",
                emissionsType = "",
                gases = new[] { true, true, true, true, true, false, true, true, true, true, true, true },
                highE = "23000000",
                injectionSelection = 11,
                lowE = "-20000",
                msaCode = "",
                overlayLevel = 0,
                pageNumber = "",
                query = "",
                reportingStatus = "ALL",
                reportingYear = strYear,
                searchOptions = "11001000",
                sectors = new[] {
                    new[] {false},
                    new[] {false, false, false, false, false},
                    new[] {false, false, false, false, false, false, false, false},
                    new[] {false, false, false, false, false, false},
                    new[] {false},
                    new[] {false, false, false},
                    new[] {false, false, false, false, false, false, false, false, false, false, false, false},
                    new[] {false, false, false, false, false, false, false, false, false, false, false},
                    new[] { true, false, false, false, false, false, false, false, false, false, true, false }
                },
                sortOrder = "0",
                state = "",
                stateLevel = "0",
                supplierSector = 0,
                trend = "current",
                tribalLandId = ""
            };

            var jsonObject = JsonConvert.SerializeObject(myData);
            var stringContent = new StringContent(jsonObject.ToString(), System.Text.Encoding.UTF8, "application/json");
            //client.Timeout = new TimeSpan(0, 5, 0);
            try
            {
                HttpResponseMessage response = await client.PostAsync("https://ghgdata.epa.gov/ghgp/service/listFacility/", stringContent);
                response.EnsureSuccessStatusCode();

                strReturnData = await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                strError = ex.Message;
                //Console.WriteLine(ex.Message);
            }

            return new APICallResult { data = strReturnData, message = strError };


        }

        private static async Task<string> GetFacilityDetails(string strID, string strYear)
        {
            string strError = "";
            string strReturnData = "";
            string html = string.Empty;
            string url = string.Format(@"https://ghgdata.epa.gov/ghgp/service/xml/{0}?id={1}&et=undefined", strYear,strID);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

            request.Method = "GET";
            //request.Headers["cache-control"] = "no-store, no-cache, must-revalidate";
            ////request.Headers["connection"] = "Keep-Alive";
            ////request.Connection = "Keep-Alive";
            //request.Headers["content-disposition"] = "attachment; filename=flight.xls";
            ////request.Headers["content-type"] = "application/ms-excel;charset=UTF-8";
            //request.ContentType = "application/ms-excel;charset=UTF-8";
            //request.Headers["keep-alive"] = "timeout=5, max=100";
            //request.Headers["pragma"] = "no-cache";
            //request.Headers["server"] = "Apache";
            //request.Headers["strict-transport-security"] = "max-age=31536000; includeSubDomainsâ";
            ////request.Headers["transfer-encoding"] = "chunked";
            ////request.SendChunked = true;
            //request.Headers["x-frame-options"] = "SAMEORIGIN";
            //request.Host = "ghgdata.epa.gov";
            //request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36";
            //request.Accept = "*/*";
            //request.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-us,en;q=0.5");




            request.AutomaticDecompression = DecompressionMethods.GZip;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                html = reader.ReadToEnd();
            }

            return html;


        }

        /// <summary>
        /// Extract GHGRP ID from string
        /// </summary>
        static string getGHGRPID(string str)
        {
            string strReturn = "";
            int iStart = 0;
            int iEnd = 0;

            try
            {
                iStart = str.IndexOf("[");
                iEnd = str.IndexOf("]");
                strReturn = str.Substring(iStart+1, iEnd - iStart-1);
            }
            catch (Exception)
            {
                
            }

            return strReturn;
        }

        /// <summary>
        /// Extract GHGRP ID and facility name from string
        /// </summary>
        static Tuple<string, string> getGHGRPID2(string str)
        {
            string strFacility = "";
            string strID = "";
            int iStart = 0;
            int iEnd = 0;

            try
            {
                iStart = str.IndexOf("[");
                iEnd = str.IndexOf("]");
                strID = str.Substring(iStart + 1, iEnd - iStart - 1);
                strFacility = str.Substring(0, iStart - 1).Trim();
            }
            catch (Exception)
            {

            }

            return Tuple.Create(strFacility, strID); ;
        }

        /// <summary>
        /// Onshore Production Report
        /// </summary>
        static List<FacilityDetails> getReport_OP(string strBasinID, string strReportYear)
        {
            int iCount = 0;
            List<FacilityDetails> lstReturns = new List<FacilityDetails>();
            Task<APICallResult> callTask = Task.Run(() => GetBasinIDList(strBasinID, strReportYear));
            callTask.Wait();
            APICallResult callReturn = callTask.Result;

            var model = JsonConvert.DeserializeObject<RootObject>(callReturn.data);
            int iTotalCount = model.data.rows.Count;
            const int interval = 10;
            int nextPercent = interval;
            for (int i = 0; i < iTotalCount; i++)
            {
                //logError(i.ToString());
                //if (i < 58)
                //{
                //    continue;
                //    //break;
                //}

                Row item = model.data.rows[i];
                string strGHGRPID = getGHGRPID(item.facility);
                Task<string> callDetailTask = Task.Run(() => GetFacilityDetails(strGHGRPID, strReportYear));
                callDetailTask.Wait();
                string strDetails = callDetailTask.Result;

                FacilityDetails fd = extractFacilityDetails(strDetails);
                lstReturns.Add(fd);

                int currentPercent = (i * 100) / iTotalCount;
                if (currentPercent >= nextPercent)
                {
                    Console.WriteLine(string.Format("      {0}% Completed!", nextPercent));
                    nextPercent = currentPercent - (currentPercent % interval) + interval;
                }

            }
            Console.WriteLine(string.Format("      {0}% Completed!", 100));

            return lstReturns;
        }

        /// <summary>
        /// Onshore Gathering and Boosting Report
        /// </summary>
        static List<FacilityDetails> getReport_OGB(string strBasinID, string strReportYear)
        {
            int iCount = 0;
            List<FacilityDetails> lstReturns = new List<FacilityDetails>();
            Dictionary<string, string> Facility_dictionary = new Dictionary<string, string>();

            Task<APICallResult> callTask = Task.Run(() => GetFacilityList(strBasinID, strReportYear));
            callTask.Wait();
            APICallResult callReturn_FacilityList = callTask.Result;

            callTask = Task.Run(() => GetBasinGeoIDList(strBasinID, strReportYear));
            callTask.Wait();
            APICallResult callReturn_BaisnGeoID = callTask.Result;

            //Get the dictionary
            var model = JsonConvert.DeserializeObject<RootObject>(callReturn_FacilityList.data);
            int iTotalCount = model.data.rows.Count;
            for (int i = 0; i < iTotalCount; i++)
            {
                Row item = model.data.rows[i];
                string strFacilityID = item.facility;
                Tuple<string, string> tFacilityID = getGHGRPID2(strFacilityID);
                Facility_dictionary.Add(tFacilityID.Item1, tFacilityID.Item2);

            }

            //Retrieve report by ID
            model = JsonConvert.DeserializeObject<RootObject>(callReturn_BaisnGeoID.data);
            iTotalCount = model.data.rows.Count;
            const int interval = 10;
            int nextPercent = interval;
            for (int i = 0; i < iTotalCount; i++)
            {
                //logError(i.ToString());
                //if (i < 58)
                //{
                //    continue;
                //    //break;
                //}

                Row item = model.data.rows[i];
                string strFacilityName = item.facility.Trim();

                if (Facility_dictionary.ContainsKey(strFacilityName))
                {
                    string strGHGRPID = Facility_dictionary[strFacilityName];
                    Task<string> callDetailTask = Task.Run(() => GetFacilityDetails(strGHGRPID, strReportYear));
                    callDetailTask.Wait();
                    string strDetails = callDetailTask.Result;

                    FacilityDetails fd = extractFacilityDetails(strDetails);
                    lstReturns.Add(fd);
                }
                

                int currentPercent = (i * 100) / iTotalCount;
                if (currentPercent >= nextPercent)
                {
                    Console.WriteLine(string.Format("      {0}% Completed!", nextPercent));
                    nextPercent = currentPercent - (currentPercent % interval) + interval;
                }

            }
            Console.WriteLine(string.Format("      {0}% Completed!", 100));

            return lstReturns;
        }

        /// <summary>
        /// Export data to excel spreadsheet
        /// </summary>
        static void exportExcel (string strExcelName)
        {
            
            using (ExcelPackage excel = new ExcelPackage())
            {
                int iExcelRow = 0;
                string headerRange = "";
                List<string[]> headerRow;

                //Bakken
                if (WIB_Report_Basin!=null & WIB_Report_BasinGeo != null)
                {
                    iExcelRow = 0;
                    var worksheet = excel.Workbook.Worksheets.Add("Bakken");
                    worksheet.Cells["B1:E1"].Merge = true;
                    worksheet.Cells["F1:J1"].Merge = true;

                    // Popular header row data
                    worksheet.Cells["A1"].Value = "Bakken";
                    worksheet.Cells["B1"].Value = "Onshore Production";
                    worksheet.Cells["F1"].Value = "Onshore Gathering and Boosting";
                    worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    headerRow = new List<string[]>()
                    {
                      new string[] { "Parent Company", "CO2 Tonnes", "CH4 Tonnes", "Gas Sales (mcf)", "Oil Sales (bbls)", "Parent Company", "CO2 Tonnes", "CH4 Tonnes", "Gas Received (mcf)", "Oil Received (bbls)" }
                    };
                    headerRange = "A2:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "2";
                    worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                    //Onshore Production
                    for (int i = 0; i < WIB_Report_Basin.Count; i++)
                    {
                        FacilityDetails fd = WIB_Report_Basin[i];
                        List<string[]> facilityRow = new List<string[]>()
                        {
                          new string[] {fd.ParentCompanyLegalName,fd.CarbonDioxide,fd.Methane,fd.GasSales,fd.OilSales }
                        };
                        iExcelRow = 3 + i;
                        string rowRange = "A" + iExcelRow.ToString() + ":" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + iExcelRow.ToString();
                        worksheet.Cells[rowRange].LoadFromArrays(facilityRow);
                    }

                    //Onshore Gathering and Boosting
                    for (int i = 0; i < WIB_Report_BasinGeo.Count; i++)
                    {
                        FacilityDetails fd = WIB_Report_BasinGeo[i];
                        List<string[]> facilityRow = new List<string[]>()
                        {
                          new string[] {fd.ParentCompanyLegalName,fd.CarbonDioxide,fd.Methane,fd.QuantityGasReceived,fd.QuantityOfHydrocarbonLiquidsReceived }
                        };
                        iExcelRow = 3 + i;
                        string rowRange = "F" + iExcelRow.ToString() + ":" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + iExcelRow.ToString();
                        worksheet.Cells[rowRange].LoadFromArrays(facilityRow);
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }



                //STX
                if (STX_Report_Basin!=null & STX_Report_BasinGeo != null)
                {
                    iExcelRow = 0;
                    var worksheet = excel.Workbook.Worksheets.Add("Eagle Ford");
                    worksheet.Cells["B1:E1"].Merge = true;
                    worksheet.Cells["F1:J1"].Merge = true;

                    // Popular header row data
                    worksheet.Cells["A1"].Value = "Eagle Ford";
                    worksheet.Cells["B1"].Value = "Onshore Production";
                    worksheet.Cells["F1"].Value = "Onshore Gathering and Boosting";
                    worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    headerRow = new List<string[]>()
                    {
                      new string[] { "Parent Company", "CO2 Tonnes", "CH4 Tonnes", "Gas Sales (mcf)", "Oil Sales (bbls)", "Parent Company", "CO2 Tonnes", "CH4 Tonnes", "Gas Received (mcf)", "Oil Received (bbls)" }
                    };
                    headerRange = "A2:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "2";
                    worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                    //Onshore Production
                    for (int i = 0; i < STX_Report_Basin.Count; i++)
                    {
                        FacilityDetails fd = STX_Report_Basin[i];
                        List<string[]> facilityRow = new List<string[]>()
                        {
                          new string[] {fd.ParentCompanyLegalName,fd.CarbonDioxide,fd.Methane,fd.GasSales,fd.OilSales }
                        };
                        iExcelRow = 3 + i;
                        string rowRange = "A" + iExcelRow.ToString() + ":" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + iExcelRow.ToString();
                        worksheet.Cells[rowRange].LoadFromArrays(facilityRow);
                    }

                    //Onshore Gathering and Boosting
                    for (int i = 0; i < STX_Report_BasinGeo.Count; i++)
                    {
                        FacilityDetails fd = STX_Report_BasinGeo[i];
                        List<string[]> facilityRow = new List<string[]>()
                        {
                          new string[] {fd.ParentCompanyLegalName,fd.CarbonDioxide,fd.Methane,fd.QuantityGasReceived,fd.QuantityOfHydrocarbonLiquidsReceived }
                        };
                        iExcelRow = 3 + i;
                        string rowRange = "F" + iExcelRow.ToString() + ":" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + iExcelRow.ToString();
                        worksheet.Cells[rowRange].LoadFromArrays(facilityRow);
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }



                //ABO
                if (APB_Report_Basin!=null & APB_Report_BasinGeo != null)
                {
                    iExcelRow = 0;
                    var worksheet = excel.Workbook.Worksheets.Add("ABO");
                    worksheet.Cells["B1:E1"].Merge = true;
                    worksheet.Cells["F1:J1"].Merge = true;

                    // Popular header row data
                    worksheet.Cells["A1"].Value = "ABO";
                    worksheet.Cells["B1"].Value = "Onshore Production";
                    worksheet.Cells["F1"].Value = "Onshore Gathering and Boosting";
                    worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    headerRow = new List<string[]>()
                    {
                      new string[] { "Parent Company", "CO2 Tonnes", "CH4 Tonnes", "Gas Sales (mcf)", "Oil Sales (bbls)", "Parent Company", "CO2 Tonnes", "CH4 Tonnes", "Gas Received (mcf)", "Oil Received (bbls)" }
                    };
                    headerRange = "A2:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "2";
                    worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                    //Onshore Production
                    for (int i = 0; i < APB_Report_Basin.Count; i++)
                    {
                        FacilityDetails fd = APB_Report_Basin[i];
                        List<string[]> facilityRow = new List<string[]>()
                        {
                          new string[] {fd.ParentCompanyLegalName,fd.CarbonDioxide,fd.Methane,fd.GasSales,fd.OilSales }
                        };
                        iExcelRow = 3 + i;
                        string rowRange = "A" + iExcelRow.ToString() + ":" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + iExcelRow.ToString();
                        worksheet.Cells[rowRange].LoadFromArrays(facilityRow);
                    }

                    //Onshore Gathering and Boosting
                    for (int i = 0; i < APB_Report_BasinGeo.Count; i++)
                    {
                        FacilityDetails fd = APB_Report_BasinGeo[i];
                        List<string[]> facilityRow = new List<string[]>()
                        {
                          new string[] {fd.ParentCompanyLegalName,fd.CarbonDioxide,fd.Methane,fd.QuantityGasReceived,fd.QuantityOfHydrocarbonLiquidsReceived }
                        };
                        iExcelRow = 3 + i;
                        string rowRange = "F" + iExcelRow.ToString() + ":" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + iExcelRow.ToString();
                        worksheet.Cells[rowRange].LoadFromArrays(facilityRow);
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }
                


                FileInfo excelFile = new FileInfo(strExcelName);
                excel.SaveAs(excelFile);
            }
        }

        /// <summary>
        /// Check whether output file can be saved
        /// </summary>
        static Boolean checkFileValid(string strExcelName)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                FileInfo excelFile = new FileInfo(strExcelName);
                try
                {
                    excel.Workbook.Worksheets.Add("Worksheet1");
                    excel.SaveAs(excelFile);
                    System.IO.File.Delete(strExcelName);
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
                
                
            }
                
        }

        /// <summary>
        /// Export error messages
        /// </summary>
        static void logError(string strError)
        {
            using (StreamWriter sw = System.IO.File.AppendText("log_Errors.txt"))
            {
                sw.WriteLine(strError);
                sw.WriteLine("");
                sw.WriteLine("");
            }

        }


        /// <summary>
        /// Export completion message
        /// </summary>
        static void logComplete(DateTime dtStart, string strProgram, string strError)
        {
            DateTime stEnd = DateTime.Now;
            TimeSpan span = stEnd.Subtract(dtStart);
            using (StreamWriter sw = System.IO.File.AppendText("log.txt"))
            {
                sw.WriteLine("*********************************************");
                sw.WriteLine(strProgram + " Completed!");
                sw.WriteLine("Started at: " + dtStart.ToString("yyyy-MM-dd HH:mm:ss"));
                sw.WriteLine("Ended at: " + stEnd.ToString("yyyy-MM-dd HH:mm:ss"));
                sw.WriteLine("Total Running Time (minutes): " + span.TotalMinutes.ToString("N1"));
                if (strError!="")
                {
                    sw.WriteLine("Error Found when Running Automation Script:");
                    sw.WriteLine(strError);
                }
                sw.WriteLine("*********************************************");
                sw.WriteLine("");
                sw.WriteLine("");
                sw.WriteLine("");
            }

        }

        /// <summary>
        /// Main process
        /// </summary>
        static void Main(string[] args)
        {
            #region App initialization   
            Console.WriteLine("Initializing Application...");

            DateTime dateStart = DateTime.Now;
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Ssl3
                                        | System.Net.SecurityProtocolType.Tls
                                        | System.Net.SecurityProtocolType.Tls11
                                        | System.Net.SecurityProtocolType.Tls12;

            string strYear = "";
            string strOutputFile = "";
            string strErr = "";
            if (File.Exists(@"Settings.txt"))
            {
                // Setting file exist
                string[] readConfig = System.IO.File.ReadAllLines(@"Settings.txt");
                strYear = (readConfig[0].Split('='))[1].Trim();
                strOutputFile = (readConfig[1].Split('='))[1].Trim();
            }
            else
            {
                strYear = (DateTime.Now.Year-1).ToString();
                strOutputFile = "EPA.xlsx";
            }

            #endregion


            #region Testing
            //RunAsync().Wait();
            //string url = @"https://ghgdata.epa.gov/ghgp/service/export?q=&tr=current&ds=O&ryr=2017&cyr=2017&lowE=-20000&highE=23000000&st=&fc=&mc=&rs=ALL&sc=0&is=11&et=&tl=&pn=undefined&ol=0&sl=0&bs=160A&g1=1&g2=1&g3=0&g4=0&g5=0&g6=0&g7=0&g8=0&g9=0&g10=0&g11=0&g12=0&s1=0&s2=0&s3=0&s4=0&s5=0&s6=0&s7=0&s8=0&s9=1&s10=0&s201=0&s202=0&s203=0&s204=0&s301=0&s302=0&s303=0&s304=0&s305=0&s306=0&s307=0&s401=0&s402=0&s403=0&s404=0&s405=0&s601=0&s602=0&s701=0&s702=0&s703=0&s704=0&s705=0&s706=0&s707=0&s708=0&s709=0&s710=0&s711=0&s801=0&s802=0&s803=0&s804=0&s805=0&s806=0&s807=0&s808=0&s809=0&s810=0&s901=0&s902=1&s903=0&s904=0&s905=0&s906=0&s907=0&s908=0&s909=0&s910=0&s911=0&sf=11001000&listExport=false"; ;

            //GetAsync(url).Wait();
            //doTesting();
            //return;
            #endregion

            #region Main
            if (!checkFileValid(strOutputFile))
            {
                logComplete(dateStart, "EPA Data Download Automation", "Output file is open. Close it and try again!");
                return;
            }

            
            try
            {
                Console.WriteLine("Start Downloading Onshore Production Data from EPA Website...");

                #region Getting data from STX (220)
                Console.WriteLine("   Retriving Data for Gulf Coast Basin");
                STX_Report_Basin = getReport_OP("220", strYear);
                #endregion

                #region Getting data from APB (160A)
                Console.WriteLine("   Retriving Data for Appalachian Basin-Eastern Overthrust");
                APB_Report_Basin = getReport_OP("160A", strYear);
                #endregion

                #region Getting data from WIB (395)
                Console.WriteLine("   Retriving Data for Williston Basin");
                WIB_Report_Basin = getReport_OP("395", strYear);
                #endregion


                Console.WriteLine("Start Downloading Onshore Gathering and Boosting Data from EPA Website...");

                #region Getting data from STX (220)
                Console.WriteLine("   Retriving Data for Gulf Coast Basin");
                STX_Report_BasinGeo = getReport_OGB("220", strYear);
                #endregion

                #region Getting data from APB (160A)
                Console.WriteLine("   Retriving Data for Appalachian Basin-Eastern Overthrust");
                APB_Report_BasinGeo = getReport_OGB("160A", strYear);
                #endregion

                #region Getting data from WIB (395)
                Console.WriteLine("   Retriving Data for Williston Basin");
                WIB_Report_BasinGeo = getReport_OGB("395", strYear);
                #endregion


                #region Export spreadsheet
                Console.WriteLine("Exporting Data to Excel Spreadsheet...");
                exportExcel(strOutputFile);
                #endregion
                
            }
            catch (Exception ex)
            {
                strErr = ex.Message;
            }
            finally
            {
                logComplete(dateStart, "EPA Data Download Automation",strErr);
            }
            #endregion

            Console.WriteLine("Automation Completed!");
        }
    }
}
