// Copyright 2018 ESRI
// 
// All rights reserved under the copyright laws of the United States
// and applicable international laws, treaties, and conventions.
// 
// You may freely redistribute and use this sample code, with or
// without modification, provided you include the original copyright
// notice and use restrictions.
// 
// See the use restrictions at <your Enterprise SDK install location>/userestrictions.txt.
// 

using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Server;
using ESRI.Server.SOESupport;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;

using Aspose.Words;
using Aspose.Words.Replacing;
using NPOI.HSSF.UserModel;

using Esri.FileGDB;

//This is REST SOE template of Enterprise SDK

//TODO: sign the project (project properties > signing tab > sign the assembly)
//      this is strongly suggested if the dll will be registered using regasm.exe <your>.dll /codebase


namespace soeXls2Odt
{
    [ComVisible(true)]
    [Guid("6abcf7fb-ae62-4357-ae90-8ec1440f82d1")]
    [ClassInterface(ClassInterfaceType.None)]
    [ServerObjectExtension("MapServer",
        AllCapabilities = "",
        DefaultCapabilities = "",
        Description = "soeXls2Odt",
        DisplayName = "soeXls2Odt",
        Properties = "odtTemplate=D:\\Win\\xls2odt\\Template3.odt"
                     ,                     
        SupportsREST = true,
        SupportsSOAP = false,
        SupportsSharedInstances = false)]
    public class soeXls2Odt : IServerObjectExtension, IObjectConstruct, IRESTRequestHandler
    {
        private string soe_name;

        private IPropertySet configProps;
        private IServerObjectHelper serverObjectHelper;
        private ServerLogger logger;
        private IRESTRequestHandler reqHandler;

        private string odtTemplate = "";        
        private string localFilePath = string.Empty;
        private string virtualFilePath = string.Empty;

        public soeXls2Odt()
        {
            soe_name = this.GetType().Name;
            logger = new ServerLogger();
            reqHandler = new SoeRestImpl(soe_name, CreateRestSchema()) as IRESTRequestHandler;
        }

        #region IServerObjectExtension Members

        public void Init(IServerObjectHelper pSOH)
        {
            serverObjectHelper = pSOH;
            IMapServer ms = (IMapServer)pSOH.ServerObject;
            String outputDir = outputDir = ms.PhysicalOutputDirectory;
            int len = 0;
            if (outputDir != null)
                len = outputDir.Length;
            if (len > 0)
            {
                String mapservicePath = outputDir.Substring(0, len - 1);
                int mapServiceIndex = mapservicePath.LastIndexOf("\\");
                virtualFilePath = mapservicePath.Substring(mapServiceIndex + 1);
                localFilePath = mapservicePath;
            }
            if (string.IsNullOrEmpty(localFilePath))
            {
                logger.LogMessage(ServerLogger.msgType.error, soe_name + ".init()", 500, "OutputDirectory is empty or missing. Reset to default.");
            }
            logger.LogMessage(ServerLogger.msgType.infoStandard, soe_name + ".init()", 200, "Initialized " + soe_name + " SOE.");
        }

        public void Shutdown()
        {
        }

        #endregion

        #region IObjectConstruct Members

        public void Construct(IPropertySet props)
        {
            configProps = props;

            this.odtTemplate = (string)props.GetProperty("odtTemplate");            
        }

        #endregion

        #region IRESTRequestHandler Members

        public string GetSchema()
        {
            return reqHandler.GetSchema();
        }

        public byte[] HandleRESTRequest(string Capabilities, string resourceName, string operationName, string operationInput, string outputFormat, string requestProperties, out string responseProperties)
        {
            return reqHandler.HandleRESTRequest(Capabilities, resourceName, operationName, operationInput, outputFormat, requestProperties, out responseProperties);
        }

        #endregion

        private RestResource CreateRestSchema()
        {
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12;

            RestResource rootRes = new RestResource(soe_name, false, RootResHandler);

            RestOperation sampleOper = new RestOperation("sampleOperation",
                                                      new string[] { "parm1", "parm2" },
                                                      new string[] { "json" },
                                                      SampleOperHandler);

            rootRes.operations.Add(sampleOper);

            RestOperation disasterXls2Odt_Op = new RestOperation("disasterXls2Odt",
                                                   new string[] { "xls file path", "save to odt path" },
                                                   new string[] { "json" },
                                                   disasterXls2Odt);
            rootRes.operations.Add(disasterXls2Odt_Op);

            RestOperation uploadFile_Op = new RestOperation("uploadFile",
                                                   new string[] { "file base64", "file save to" },
                                                   new string[] { "json" },
                                                   uploadFile);
            rootRes.operations.Add(uploadFile_Op);

            RestOperation downloadFile_Op = new RestOperation("downloadFile",
                                                   new string[] { "server file path" },
                                                   new string[] { "json" },
                                                   downloadFile);
            rootRes.operations.Add(downloadFile_Op);


            return rootRes;
        }

        private byte[] RootResHandler(NameValueCollection boundVariables, string outputFormat, string requestProperties, out string responseProperties)
        {
            responseProperties = null;

            JsonObject result = new JsonObject();
            result.AddString("hello", "world");

            return Encoding.UTF8.GetBytes(result.ToJson());
        }

        private byte[] createErrorObject(int codeNumber, String errorMessageSummary, String[] errorMessageDetails)
        {
            if (errorMessageSummary.Length == 0 || errorMessageSummary == null)
            {
                throw new Exception("Invalid error message specified.");
            }

            JSONObject errorJSON = new JSONObject();
            errorJSON.AddLong("code", codeNumber);
            errorJSON.AddString("message", errorMessageSummary);

            if (errorMessageDetails == null)
            {
                errorJSON.AddString("details", "No error details specified.");
            }
            else
            {
                String errorMessages = "";
                for (int i = 0; i < errorMessageDetails.Length; i++)
                {
                    errorMessages = errorMessages + errorMessageDetails[i] + "\n";
                }

                errorJSON.AddString("details", errorMessages);
            }

            JSONObject error = new JSONObject();
            errorJSON.AddJSONObject("error", errorJSON);

            return Encoding.UTF8.GetBytes(errorJSON.ToJSONString(null));
        }

        private byte[] SampleOperHandler(NameValueCollection boundVariables,
                                                  JsonObject operationInput,
                                                      string outputFormat,
                                                      string requestProperties,
                                                  out string responseProperties)
        {
            responseProperties = null;

            string parm1Value;
            bool found = operationInput.TryGetString("parm1", out parm1Value);
            if (!found || string.IsNullOrEmpty(parm1Value))
                throw new ArgumentNullException("parm1");

            string parm2Value;
            found = operationInput.TryGetString("parm2", out parm2Value);
            if (!found || string.IsNullOrEmpty(parm2Value))
                throw new ArgumentNullException("parm2");

            JsonObject result = new JsonObject();
            result.AddString("parm1", parm1Value);
            result.AddString("parm2", parm2Value);

            try
            {
                Geodatabase.Delete("D:/Win/xls2Odt/FeatureDatasetDemo.gdb");
            }
            catch (FileGDBException ex)
            {
            }

                Geodatabase geodatabase = Geodatabase.Create("D:/Win/xls2Odt/FeatureDatasetDemo.gdb");
            
            string featureDatasetDef = "";
            using (StreamReader sr = new StreamReader("D:/Win/xls2Odt/TransitFD.xml"))
            {
                while (sr.Peek() >= 0)
                {
                    featureDatasetDef += sr.ReadLine() + "\n";
                }
                sr.Close();
            }
            geodatabase.CreateFeatureDataset(featureDatasetDef);
            string tableDef = "";
            using (StreamReader sr = new StreamReader("D:/Win/xls2Odt/BusStopsTable.xml"))
            {
                while (sr.Peek() >= 0)
                {
                    tableDef += sr.ReadLine() + "\n";
                }
                sr.Close();
            }            
            Esri.FileGDB.Table table = geodatabase.CreateTable(tableDef, "\\Transit");
            
            // Close the table.
            table.Close();

            // Close the geodatabase
            geodatabase.Close();


            return Encoding.UTF8.GetBytes(result.ToJson());
        }

        private byte[] downloadFile(NameValueCollection boundVariables,
                                                  JsonObject operationInput,
                                                      string outputFormat,
                                                      string requestProperties,
                                                  out string responseProperties)
        {
            responseProperties = "";

            string download_file_name;
            bool found = operationInput.TryGetString("server file path", out download_file_name);
            if (!found || string.IsNullOrEmpty(download_file_name))
                throw new ArgumentNullException("server file path");

            if( !File.Exists(download_file_name) ) 
                throw new ArgumentNullException("server file not found");

            var extension = System.IO.Path.GetExtension(download_file_name);

            string fileId = Guid.NewGuid().ToString("N");
            string fileName = "testFile_" + fileId + extension;
            string file = localFilePath + "\\" + fileName;

            byte[] fileBytes = System.IO.File.ReadAllBytes(download_file_name);
            System.IO.File.WriteAllBytes(file, fileBytes);
            long fileSize = new System.IO.FileInfo(file).Length;

            if (outputFormat == "json")
            {
                responseProperties = "{\"Content-Type\" : \"application/json\"}";
                string requestURL = ServerUtilities.GetServerEnvironment().Properties.GetProperty("RequestContextURL") as string;
                string fileVirutualURL = requestURL + "/rest/directories/arcgisoutput/" + virtualFilePath + "/" + fileName;
                JsonObject jsonResult = new JsonObject();
                jsonResult.AddString("url", fileVirutualURL);
                jsonResult.AddString("fileName", fileName);
                jsonResult.AddString("fileSizeBytes", Convert.ToString(fileSize));
                return Encoding.UTF8.GetBytes(jsonResult.ToJson());

            }
            else if (outputFormat == "file")
            {
                responseProperties = "{\"Content-Type\" : \"application/octet-stream\",\"Content-Disposition\": \"attachment; filename=" + fileName + "\"}";
                return System.IO.File.ReadAllBytes(file);
            }
            return Encoding.UTF8.GetBytes("");
        }

        private byte[] uploadFile(NameValueCollection boundVariables,
                                                    JsonObject operationInput,
                                                    string outputFormat,
                                                    string requestProperties,
                                                    out string responseProperties)
        {
            responseProperties = null;

            string fileBase64;
            bool found = operationInput.TryGetString("file base64", out fileBase64);
            if (!found || string.IsNullOrEmpty(fileBase64))
                throw new ArgumentNullException("file base64");

            string saveFilePath;
            found = operationInput.TryGetString("file save to", out saveFilePath);
            if (!found || string.IsNullOrEmpty(saveFilePath))
                throw new ArgumentNullException("file save to");

            // decode 存檔
            byte[] fileBytes = Convert.FromBase64String(fileBase64);
            File.WriteAllBytes(saveFilePath, fileBytes);

            return Encoding.UTF8.GetBytes( "{\"Success\":true,\"Message\":\"OK\"}" );
        }


        private byte[] disasterXls2Odt(NameValueCollection boundVariables,
                                                    JsonObject operationInput,
                                                    string outputFormat,
                                                    string requestProperties,
                                                    out string responseProperties)
        {
            responseProperties = null;

            string xlsFilePath;
            bool found = operationInput.TryGetString("xls file path", out xlsFilePath);
            if (!found || string.IsNullOrEmpty(xlsFilePath))
                throw new ArgumentNullException("xls file path");

            string odtSavePath;
            found = operationInput.TryGetString("save to odt path", out odtSavePath);
            if (!found || string.IsNullOrEmpty(odtSavePath))
                throw new ArgumentNullException("save to odt path");

            // 轉換
            string retJson = disasterXls2Odt_sub(xlsFilePath, odtSavePath);

            return Encoding.UTF8.GetBytes( retJson);
        }

        private string disasterXls2Odt_sub(string xlsFilePath, string odtSavePath)
        {
            string retJson = "{";
            try
            {
                Document doc = new Document(this.odtTemplate);

                Dictionary<string, string> replacements = ReadXlsPage1(xlsFilePath);

                // 套入 odt 
                foreach (var entry in replacements)
                {
                    string placeholder = string.Format("${{{0}}}", entry.Key);
                    doc.Range.Replace(placeholder, entry.Value, new FindReplaceOptions(FindReplaceDirection.Forward));
                }

                // 讀取 excel 次頁地號項目
                var listItems = ReadXlsPage2(xlsFilePath);

                // 進行替換
                for(int i=0;i<5;i++) {
                    string placeholder = string.Format("${{{0}}}", "sectName_"+i);
                    doc.Range.Replace(placeholder, listItems[i,0], new FindReplaceOptions(FindReplaceDirection.Forward));
                    placeholder = string.Format("${{{0}}}", "landNo_"+i);
                    doc.Range.Replace(placeholder, listItems[i,1], new FindReplaceOptions(FindReplaceDirection.Forward));
                    placeholder = string.Format("${{{0}}}", "area_"+i);
                    doc.Range.Replace(placeholder, listItems[i,2], new FindReplaceOptions(FindReplaceDirection.Forward));
                }

                // 讀取作物種類套入 checkbox
                string cropType = ReadPage2CropType(xlsFilePath);

                string holder = string.Format("${{{0}}}", "cropType");
                doc.Range.Replace(holder, cropType, new FindReplaceOptions(FindReplaceDirection.Forward));

                // 讀取作物補助各欄
                replacements = ReadXlsPage2CropItems(xlsFilePath);

                // 套入 odt
                foreach (var entry in replacements)
                {
                    string placeholder = string.Format("${{{0}}}", entry.Key);
                    doc.Range.Replace(placeholder, entry.Value, new FindReplaceOptions(FindReplaceDirection.Forward));
                }

                // 儲存修改後的文件為 ODT
                doc.Save(odtSavePath, SaveFormat.Odt);

                retJson += "\"Success\":true,";
                retJson += "\"Message\":\""+odtSavePath+"\"";
            }
            catch (Exception ex)
            {
                retJson += "\"Success\":false,";
                retJson += "\"Message\":\""+ex.ToString()+"\"";
            }

            retJson += "}";

            return retJson;
        }

        private Dictionary<string, string> ReadXlsPage2CropItems(string filePath)
        {
            Dictionary<string, string> data = new Dictionary<string, string>();
            data["cropName"] = "";
            data["cropDegree"] = "";
            data["cropArea"] = "";
            data["cropMoney"] = "";

            using (var file = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new HSSFWorkbook(file);
                var worksheet = workbook.GetSheetAt(1);

                int rowIndex = 0;
                while (rowIndex<5 && worksheet.GetRow(rowIndex+6) != null && !worksheet.GetRow(rowIndex+6).GetCell(0).ToString().Equals("合計") )
                {
                    data["cropName"] += ">"+worksheet.GetRow(rowIndex+6).GetCell(11).ToString()+"\r\n";
                    data["cropDegree"] += ">"+worksheet.GetRow(rowIndex+6).GetCell(14).ToString()+"\r\n";
                    data["cropArea"] += ">"+worksheet.GetRow(rowIndex+6).GetCell(13).ToString()+"\r\n";
                    data["cropMoney"] += ">"+worksheet.GetRow(rowIndex+6).GetCell(16).ToString()+"\r\n";
                    rowIndex++;
                }
            }

            return data;
        }

        private string ReadPage2CropType(string xlsPath)
        {
            // 讀取 xls 作物，準備依此勾選
            string cropItems = "";   // 空白區隔，串成一串即可(因僅需判斷是否有填此作物依據)
            using (var file = new System.IO.FileStream(xlsPath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new HSSFWorkbook(file);
                var worksheet = workbook.GetSheetAt(1);

                int rowIndex = 0;
                while (rowIndex<5 && worksheet.GetRow(rowIndex+6) != null && !worksheet.GetRow(rowIndex+6).GetCell(0).ToString().Equals("合計") )
                {
                    cropItems += worksheet.GetRow(rowIndex+6).GetCell(11).ToString()+" ";
                    rowIndex++;
                }
            }

            Dictionary<string, string> fieldMapping = new Dictionary<string, string>
            {
                {"稻米", "□稻米"},
                {"雜糧", "□雜糧"},
                {"果樹", "□果樹"},
                {"花卉", "□花卉"},
                {"菇類", "□菇類"},
                {"蔬菜", "□蔬菜"},
                {"特用作物", "□特用作物(□荖花荖葉□其他)"},
                {"養蜂", "□養蜂(□蜂箱□蜂群(因蜜源缺乏之公告，以完成農民從事養蜂事實申報及登錄作業程序者為限)"},
                {"結構型鋼骨溫網室", "□結構型鋼骨溫網室"},
                {"簡易式塑膠布溫網室", "□簡易式塑膠布溫網室"},
                {"水平棚架網室", "□水平棚架網室"},
                {"溫網室以外之農業設施", "□溫網室以外之農業設施"},
                {"菇舍", "□菇舍"},
                {"製茶設備設施", "□製茶設備設施"}
            };

            StringBuilder checkboxSection = new StringBuilder();

            // 逐一檢查 XLS 中的欄位值並替換勾選符號
            foreach (var entry in fieldMapping)
            {
                // 判斷此筆組合的正確 value(該勾的勾)
                string itemText = getEntryValue(entry.Key,entry.Value,cropItems);
                checkboxSection.Append(itemText );
            }

            return checkboxSection.ToString();

        }

        private string getEntryValue(string check_key, string check_value, string crop_items)
        {
            string ret_value = check_value;

            // 判斷方式，如果 check_key 直接在 crop_items 中，則直接替換成實心方塊回傳
            if( crop_items.Contains(check_key) )
                ret_value = check_value.Replace("□", "■");
            else if( check_key == "果樹") {
                if( crop_items.Contains("洋香瓜") )
                    ret_value = check_value.Replace("□", "■");
            }

            return ret_value;
        }


        private Dictionary<string, string> ReadXlsPage1(string filePath)
        {
            Dictionary<string, string> data = new Dictionary<string, string>();

            using (var file = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // 使用 HSSFWorkbook 來讀取 .xls 格式
                var workbook = new HSSFWorkbook(file);
                var worksheet = workbook.GetSheetAt(0); // 取得第一個工作表

                data["disasterName"] = worksheet.GetRow(3).GetCell(3).ToString();
                data["applyDate"] = worksheet.GetRow(3).GetCell(6).ToString();
                data["applyName"] = worksheet.GetRow(4).GetCell(3).ToString();
                data["idNumber"] = worksheet.GetRow(4).GetCell(6).ToString();
                data["address"] = worksheet.GetRow(5).GetCell(3).ToString();
                data["telephone"] = worksheet.GetRow(5).GetCell(6).ToString();
            }

            return data;
        }

        private string[,] ReadXlsPage2(string filePath)
        {
            string[,] listItems = new string[5,3];
            for(int i=0;i<5;i++)
                for(int j=0;j<3;j++)
                    listItems[i,j]="";

            using (var file = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new HSSFWorkbook(file);
                var worksheet = workbook.GetSheetAt(1);

                // 鄉鎮
                var city = worksheet.GetRow(2).GetCell(1).ToString();

                int rowIndex = 0;
                while (rowIndex<5 && worksheet.GetRow(rowIndex+6) != null && !worksheet.GetRow(rowIndex+6).GetCell(0).ToString().Equals("合計") )
                {
                    listItems[rowIndex,0] = city+worksheet.GetRow(rowIndex+6).GetCell(1).ToString()+"段"+
                                            worksheet.GetRow(rowIndex+6).GetCell(2).ToString()+"小段";
                    listItems[rowIndex,1] = worksheet.GetRow(rowIndex+6).GetCell(3).ToString()+"地號";
                    listItems[rowIndex,2] = worksheet.GetRow(rowIndex+6).GetCell(5).ToString();   // 用權利面積 

                    rowIndex++;
                }
            }

            return listItems;
        }


    }
}
