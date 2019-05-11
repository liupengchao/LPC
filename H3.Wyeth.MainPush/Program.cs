using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using NPOI.HSSF;
using H3.SHHS.Excel;
using H3.BizBus;
using System.IO;
using System.Configuration;

namespace H3.Wyeth.MainPush
{
    class Program
    {
        private static string ContractSheetName = "合同信息";
        private static string StoreSheetName = "门店信息";
        private static string FilePath = ConfigurationManager.AppSettings["InitFilePath"] + string.Empty;

        private static com.h3yun.www.BizObjectService _bizService = null;
        public static com.h3yun.www.BizObjectService bizService
        {
            get
            {
                if (_bizService == null)
                {
                    com.h3yun.www.Authentication au = new com.h3yun.www.Authentication();
                    au.Secret = "BFFs4ZtnTzTgSOU5ToguNPvWDDptWKYqfoJLPPSyx/H8j+/9eFliyA==";
                    au.EngineCode = "oyus4n0b3vz0sjaanfk3kxjg2";
                    _bizService = new com.h3yun.www.BizObjectService();
                    _bizService.AuthenticationValue = au;
                    _bizService.Timeout = int.MaxValue;
                }
                return _bizService;
            }
        }

        public static string PutFilePath
        {
            get
            {
                return System.Configuration.ConfigurationManager.AppSettings["PutFilePath"];
            }
        }

        static string SFTPAddress
        {
            get
            {
                return ConfigurationManager.AppSettings["SFTPAddress"] + string.Empty;
            }
        }
        static string SFTPPort
        {
            get
            {
                return ConfigurationManager.AppSettings["SFTPPort"] + string.Empty;
            }
        }
        static string SFTPUserName
        {
            get
            {
                return ConfigurationManager.AppSettings["SFTPUserName"] + string.Empty;
            }
        }
        static string SFTPPassword
        {
            get
            {
                return ConfigurationManager.AppSettings["SFTPPassword"] + string.Empty;
            }
        }
        static string RemotePath
        {
            get
            {
                return ConfigurationManager.AppSettings["RemotePath"] + string.Empty;
            }
        }


        public static Dictionary<string, string> MappingColumn
        {
            get
            {
                Dictionary<string, string> mapping = new Dictionary<string, string>();
                mapping.Add("ContractCode", "合同编码");
                mapping.Add("ApproveState", "审批状态");
                mapping.Add("Zone", "大区");
                mapping.Add("Region", "区域");
                mapping.Add("City", "城市");
                mapping.Add("DealerName", "经销商名称");
                mapping.Add("DealerCode", "经销商代码");
                mapping.Add("DealerRegion", "经销商区域");
                mapping.Add("DealerZone", "经销商大区");
                mapping.Add("DealerCity", "经销商城市");
                mapping.Add("CustomerGroupType", "客户组织类型");
                mapping.Add("CustomerGroupName", "客户组织名称");
                mapping.Add("StoreCount", "客户门店数");
                mapping.Add("ContractStoreCount", "合同门店数");
                mapping.Add("MainStoreCount", "主推门店数");
                mapping.Add("MainStoreRate", "主推门店比率");
                mapping.Add("FirstMonthTarget", "第1个月目标");
                mapping.Add("SecondMonthTarget", "第2个月目标");
                mapping.Add("ThirdMonthTarget", "第3个月目标");
                mapping.Add("InmktTargetTTL", "Inmkt目标TTL");
                mapping.Add("FirstQHYJJ", "第1个Q惠赢基金");
                mapping.Add("HYJJTTL", "惠赢基金TTL");
                mapping.Add("HYJJRate", "惠赢基金费比");
                mapping.Add("IMSTTL", "IMS达成TTL");
                mapping.Add("IMSTTLRate", "IMS达成TTL%");
                mapping.Add("GrowthRate", "全品牌目标对比前3月增长率");
                mapping.Add("Remark", "备注");
                return mapping;
            }
        }

        public static Dictionary<string, string> ChildMappingColumn
        {
            get
            {
                Dictionary<string, string> child = new Dictionary<string, string>();
                child.Add("IsMainPush", "是否主推门店");
                child.Add("IMSTTL", "IMS达成TTL");
                child.Add("StoreCode", "门店代码");
                return child;
            }
        }

        public static Dictionary<string,string> SFTPMappingColumn
        {
            get
            {
                Dictionary<string, string> mapping = new Dictionary<string, string>();
                mapping.Add("ContractCode", "合同编号");
                mapping.Add("IsValid", "是否签订");
                mapping.Add("CreatedBy", "签订人");
                mapping.Add("CreatedTime", "签订时间");
                mapping.Add("ContractLocation", "签订位置");
                mapping.Add("Lng", "经度");
                mapping.Add("Lat", "维度");
                mapping.Add("ChangedCount", "变更次数");
                //mapping.Add("PictureUrl", "合同照片");
                return mapping;
            }
        }

        static void __Main(string[] args)
        {
            //WriteLog("同步合同信息开始...");
            //SyncContract();
            //WriteLog("同步合同信息结束...");
            WriteLog("上传合同签订至SFTP开始...");
            UpdateFileToSFTP();
            WriteLog("上传合同签订至SFTP结束...");
            Console.ReadKey();
            Console.ReadKey();
        }


        #region------------------------------同步主推合同基础数据方法

        static void SyncContract()
        {
            string path = FilePath;
            DataTable contract = GetContract(path);
            DataTable store = GetContractStore(path);
            if (contract != null && contract.Rows.Count > 0)
            {
                foreach (DataRow row in contract.Rows)
                {
                    try
                    {
                        string contractCode = row["合同编码"] + string.Empty;
                        Dictionary<string, object> dic = new Dictionary<string, object>();
                        foreach (string key in MappingColumn.Keys)
                        {
                            dic.Add(key, row[MappingColumn[key]] + string.Empty);
                        }
                        DataRow[] childRows = store.Select("合同编号='" + contractCode + "'");
                        if (childRows != null && childRows.Length > 0)
                        {
                            List<Dictionary<string, string>> children = new List<Dictionary<string, string>>();
                            foreach (DataRow c in childRows)
                            {
                                string storeCode = c["门店代码"] + string.Empty;
                                string storeId = GetStore(storeCode);
                                if (!string.IsNullOrEmpty(storeId))
                                {
                                    Dictionary<string, string> d = new Dictionary<string, string>();
                                    d.Add("Store", storeId);
                                    foreach (string k in ChildMappingColumn.Keys)
                                    {
                                        d.Add(k, c[ChildMappingColumn[k]] + string.Empty);
                                    }
                                    children.Add(d);
                                }
                            }
                            dic.Add("D000365ContractStore", children);
                        }
                        string contractId = GetContractId(contractCode);
                        if (!string.IsNullOrEmpty(contractId))
                        {
                            string r = bizService.UpdateBizObject("D000365Contract", contractId, Newtonsoft.Json.JsonConvert.SerializeObject(dic));
                            if(r.IndexOf("成功")>-1)
                            {
                                WriteLog("更新合同信息成功------->ContractCode:" +contractCode+",ContractId:"+contractId );
                            }
                            else
                            {
                                WriteLog("更新合同信息失败------->ContractId:"+contractId+",ContractCode:" + contractCode+",Error:"+r);
                            }
                        }
                        else
                        {
                            string r = bizService.CreateBizObject("D000365Contract", Newtonsoft.Json.JsonConvert.SerializeObject(dic), true);
                            if (r.IndexOf("成功") > -1)
                            {
                                WriteLog("创建合同信息成功------->ContractCode:" + contractCode);
                            }
                            else
                            {
                                WriteLog("创建合同信息失败------->ContractCode:" + contractCode + ",Error:" + r);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLog("创建合同信息错误---->Error:" + ex.Message);
                    }
                }
            }
        }

        static string GetStore(string storeCode)
        {
            if (string.IsNullOrEmpty(storeCode)) return null;
            Filter f = new Filter();
            And and = new And();
            //and.Add(new ItemMatcher("SKUID", ComparisonOperatorType.NotEqual, ""));
            and.Add(new ItemMatcher("RetailCode", ComparisonOperatorType.Equal, storeCode));
            f.Matcher = and;
            string fstr = H3.BizBus.BizStructureUtility.FilterToJson(f);
            string result = string.Empty;
            result = bizService.LoadBizObjects("D000365ShopInfo", fstr);
            List<Dictionary<string, string>> stores = ConvertToDic(result);
            if (stores != null && stores.Count > 0)
            {
                Dictionary<string, string> store = stores[0];
                return store["ObjectId"];
            }
            return null;
        }

        static string GetContractId(string contractCode)
        {
            if (string.IsNullOrEmpty(contractCode)) return null;
            Filter f = new Filter();
            And and = new And();
            //and.Add(new ItemMatcher("SKUID", ComparisonOperatorType.NotEqual, ""));
            and.Add(new ItemMatcher("ContractCode", ComparisonOperatorType.Equal, contractCode));
            f.Matcher = and;
            string fstr = H3.BizBus.BizStructureUtility.FilterToJson(f);
            string result = string.Empty;
            result = bizService.LoadBizObjects("D000365Contract", fstr);
            List<Dictionary<string, string>> stores = ConvertToDic(result);
            if (stores != null && stores.Count > 0)
            {
                Dictionary<string, string> store = stores[0];
                return store["ObjectId"];
            }
            return null;
        }

        public static List<Dictionary<string, string>> ConvertToDic(string result)
        {
            if (string.IsNullOrEmpty(result)) return null;
            List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();
            try
            {
                Dictionary<string, object> d = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(result);
                List<Dictionary<string, object>> list1 = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(d["Data"] + string.Empty);
                foreach (Dictionary<string, object> dic in list1)
                {
                    if (dic != null && dic.Count > 0)
                    {
                        Dictionary<string, string> dd = new Dictionary<string, string>();
                        foreach (string key in dic.Keys)
                        {
                            dd[key] = dic[key] + string.Empty;
                        }
                        list.Add(dd);
                    }
                }
                //list = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(d["Data"] + string.Empty);
            }
            catch (Exception ex)
            {
                WriteLog("解析H3数据失败------>Error:" + ex.Message);
            }
            return list;

        }

        static DataTable GetContract(string path)
        {
            if (string.IsNullOrEmpty(path)) return null;
            return NPOIHelper.ExcelSheetImportToDataTable(path, ContractSheetName);
        }

        static DataTable GetContractStore(string path)
        {
            if (string.IsNullOrEmpty(path)) return null;
            return NPOIHelper.ExcelSheetImportToDataTable(path, StoreSheetName);
        }

        #endregion

        #region---------------------------------上传文件至SFTP
        
        static void UpdateFileToSFTP()
        {
            string[] files = GetContractFile();
            if(files!=null&&files.Length>0)
            {
                SFTPHelper sFTPHelper = new SFTPHelper(SFTPAddress, SFTPPort, SFTPUserName, SFTPPassword);
                bool connect= sFTPHelper.Connect();
                foreach(string file in files)
                {
                    try
                    {
                        string fileName = Path.GetFileName(file);
                        sFTPHelper.Put(file, RemotePath+fileName);
                        WriteLog("上传文件至SFTP成功------>File:" + file);
                    }
                    catch(Exception ex)
                    {
                        WriteLog("上传文件至SFTP失败------>File:" + file + ",Error:" + ex.Message);
                    }
                }
            }
        }

        static string[] GetContractFile()
        {
            List<string> files = new List<string>();
            try
            {
                Filter f = new Filter();
                And and = new And();
                //and.Add(new ItemMatcher("SKUID", ComparisonOperatorType.NotEqual, ""));
                and.Add(new ItemMatcher("SFTPState", ComparisonOperatorType.Equal, "未同步"));
                f.Matcher = and;
                string fstr = H3.BizBus.BizStructureUtility.FilterToJson(f);
                string result = string.Empty;
                result = bizService.LoadBizObjects("D000365MainPushContract", fstr);
                List<Dictionary<string, string>> contracts = ConvertToDic(result);
                if (contracts != null && contracts.Count > 0)
                {
                    string fileName = "DD" + DateTime.Now.ToString("yyyyMMddhhmmss") + "Signed";
                    DataTable dt = new DataTable();
                    foreach (string key in SFTPMappingColumn.Keys)
                    {
                        dt.Columns.Add(SFTPMappingColumn[key]);
                    }
                    dt.Columns.Add("合同照片");
                    DataRow header = dt.NewRow();
                    foreach(DataColumn column in dt.Columns)
                    {
                        header[column.ColumnName] = column.ColumnName;
                    }
                    dt.Rows.Add(header);
                    foreach (Dictionary<string, string> d in contracts)
                    {
                        try
                        {
                            DataRow row = dt.NewRow();
                            foreach (string key in SFTPMappingColumn.Keys)
                            {
                                if (d.ContainsKey(key))
                                {
                                    if (key == "ContractLocation")
                                    {
                                        row[SFTPMappingColumn[key]] = GetAddress(d[key]);
                                    }
                                    else
                                    {
                                        row[SFTPMappingColumn[key]] = d[key];
                                    }
                                }
                            }
                            if (d.ContainsKey("FileID"))
                            {
                                if (!string.IsNullOrEmpty(d["FileID"]))
                                {
                                    string[] fileId = d["FileID"].Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                                    row["合同照片"] = GetFileViewUrl(fileId);
                                }
                            }
                            //更新数据为已同步状态
                            object o = new {
                                SFTPState="已同步" 
                            };
                            string upResult= bizService.UpdateBizObject("D000365MainPushContract", d["ObjectId"], Newtonsoft.Json.JsonConvert.SerializeObject(o));
                            if(!string.IsNullOrEmpty(upResult)&&upResult.IndexOf("成功")>-1)
                            {
                                dt.Rows.Add(row);
                            }
                            else
                            {
                                WriteLog("更新合同同步状态失败---->ObjectId:" + (d.ContainsKey("ObjectId") ? d["ObjectId"] : "") + ",ContractCode:" + (d.ContainsKey("ContractCode") ? d["ContractCode"] : "") + ",Error:" + upResult);
                            }
                        }
                        catch(Exception ex)
                        {
                            WriteLog("获取合同签订信息失败---->ObjectId:" + (d.ContainsKey("ObjectId") ? d["ObjectId"] : "") + ",ContractCode:" + (d.ContainsKey("ContractCode") ? d["ContractCode"] : "") + ",Error:" + ex.Message);
                        }
                    }
                    DataSet st = new DataSet("主推合同");
                    //dt.Rows.InsertAt(header, 0);
                    st.Tables.Add(dt);
                    ImportSource.IImportSource source = ImportSource.ImportSourceFactory.Create(st);

                    ImportData.IImportData builder = ImportData.ImportBuilder.Create(source, ImportData.DataType.Csv);

                    if (!Directory.Exists(PutFilePath))
                    {
                        Directory.CreateDirectory(PutFilePath);
                    }
                    NPOIHelper.ConvertTableToCSV(builder, PutFilePath, fileName + ".csv");
                    files.Add(PutFilePath + fileName + ".csv");
                }
            }
            catch(Exception ex)
            {
                WriteLog("获取合同文件失败-------->Error:" + ex.Message);
            }
            return files.ToArray();
        }


        static string GetFileViewUrl(string[] fileid)
        {
            return string.Empty;
        }

        static string GetAddress(string address)
        {
            if(!string.IsNullOrEmpty(address))
            {
                Dictionary<string, object> d = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(address);
                if (d != null && d.Count > 0 && d.ContainsKey("Address"))
                {
                    string add = d["Address"] + string.Empty;
                    add = ReplaceSpecilWord(add);
                    return add;
                }
            }
            return string.Empty;
        }

        private static string ReplaceSpecilWord(string theString)
        {
            if (string.IsNullOrEmpty(theString)) return theString;
            theString = theString.Replace(">", "&gt;");
            theString = theString.Replace("<", "&lt;");
            theString = theString.Replace(" ", "&nbsp;");
            theString = theString.Replace("\"", "&quot;");
            theString = theString.Replace("\'", "&#39;");
            theString = theString.Replace("\\", "\\\\");//对斜线的转义  
            theString = theString.Replace("\n", "\\n");
            theString = theString.Replace("\r", "\\r");
            theString= theString.Replace(",", "，");
            return theString;
        }

        #endregion

        /// <summary>
        /// 写入日志方法
        /// </summary>
        /// <param name="Message">日志信息</param>
        public static void WriteLog(string Message)
        {
            if (string.IsNullOrEmpty(Message)) return;
            //日志保存路径，不包括文件名
            string filePath = System.AppDomain.CurrentDomain.BaseDirectory + "log";
            //日志完整路径，包括文件名
            string logFileName = DateTime.Now.ToString("yyyy-MM-dd") + ".log";

            WriteLog(filePath, logFileName, Message);
            Console.WriteLine("【" + DateTime.Now.ToString() + "】    " + Message);
        }

        private static void WriteLog(string Path, string FileName, string Message)
        {
            if (string.IsNullOrEmpty(Message)) return;
            //日志保存路径，不包括文件名
            string filePath = Path;
            //日志完整路径，包括文件名
            string logFileName = Path + "\\" + FileName;

            //文件不存在，则创建新文件
            if (!Directory.Exists(filePath))
            {
                try
                {
                    //按照路径创建目录
                    Directory.CreateDirectory(filePath);
                }
                catch (System.Exception e)
                {
                    throw new System.Exception(e + "创建目录失败！");
                }
            }
            if (!File.Exists(logFileName))
            {
                FileStream filestream = null;
                try
                {
                    filestream = File.Create(logFileName);
                    /*创建日志头部*/
                    filestream.Dispose();
                    filestream.Close();
                }
                catch (System.Exception ex)
                {
                    throw new System.Exception(ex + "创建日志文件失败");
                }
            }
            //true 如果日志文件存在则继续追加日志 
            System.IO.StreamWriter sw = null;
            try
            {
                sw = new System.IO.StreamWriter(logFileName, true, System.Text.Encoding.UTF8);
                sw.WriteLine("【" + System.DateTime.Now.ToString() + "】" + "【" + Message + "】");
                //return true;
            }
            catch (System.Exception ex)
            {
                //return false;
                throw new System.Exception(ex + "写入日志失败，检查！");
            }
            finally
            {
                sw.Flush();
                sw.Dispose();
                sw.Close();
            }
        }
    }
}
