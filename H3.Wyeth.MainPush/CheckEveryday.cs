using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using H3.SHHS.Excel;
using System.Data;
using Newtonsoft.Json;
using H3.BizBus;
using System.Globalization;

namespace H3.Wyeth.MainPush
{
    public class CheckEveryday
    {
        private static com.h3yun.www.BizObjectService _bizService = null;
        public static com.h3yun.www.BizObjectService bizService
        {
            get
            {
                if (_bizService == null)
                {
                    com.h3yun.www.Authentication au = new com.h3yun.www.Authentication();
                    au.Secret = "RgAZzBSv6P47q636oiKG6aYXm+5v4pbe/HPGmDiDIizpLJFKQeVy1Q==";
                    au.EngineCode = "vdgh9dwtgyyj88x64swnm5vt6";
                    _bizService = new com.h3yun.www.BizObjectService();
                    _bizService.AuthenticationValue = au;
                    _bizService.Timeout = int.MaxValue;
                }
                return _bizService;
            }
        }

        public static string FilePath
        {
            get
            {
                return System.Configuration.ConfigurationManager.AppSettings["check"];
            }
        }

        static void Main(string[] args)
        {
            //DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            //dtFormat.ShortDatePattern = "yyyy-MM-dd hh:mm:ss";
            //DateTime dt = Convert.ToDateTime("2018-6-30  23:18");
            //var list= LoadObject();
            try
            {
                BatInsert();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {

            }

        }
        public static void BatInsert()
        {
            DataTable dt = NPOIHelper.ExcelSheetImportToDataTable(FilePath, "每日检查表模拟");
            Dictionary<string, object> dic = null;
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
            if (dt == null && dt.Rows.Count <= 0)
            {
                Console.WriteLine("原数据出错或读取为空");
                return;
            }
            dt.Rows.RemoveAt(0);
            //var rows= dt.Select("F0000001 = '01-6月-2018' ");
            //dt.Rows.AsQueryable()

            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName == "F0000016") col.ColumnName = "F00000_1";
                if (col.ColumnName == "F0000034") col.ColumnName = "F0000016";//旅店名称

                if (col.ColumnName == "F0000015") col.ColumnName = "F00000_2";//人员
                if (col.ColumnName == "F0000033") col.ColumnName = "F0000015";
            }
            List<Dictionary<string, string>> _list = LoadObject();
            foreach (DataRow row in dt.Rows)
            {
                dic = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                {
                    if (col.ColumnName == "F0000001")//日期
                    {
                        DateTimeFormatInfo dtFormat = new System.Globalization.DateTimeFormatInfo();
                        dtFormat.ShortDatePattern = "dd/MM/yyyy";
                        if (string.IsNullOrWhiteSpace(row[col].ToString())) continue;
                        row[col] = row[col].ToString().Replace("6月", "06");
                        try
                        {
                            dic.Add(col.ColumnName, Convert.ToDateTime(row[col], dtFormat));
                        }
                        catch (Exception ex)
                        {

                            throw;
                        }

                    }
                    else if (col.ColumnName == "F0000013")//上报时间
                    {
                        //if (row[col] != new { })
                        //{
                        //    row[col] = Convert.ToDateTime(row[col]);
                        //}
                        dic.Add(col.ColumnName, row[col]);
                    }
                    else if (col.ColumnName == "F0000002")//选择人员旅店
                    {
                        var bizs = _list.Where(P => P["F0000010"].ToString().Contains(row["F0000016"].ToString()) ||     //按旅店名称和创建人筛选
                            row["F0000016"].ToString().Contains(P["F0000010"].ToString()));//
                        if (bizs != null && bizs.Any())
                        {
                            //if(bizs.Count()>1)
                            //{
                            //    var biz = bizs.Where(P => P["F0000003"].ToString().Contains(row["CreatedBy"].ToString())
                            //       || row["CreatedBy"].ToString().Contains(P["F0000003"].ToString().ToString()))
                            //       .FirstOrDefault()??bizs.FirstOrDefault();

                            //}
                            var biz = bizs.FirstOrDefault();
                            if (biz.ContainsKey("F0000001"))
                            {
                                row[col] = biz["ObjectId"];
                                row["F0000015"] = biz["F0000001"];//旅店id
                            }
                        }
                        dic.Add(col.ColumnName, row[col]);
                    }
                    else
                    {
                        dic.Add(col.ColumnName, row[col]);
                    }
                }
                list.Add(dic);
            }
            List<string> ids = new List<string>();
            foreach (var item in list)
            {
                //string result = bizService.CreateBizObject("D000765daycheck", JsonConvert.SerializeObject(item), true);
                //ids.Add(result);
                //Console.WriteLine(result);
            }
            Console.WriteLine("导入完成");
            //Program.WriteLog(JsonConvert.SerializeObject(list));
            Console.ReadKey();
        }
        public static void BatDelete()
        {

        }
        public static List<Dictionary<string, string>> LoadObject()
        {
            Filter f = new Filter();
            And and = new And();
            //and.Add(new ItemMatcher("SKUID", ComparisonOperatorType.NotEqual, ""));
            //and.Add(new ItemMatcher("", ComparisonOperatorType.Equal, ""));
            f.Matcher = and;
            f.FromRowNum = 0;
            f.ToRowNum = 1000;
            string fstr = H3.BizBus.BizStructureUtility.FilterToJson(f);
            string result = string.Empty;
            result = bizService.LoadBizObjects("D000765hotelemp", fstr);

            return Program.ConvertToDic(result);
        }
        public static Dictionary<string, string> GetAllColumnFields
        {
            get
            {
                return new Dictionary<string, string>(){
                    {"CreatedBy","上报人ID"           },
                    {"CreatedBy_Name","上报人"        },
                    {"CreatedTime","创建时间"         },
                    {"F0000001","日期"                },
                    {"F0000002","选择旅店ID"          },
                    {"F0000002_Name","选择旅店"       },
                    {"F0000004", "旅馆登记系统"       },
                    {"F0000005", "旅客信息及时发送"   },
                    {"F0000006", "监控系统是否正常"   },
                    {"F0000007", "上网认证系统"       },
                    {"F0000008", "110报警系统"        },
                    {"F0000009", "消防设施"           },
                    {"F0000010", "禁毒标示标牌"       },
                    {"F0000011", "旅客纸质登记"       },
                    {"F0000012", "旅馆24小时值班制度" },
                    {"F0000013", "上报时间"           },
                    {"F0000014", "分值"               },
                    {"F0000015", "旅店ID"             },
                    {"F0000015_Name", "旅店"          },
                    {"F0000016","旅店名称"            },
                    {"F0000018","登记系统不正常说明"  },
                    {"F0000020","旅客信息未发送说明"  },
                    {"F0000022","监控系统不正常说明"  },
                    {"F0000024","上网认证系统不正说明"},
                    {"F0000026","110不正常说明"       },
                    {"F0000028","消防设施未齐全说明"  },
                    {"F0000030","禁毒标牌不齐全说明"  },
                    {"F0000032","旅客纸质未登记说明"  },
                    {"F0000034","值班制度未到位说明"  },
                    {"ModifiedTime", "修改时间"       },
                    {"Name",  "数据标题"              },
                    {"OwnerDeptId_Name",  "钉钉部门"  },
                    {"OwnerId", "拥有者ID"            },
                    {"OwnerId_Name","拥有者"          },
                };
            }
        }
        public static string ReturnDicKey(string value)
        {
            string field = null;
            var dic = CheckEveryday.GetAllColumnFields;
            foreach (string key in CheckEveryday.GetAllColumnFields.Keys)
            {
                if (dic[key] == value)
                {
                    field = key;
                    break;
                }
            }
            return field;
        }
    }

}
