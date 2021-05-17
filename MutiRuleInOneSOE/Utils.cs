using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml.Linq;

using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;

namespace MultiRuleInOneSOE
{
    class Utils
    {
        /// <summary>
        /// 数据存储位置
        /// </summary>
        public static string workSpace;
        public static string wordDir;
        public static string fileGDB = null;
        public static string name_MXD = null;

        public static string GHXFileGDB = null;
        public static string GHX_MXD = null;

        public static List<int> running = new List<int>();
        
        //土地类型代码与中文名称对照表
        public static List<Dictionary<string, string>> tdlxDicList = new List<Dictionary<string, string>>();
        //图层序号与图层子类属性字段名对照表
        public static  List<Dictionary<string, string>> layerFieldNameDicList = new  List<Dictionary<string, string>>();
        /// <summary>
        /// 空间参考
        /// </summary>
        public static string refStr = "";
        /// <summary>
        /// 检查范围的外包矩形
        /// </summary>
        public static IEnvelope pfwEnv = null;

        /// <summary>
        /// 图片数组
        /// </summary>
        public static string picStr = null;
        /// <summary>
        /// 服务包含的所有图层
        /// </summary>
        public static List<IFeatureClass> featureClsArray = null;
       // 图层类型
        public static  List<string> layerTypeArray=null ;

        /// <summary>
        /// 读取配置文件，确保目录权限
        /// </summary>
        /// <param name="xmlFilename"></param>
        /// <returns></returns>
        
        public static void getConfigDatas(string xmlFilePath)
        {
            if (System.IO.File.Exists(xmlFilePath))
            {
                string tdlx = "";
                XElement xElement = XElement.Load(xmlFilePath);

                workSpace = xElement.Elements("WorkSpace").Elements("RootDir").First().Value;
                wordDir = xElement.Elements("WorkSpace").Elements("WordDir").First().Value;
                //一张图合规检测
                Utils.fileGDB = workSpace + @"planDatas.gdb";
                Utils.name_MXD = "planTheme";

                //规划线数据库与出图模板
                Utils.GHXFileGDB = workSpace + @"GHJCDatas.gdb";
                Utils.GHX_MXD = "GHXtemplate";
                int count = xElement.Elements("tdlx").Elements().Count(); ;

                tdlxDicList.Clear();
                layerFieldNameDicList.Clear();

                for (int i = 0; i < count; i++)
                {
                    Dictionary<string, string> tdlxDic = new Dictionary<string, string>();
                    if (xElement.Elements("tdlx").Elements("L"+i.ToString()).Any())
                    {
                        tdlx = xElement.Elements("tdlx").Elements("L" + i.ToString()).First().Value.Replace("\n", "").Replace("\t", "").Replace(" ", "");

                        if (!String.IsNullOrEmpty(tdlx))
                        {
                            if (tdlx.EndsWith(","))
                            {
                                tdlx = tdlx.Remove(tdlx.Length - 1);
                            }
                            string[] lxs = tdlx.Split(',');
                            foreach (string item in lxs)
                            {
                                string[] lx = item.Split('-');

                                tdlxDic.Add(lx[0], lx[1]);
                            }
                        }
                           
                        tdlxDicList.Add(tdlxDic);
                    }
                }
                   
                string LayerFieldName = "";
                if (xElement.Elements("LayerFieldName").Any())
                {
                    LayerFieldName = xElement.Elements("LayerFieldName").First().Value.Replace("\n", "").Replace("\t", "").Replace(" ", "");
                    if (LayerFieldName.EndsWith(","))
                    {
                        LayerFieldName = LayerFieldName.Remove(LayerFieldName.Length - 1);
                    }
                    string[] layer = LayerFieldName.Split(',');

                    foreach (string name in layer)
                    {
                        Dictionary<string, string> layerFieldDic = new Dictionary<string, string>();
                        string[] field = name.Split(':');
                        layerFieldDic.Add("layerNo", field[0]);
                        layerFieldDic.Add("layerName", field[1]);
                        layerFieldDic.Add("propName", field[2]);
                        layerFieldDic.Add("type", field[3]);
                        layerFieldNameDicList.Add(layerFieldDic);
                    }

                }

            }
            
        }

        /// <summary>
        /// 获取文件中的数据
        /// </summary>
        /// <param name="args"></param>
        public static string fileToString(String filePath)
        {
            StringBuilder strData = new StringBuilder();
            try
            {
                string line;
                // 创建一个 StreamReader 的实例来读取文件 ,using 语句也能关闭 StreamReader
                using (System.IO.StreamReader sr = new System.IO.StreamReader(filePath))
                {
                    // 从文件读取并显示行，直到文件的末尾
                    while ((line = sr.ReadLine()) != null)
                    {
                        //Console.WriteLine(line);
                        strData.Append(line);
                    }
                }
            }
            catch (Exception e)
            {
                // 向用户显示出错消息
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
            return strData.ToString();
        }
        /// <summary>
        /// 获取土地类型代码对应的中文名称
        /// </summary>
        /// <param name="keyCode"></param>
        /// <returns></returns>
        public static string getKeyType(string keyCode,int layerId)
        {
            string value=keyCode;
            foreach (string key in Utils.tdlxDicList[layerId].Keys)
            {
                if (key == keyCode)
                {
                    Utils.tdlxDicList[layerId].TryGetValue(key, out value);
                }
            }

            return value;
        }
    }
}
