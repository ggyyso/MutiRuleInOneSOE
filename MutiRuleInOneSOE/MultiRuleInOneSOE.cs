using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;

using System.Collections.Specialized;

using System.Runtime.InteropServices;

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Server;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.SOESupport;
using ESRI.ArcGIS.Output;
using ESRI.ArcGIS.Display;

using Aspose.Words;
using Aspose.Words.Tables;
using System.Collections;

namespace MultiRuleInOneSOE
{
    [ComVisible(true)]
    [Guid("5ac335c3-4ab8-4227-9e17-0137ff25f915")]
    [ClassInterface(ClassInterfaceType.None)]
    [ServerObjectExtension("MapServer",//use "MapServer" if SOE extends a Map service and "ImageServer" if it extends an Image service.
        AllCapabilities = "针对多规合一项目控制线检测功能ArcGIS服务扩展",
        DefaultCapabilities = "空间关系运算、要素相关操作、项目与控制线线检测结果、检测报告Word/pdf生成",
        Description = "多规合一项目控制线检测功能ArcGIS服务扩展",
        DisplayName = "MultiRuleInOneSOE",
        Properties = "Jason Wong修改20210105",
        SupportsREST = true,
        SupportsSOAP = false)]
    public class MultiRuleInOneSOE : IServerObjectExtension, IObjectConstruct, IRESTRequestHandler
    {
        private string soe_name;

        private IPropertySet configProps;
        private IServerObjectHelper serverObjectHelper;
        private ServerLogger logger;
        private IRESTRequestHandler reqHandler;

        /// <summary>
        /// word文档输出表格
        /// </summary>
        DataTable tbl = new DataTable();


        /// <summary>
        /// 项目信息
        /// </summary>
        private object[] proInfos;

        /// <summary>
        /// 返回空间参考
        /// </summary>
        private string toSpatialReference;


        public MultiRuleInOneSOE()
        {
            soe_name = this.GetType().Name;
            logger = new ServerLogger();
            reqHandler = new SoeRestImpl(soe_name, CreateRestSchema()) as IRESTRequestHandler;
        }

        #region IServerObjectExtension Members

        public void Init(IServerObjectHelper pSOH)
        {
            //增加调试
            //System.Diagnostics.Debugger.Launch();

            serverObjectHelper = pSOH;

            logger.LogMessage(ServerLogger.msgType.infoStandard, this.soe_name + ".init()", 200, "Initialized " + this.soe_name + " SOE.");

        }

        public void Shutdown()
        {
            this.tbl = null;
            this.proInfos = null;
            Utils.pfwEnv = null;
            Utils.featureClsArray = null;
        }

        #endregion

        #region IObjectConstruct Members

        public void Construct(IPropertySet props)
        {
            configProps = props;

            try
            {
                Utils.featureClsArray = new List<IFeatureClass>();
                Utils.layerTypeArray = new List<string>();
                //读取图层信息
                IMapServer3 pMapServer = serverObjectHelper.ServerObject as IMapServer3;
                string mapName = pMapServer.DefaultMapName;
                IMapLayerInfos layerinfs = pMapServer.GetServerInfo(mapName).MapLayerInfos;
                IMapServerDataAccess dataAccess = pMapServer as IMapServerDataAccess;
                for (int i = 0; i < layerinfs.Count; i++)
                {
                    IDataset pds = dataAccess.GetDataSource(mapName, i) as IDataset;
                    IFeatureClass pfs = pds as IFeatureClass;
                    Utils.featureClsArray.Add(pds as IFeatureClass);
                    IFeatureDataset pfd = pds as IFeatureDataset;

                    Utils.layerTypeArray.Add(pfs.ShapeType.ToString());

                    //读取空间参考，用来创建客户端绘制的查询范围
                    if (i == 0)
                    {
                        int refLengh;
                        IGeoDataset pGeoDataset = pds as IGeoDataset;
                        IESRISpatialReferenceGEN EsriSR = pGeoDataset.SpatialReference as IESRISpatialReferenceGEN;
                        EsriSR.ExportToESRISpatialReference(out Utils.refStr, out refLengh);
                    }

                }
                //防止配置参数临时修改，每次请求时读取配置,此处取消
                //Utils.getConfigDatas(@"C:\soetemp\config.xml");



            }
            catch (System.Exception ex)
            {
                throw new Exception("Load server layer error");
            }

            logger.LogMessage(ServerLogger.msgType.infoStandard, this.soe_name + ".Construct()", 200, "Get Service Infos " + this.soe_name + " SOE.");

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
            RestResource rootRes = new RestResource(soe_name, false, RootResHandler);

            //信息校核
            //layers为多种规划图层
            RestOperation XXJHOper = new RestOperation("XXJHSOE",
                                                      new string[] { "checkFeatures", "layers", "sr" },
                                                      new string[] { "json" },
                                                      XxjhOperHandler);

            rootRes.operations.Add(XXJHOper);

            //多种规划检测 多规合一第三版本 ，分图层，图层中分类型检测并输出报表（含地图）--规划检测
            //layers为多种规划，
            RestOperation GuiHuaCheckOper = new RestOperation("GuiHuaCheckSOE",
                                                      new string[] { "projectInfos", "checkFeatures", "layers", "servicesfields", "sr" },
                                                      new string[] { "json" },
                                                      YZTHgjcOperHandler2);

            rootRes.operations.Add(GuiHuaCheckOper);

            //一张图合规检测， 分图层检测并输出报表//合规检测
            //layers为多种规划图层
            RestOperation ControlLineHcjcOper = new RestOperation("YZTHgjcSOE",
                                                     new string[] { "projectInfos", "checkFeatures", "layers", "servicesfields", "sr" },
                                                     new string[] { "json" },
                                                     YZTHgjcOperHandler2);
            rootRes.operations.Add(ControlLineHcjcOper);

            return rootRes;
        }

        private byte[] RootResHandler(NameValueCollection boundVariables, string outputFormat, string requestProperties, out string responseProperties)
        {
            responseProperties = null;

            JsonObject result = new JsonObject();
            result.AddString("功能", "一张图合规检测服务扩展");

            return Encoding.UTF8.GetBytes(result.ToJson());
        }


        /// <summary>
        /// 各种规划数据占用，信息校核
        /// </summary>
        /// <param name="boundVariables"></param>
        /// <param name="operationInput"></param>
        /// <param name="outputFormat"></param>
        /// <param name="requestProperties"></param>
        /// <param name="responseProperties"></param>
        /// <returns></returns>
        private byte[] XxjhOperHandler(NameValueCollection boundVariables,
                                                  JsonObject operationInput,
                                                      string outputFormat,
                                                      string requestProperties,
                                                  out string responseProperties)
        {
            responseProperties = null;

            //读取配置文件
            List<ydxzClass> ydxzList = new List<ydxzClass>();

            Utils.getConfigDatas(@"C:\soetemp\config.xml");


            //输入检测要素
            JsonObject jsonCheckInfos;
            bool found = operationInput.TryGetJsonObject("checkFeatures", out jsonCheckInfos);
            if (!found || jsonCheckInfos == null)
                throw new ArgumentNullException("checkFeatures");

            //输入检测图层号,不输入时，默认检测全部
            object layersValue;
            found = operationInput.TryGetObject("layers", out layersValue);
            if (!found || layersValue == null)
            {
                string layerString = "";
                foreach (Dictionary<string, string> d in Utils.layerFieldNameDicList)
                {
                    string layerid = "";
                    d.TryGetValue("layerNo", out layerid);
                    layerString += layerid + ",";
                }
                layersValue = layerString.Substring(0, layerString.Length - 1);
            }


            //获取输入面要素
            object obj = null;
            jsonCheckInfos.TryGetObject("polygons", out obj);

            object[] featuresObj = (object[])obj;
            //输入返回数据空间参考
            object spatialRef;
            found = operationInput.TryGetObject("sr", out spatialRef);
            if (!found)
            {
                //throw new ArgumentNullException("sr");
                toSpatialReference = "4490";
            }
            else
            {
                toSpatialReference = spatialRef.ToString();
            }

            //图层集合
            string[] layers = layersValue.ToString().Split(',');

            //结果集合
            List<JsonObject> resultArray = new List<JsonObject>();




            List<SortedDictionary<string, object>> paramDicList = new List<SortedDictionary<string, object>>();

            List<SortedDictionary<string, object>> dikuaiAreaList = new List<SortedDictionary<string, object>>();

            for (int i = 0; i < featuresObj.Length; i++)
            {
                //计算输入面要素面积
                double prjArea = 0;
                SortedDictionary<string, object> geoDic = new SortedDictionary<string, object>();
                SortedDictionary<string, object> dikuaiArea = new SortedDictionary<string, object>();

                JsonObject fsJson = (JsonObject)featuresObj[i];
                JsonObject geo;
                JsonObject attri;
                string attName;
                fsJson.TryGetJsonObject("geometry", out geo);
                fsJson.TryGetJsonObject("attributes", out attri);
                attri.TryGetString("name",out attName);
                IPolygon input_pg = Conversion.ToGeometry(geo, esriGeometryType.esriGeometryPolygon) as IPolygon;
                input_pg.Close();
                ITopologicalOperator pBoundaryTop = input_pg as ITopologicalOperator;
                pBoundaryTop.Simplify();

                //所以，在对Geometry进行处理之前，增加了如下代码：
                IGeometry prjGeo= SpatialOperations.projectGeometry(input_pg, 4490);
                geoDic.Add("geo", prjGeo);
                geoDic.Add("name", attName); 
                if (input_pg.SpatialReference is ProjectedCoordinateSystem)
                {
                    IArea pArea = input_pg as IArea;
                    prjArea += pArea.Area;
                }
                else
                {
                    int wkid = 4545;
                    if (input_pg.Envelope.Envelope.MMax < 106.5)
                    {
                        wkid = 4544;
                    }
                    IClone pgeoColone = input_pg as IClone;
                    IGeometry input_pg_pro = SpatialOperations.projectGeometry(pgeoColone.Clone() as IGeometry, wkid);
                    IArea pArea = input_pg_pro as IArea;
                    prjArea += pArea.Area;
                }
                //geoDic.Add("area", prjArea);
               
                paramDicList.Add(geoDic);
                dikuaiArea.Add("name", attName);
                dikuaiArea.Add("area", prjArea);
                dikuaiAreaList.Add(dikuaiArea);
            }

            //返回结果
            SortedDictionary<string, resultClass> returnDic = new SortedDictionary<string, resultClass>();

            Utils.picStr = "";

            tbl.Rows.Clear();


            List<SortedDictionary<string, SortedDictionary<string, double>>> LayerStaticDictList = new List<SortedDictionary<string, SortedDictionary<string, double>>>();
            //遍历图层
            for (int i = 0; i < layers.Length; i++)
            {
                // IFeatureClass outFS = null;
                string fieldName = "";
                Utils.layerFieldNameDicList[int.Parse(layers[i])].TryGetValue("propName", out fieldName);
                //outFS = SpatialOperations.spatialIntersect(Utils.fileGDB, inputFeatureClass, int.Parse(layers[i]));
                IFeatureClass kzxFs = Utils.featureClsArray[int.Parse(layers[i])];
               
                ISpatialFilter spatialFilter = new SpatialFilterClass();
                spatialFilter.GeometryField = kzxFs.ShapeFieldName;
                spatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
                //遍历地块,进行空间检索，将检索结果按地块存储。
                SortedDictionary<string, List<IFeature>> searchFeatureDic = new SortedDictionary<string, List<IFeature>>();
                foreach (SortedDictionary<string, object> geoDic in paramDicList)
                {
                    List<IFeature> searchFeatureList = new List<IFeature>();
                    List<int> featureId = new List<int>();

                    object GeoObj;
                    geoDic.TryGetValue("geo", out GeoObj);
                    spatialFilter.Geometry = GeoObj as IGeometry;
                    IFeatureCursor featureCursor = kzxFs.Search(spatialFilter, false);
                    IFeature searchF = null;
                    while ((searchF = featureCursor.NextFeature()) != null)
                    {
                        if (featureId.Contains(searchF.OID))
                        {
                            continue;
                        }
                        featureId.Add(searchF.OID);
                        searchFeatureList.Add(searchF);
                    }
                    Marshal.ReleaseComObject(featureCursor);
                    object dikuaiName;
                    geoDic.TryGetValue("name",out dikuaiName);
                    searchFeatureDic.Add(dikuaiName.ToString(), searchFeatureList);
                }


                int fieldIndex = kzxFs.FindField(fieldName);

                double dArea = 0;
   
                List<JsonObject> geosArray = new List<JsonObject>();

                string fieldValue = "";

                SortedDictionary<string, SortedDictionary<string, double>> layerDikuaiDic = new SortedDictionary<string, SortedDictionary<string, double>>();

                //对检索到的所有空间要素与所有地块进行空间求交
                foreach (string keyDikuaiName in searchFeatureDic.Keys )
                {
                    //不同类型地类面积统计
                    SortedDictionary<string, double> dileiDic = new SortedDictionary<string, double>();
                    
                    List<IFeature> featureList = null;
                    searchFeatureDic.TryGetValue(keyDikuaiName, out featureList);

                    IGeometry dikuaiGeo = GeitDiKuaiGeo(keyDikuaiName, paramDicList);
                    if (null == dikuaiGeo)
                    {
                        continue;
                    }
                    List<IGeometry> diKuaiGeoList=new List<IGeometry>();
                    diKuaiGeoList.Add(dikuaiGeo);
                    //对检索到的要素与单个地块求交
                    foreach (IFeature pfeature in featureList)
                    {
                        //每个要素于 与当前地块求交集
                        List<IGeometry> intersectGeoList = SpatialOperations.spatialIntersect(pfeature.ShapeCopy, diKuaiGeoList);
                        if (intersectGeoList.Count==0)
                        {
                            continue;
                        }

                        foreach (IGeometry intersectGeo in intersectGeoList)
                        {
                            if (intersectGeo.GeometryType == esriGeometryType.esriGeometryPolygon)
                            {

                                //JsonObject featureJson;

                                //if (toSpatialReference.Equals("4326") || toSpatialReference.Equals("4490"))
                                //{
                                //    featureJson = Conversion.ToJsonObject(SpatialOperations.projectGeometry(intersectGeo, 4326));
                                //}
                                //else if (toSpatialReference.Equals("4545"))
                                //{
                                //    featureJson = Conversion.ToJsonObject(SpatialOperations.projectGeometry(intersectGeo, 4545));
                                //}
                                //else
                                //{
                                //    featureJson = Conversion.ToJsonObject(intersectGeo);
                                //}
                                //子类型属性值
                                if (-1 != fieldIndex)
                                {
                                    fieldValue = pfeature.get_Value(fieldIndex).ToString();
                                    fieldValue = Utils.getKeyType(fieldValue, int.Parse(layers[i]));
                                }
                                else
                                {//不分类型时，直接取配置的名称
                                    fieldValue = fieldName;
                                }


                                IArea pArea;
                                if (intersectGeo.SpatialReference is ProjectedCoordinateSystem)
                                {
                                    pArea = pfeature as IArea;

                                }
                                else
                                {
                                    int wkid = 4545;
                                    if (intersectGeo.Envelope.Envelope.MMax < 106.5)
                                    {
                                        wkid = 4544;
                                    }
                                    IGeometry input_pg_pro = SpatialOperations.projectGeometry(intersectGeo, wkid);
                                    pArea = input_pg_pro as IArea;

                                }

                                dArea = 0;
                                if (dileiDic.ContainsKey(fieldValue))
                                {
                                    dileiDic.TryGetValue(fieldValue, out dArea);
                                    dileiDic.Remove(fieldValue);
                                    dileiDic.Add(fieldValue, dArea + pArea.Area);
                                }
                                else
                                {
                                    dileiDic.Add(fieldValue, pArea.Area);

                                }

                            }//end if
                        }

                    }//end for结束当前地块求交
                    //记录当前地块与各类地类类型相交的面积
                    layerDikuaiDic.Add(keyDikuaiName, dileiDic);
                }
                LayerStaticDictList.Add(layerDikuaiDic);

            }//endfor结束图层遍历


            JsonObject result = new JsonObject();
            result.AddArray("Results", LayerStaticDictList.ToArray());

            result.AddObject("RegionArea", dikuaiAreaList);

            return Encoding.UTF8.GetBytes(result.ToJson());
        }

        private IGeometry GeitDiKuaiGeo(string keyDikuaiName, List<SortedDictionary<string, object>> paramDicList)
        {
            foreach (SortedDictionary<string, object> paramDic in paramDicList)
            {
                object dikuaiName;
                object dikuaiGeo;

                paramDic.TryGetValue("name",out dikuaiName);
                if (keyDikuaiName.Equals(dikuaiName))
                {
                    paramDic.TryGetValue("geo", out dikuaiGeo);
                    return dikuaiGeo as IGeometry;
                }
            }
            return null;
        }

        /// <summary>
        /// 一张图控制线合规检测接口
        /// </summary>
        /// <param name="boundVariables"></param>
        /// <param name="operationInput"></param>
        /// <param name="outputFormat"></param>
        /// <param name="requestProperties"></param>
        /// <param name="responseProperties"></param>
        /// <returns></returns>
        private byte[] ControlLineHgjcHandler(NameValueCollection boundVariables,
                                                 JsonObject operationInput,
                                                     string outputFormat,
                                                     string requestProperties,
                                                 out string responseProperties)
        {
            responseProperties = null;
            //检测结论
            Dictionary<int, string> jcztDic = new Dictionary<int, string>();
            //读取配置文件
            List<ydxzClass> ydxzList = new List<ydxzClass>();

            Utils.getConfigDatas(@"C:\soetemp\config.xml");
            //string configStr = Utils.fileToString(System.IO.Path.Combine(Utils.workSpace, "土地利用合规检测配置.json"));
            //IJSONArray configJsonArray = new JSONArrayClass(); ;
            //configJsonArray.ParseString(configStr);
            //for (int i = 0; i < configJsonArray.Count; i++)
            //{
            //    ydxzClass ydxz = new ydxzClass();

            //    IJSONObject objConfig = null;
            //    configJsonArray.TryGetValueAsObject(i, out objConfig);
            //    string type=null;
            //    string code=null;

            //    objConfig.TryGetValueAsString("type", out type);
            //    objConfig.TryGetValueAsString("code", out code);
            //    ydxz.Type = type;
            //    ydxz.Code = code;

            //    List<jcxClass> kzxJcxList = new List<jcxClass>();

            //    IJSONArray kzxjcArray = null;

            //    objConfig.TryGetValueAsArray("kzxjc", out kzxjcArray);
            //    if (null != kzxjcArray)
            //    {
            //        for (int j = 0; j < kzxjcArray.Count; j++)
            //        {
            //            jcxClass jcx = new jcxClass();
            //            IJSONObject objJc = null;
            //            kzxjcArray.TryGetValueAsObject(j, out objJc);
            //            string jctype = null;
            //            bool condi = false;
            //            objJc.TryGetValueAsString("type", out jctype);
            //            objJc.TryGetValueAsBoolean("condi", out condi);
            //            jcx.Type = jctype;
            //            jcx.Condi = condi;
            //            kzxJcxList.Add(jcx);
            //        }
            //    }

            //    ydxz.Kzxjc = kzxJcxList;


            //    List<jcxClass> kcJcxList = new List<jcxClass>();
            //    IJSONArray kcjcArray = null;

            //    objConfig.TryGetValueAsArray("kcjc", out kcjcArray);
            //    if (null != kcjcArray)
            //    {
            //        for (int j = 0; j < kcjcArray.Count; j++)
            //        {
            //            jcxClass jcx = new jcxClass();
            //            IJSONObject objJc = null;
            //            kcjcArray.TryGetValueAsObject(j, out objJc);
            //            string jctype = null;
            //            bool condi = false;
            //            objJc.TryGetValueAsString("type", out jctype);
            //            objJc.TryGetValueAsBoolean("condi", out condi);
            //            jcx.Type = jctype;
            //            jcx.Condi = condi;
            //            kcJcxList.Add(jcx);
            //        }
            //    }

            //    ydxz.Kchgjc = kcJcxList;
            //}


            //输入项目信息
            JsonObject jsonProInfos;
            bool found = operationInput.TryGetJsonObject("projectInfos", out jsonProInfos);
            if (!found || jsonProInfos == null)
                throw new ArgumentNullException("projectInfos");

            //输入检测要素
            JsonObject jsonCheckInfos;
            found = operationInput.TryGetJsonObject("checkFeatures", out jsonCheckInfos);
            if (!found || jsonCheckInfos == null)
                throw new ArgumentNullException("checkFeatures");

            //输入检测控制线图层号
            object layersValue;
            found = operationInput.TryGetObject("layers", out layersValue);
            if (!found || layersValue == null)
                throw new ArgumentNullException("layers");

            //输入对应检测规则
            //object relationsValue;
            //found = operationInput.TryGetObject("relations", out relationsValue);
            //if (!found || relationsValue == null)
            //    throw new ArgumentNullException("relations");

            //输入对应检测结果输出字段
            //object servicesfieldsValue;
            //found = operationInput.TryGetObject("servicesfields", out servicesfieldsValue);
            //if (!found)
            //    throw new ArgumentNullException("servicesfields");


            //输入返回数据空间参考
            object spatialRef;
            found = operationInput.TryGetObject("sr", out spatialRef);
            if (!found)
                throw new ArgumentNullException("sr");

            object obj = null;
            jsonCheckInfos.TryGetObject("polygons", out obj);

            object[] featuresObj = (object[])obj;

            //获取项目信息
            object pro_obj = null;
            jsonProInfos.TryGetObject("projectInfos", out pro_obj);

            proInfos = (object[])pro_obj;

            //图层集合
            string[] layers = layersValue.ToString().Split(',');
            //字段集合
            //string[] fieldsArr = servicesfieldsValue.ToString().Split(';');

            toSpatialReference = spatialRef.ToString();

            List<JsonObject> resultArray = new List<JsonObject>();


            IFeatureClass inputFeatureClass = FeatureOperations.OpenFeatureClass(Utils.fileGDB, "inputs", "input");

            if (null == inputFeatureClass)
            {
                return null;
            }
            else
            {
                FeatureOperations.DeleteAllFeatureFromFeatureClass(inputFeatureClass);
            }

            double prjArea = 0;
            for (int i = 0; i < featuresObj.Length; i++)
            {
                JsonObject fsJson = (JsonObject)featuresObj[i];
                object[] geos;
                fsJson.TryGetArray("geometries", out geos);

                IPolygon input_pg = Conversion.ToGeometry((JsonObject)geos[0], esriGeometryType.esriGeometryPolygon) as IPolygon;

                IFeature pFeature = inputFeatureClass.CreateFeature();
                pFeature.Shape = input_pg;

                if (i == 0)
                {
                    Utils.pfwEnv = input_pg.Envelope;
                }
                else
                {
                    Utils.pfwEnv.Union(input_pg.Envelope);
                }

                pFeature.Store();

                IPolygon polygon_pro = Conversion.ToGeometry((JsonObject)geos[0], esriGeometryType.esriGeometryPolygon) as IPolygon;

                if (input_pg.SpatialReference is ProjectedCoordinateSystem)
                {
                    IArea pArea = input_pg as IArea;
                    prjArea += pArea.Area;
                }
                else
                {
                    //4545 CGCS2000_3_Degree_GK_CM_108E
                    IGeometry input_pg_pro = SpatialOperations.projectGeometry(polygon_pro, 4545);
                    IArea pArea = input_pg_pro as IArea;
                    prjArea += pArea.Area;
                }

            }
            //删除结果图层
            for (int i = 0; i < layers.Length; i++)
            {
                IFeatureClass kzxFs;


                kzxFs = Utils.featureClsArray[int.Parse(layers[i])];


                IFeatureClass IntersectResultFeatureClass = FeatureOperations.OpenFeatureClass(Utils.fileGDB, "results", kzxFs.AliasName + @"IntersectResult");
                if (IntersectResultFeatureClass != null)
                {
                    FeatureOperations.DeleteAllFeatureFromFeatureClass(IntersectResultFeatureClass);
                }


            }

            //返回结果
            Dictionary<string, resultClass> returnDic = new Dictionary<string, resultClass>();

            Utils.picStr = "";

            tbl.Rows.Clear();
            string[] jlArray = new string[] { "通过", "通过", "通过", "通过" };
            //总结论
            bool totalJl = true;
            for (int i = 0; i < layers.Length; i++)
            {
                IFeatureClass outFS = null;


                outFS = SpatialOperations.spatialIntersect(Utils.fileGDB, inputFeatureClass, int.Parse(layers[i]));

                if (outFS != null)
                {
                    IFeatureCursor pfeatureCursor = outFS.Search(null, false);
                    IFeature pfeature = null;
                    double dArea = 0;
                    List<JsonObject> geosArray = new List<JsonObject>();

                    while (null != (pfeature = pfeatureCursor.NextFeature()))
                    {
                        if (pfeature.Shape.GeometryType == esriGeometryType.esriGeometryPolygon)
                        {

                            //JsonObject featureJson;

                            //if (toSpatialReference.Equals("4326"))
                            //{
                            //    featureJson = Conversion.ToJsonObject(SpatialOperations.projectGeometry(pfeature.Shape, 4326));
                            //}
                            //else if (toSpatialReference.Equals("4545"))
                            //{
                            //    featureJson = Conversion.ToJsonObject(SpatialOperations.projectGeometry(pfeature.Shape, 4545));
                            //}
                            //else
                            //{
                            //    featureJson = Conversion.ToJsonObject(pfeature.Shape);
                            //}

                            IArea pArea;
                            if (pfeature.Shape.SpatialReference is ProjectedCoordinateSystem)
                            {
                                pArea = pfeature.Shape as IArea;
                                dArea += pArea.Area;
                            }
                            else
                            {
                                IGeometry input_pg_pro = SpatialOperations.projectGeometry(pfeature.Shape, 4545);
                                pArea = input_pg_pro as IArea;
                                dArea += pArea.Area;
                            }


                            //string fields = fieldsArr[i];

                            //string[] field = fields.ToString().Split(',');

                            //for (int k = 0; k < field.Length; k++)
                            //{

                            //    if (field[k] != "" && pfeature.Fields.FindField(field[k]) != -1)
                            //    {
                            //        if (field[k] == "MJ")
                            //        {
                            //            featureJson.AddObject("MJ", Convert.ToDouble(Math.Round(pArea.Area, 2).ToString("0.0")));
                            //        }
                            //        else
                            //        {

                            //            if (pfeature.get_Value(pfeature.Fields.FindField(field[k])) != null)
                            //            {
                            //                featureJson.AddObject(field[k], pfeature.get_Value(pfeature.Fields.FindField(field[k])).ToString());
                            //            }

                            //        }

                            //    }

                            //}

                            //geosArray.Add(featureJson);
                        }//end if
                    }//end while

                    Marshal.ReleaseComObject(pfeatureCursor);
                    dArea = Convert.ToDouble(Math.Round(dArea, 2).ToString("0.0"));
                    resultClass rs = new resultClass();
                    rs.Area = dArea;
                    rs.Result = "通过";
                    //IGeoDataset geoDs=Utils.featureClsArray[int.Parse(layers[i])] as IGeoDataset;
                    //IDataset ds=geoDs as IDataset;
                    //returnDic.Add(ds.Name, rs);
                    // FillDataTable(i, dArea, "侵占");

                    JsonObject ItemObj = new JsonObject();
                    //ItemObj.AddArray("geometries", geosArray.ToArray());
                    int code = 2;
                    ItemObj.AddObject("area", dArea);
                    ItemObj.AddObject("layer", i);
                    ItemObj.AddObject("code", code);
                    if (layers[i] == "0" || layers[i] == "1" || layers[i] == "2")
                    {
                        if (code == 2)
                        {
                            jlArray[0] = "不通过";
                            totalJl = false;
                        }
                    }
                    else if (layers[i] == "3" || layers[i] == "4" || layers[i] == "5")
                    {
                        if (code == 2)
                        {
                            jlArray[1] = "不通过";
                            totalJl = false;
                        }
                    }
                    else
                    {
                        if (code == 2)
                        {
                            jlArray[2] = "不通过";
                            totalJl = false;
                        }
                    }
                    resultArray.Add(ItemObj);

                }
            }
            if (!totalJl)
            {
                jlArray[3] = "不通过";
            }

            string wordName = CreateHgjcWord(proInfos, prjArea, resultArray, jlArray);

            JsonObject fsJson1 = (JsonObject)proInfos[0];
            object obj_Area = null;
            fsJson1.TryGetObject("area", out obj_Area);
            if (obj_Area == null)
            {
                obj_Area = prjArea;
            }
            JsonObject result = new JsonObject();
            result.AddArray("Results", resultArray.ToArray());
            result.AddString("WordName", wordName);
            result.AddDouble("RegionArea", Double.Parse(obj_Area.ToString()));
            result.AddObject("jl", jlArray);


            return Encoding.UTF8.GetBytes(result.ToJson());
        }

        /// <summary>
        /// 一张图多种规划检测第二版（汉中20210303）
        /// </summary>
        /// <param name="boundVariables"></param>
        /// <param name="operationInput"></param>
        /// <param name="outputFormat"></param>
        /// <param name="requestProperties"></param>
        /// <param name="responseProperties"></param>
        /// <returns></returns>
        private byte[] YZTHgjcOperHandler(NameValueCollection boundVariables,
                                                JsonObject operationInput,
                                                    string outputFormat,
                                                    string requestProperties,
                                                out string responseProperties)
        {
            responseProperties = null;
            try
            {
                Utils.getConfigDatas(@"C:\soetemp\config.xml");
                ////判断锁文件是否存在，若存在就等待，
                //while (File.Exists(Utils.workSpace + "lock.lock"))
                //{
                //    System.Threading.Thread.Sleep(5000);
                //}
                ////创建锁文件
                //StreamWriter sw = new StreamWriter(Utils.workSpace + "lock.lock");
                //sw.Close();

                //读取配置文件
                List<ydxzClass> ydxzList = new List<ydxzClass>();

                //输入项目信息
                JsonObject jsonProInfos;
                bool found = operationInput.TryGetJsonObject("projectInfos", out jsonProInfos);
                //if (!found || jsonProInfos == null)
                //    throw new ArgumentNullException("projectInfos");

                //输入检测要素
                JsonObject jsonCheckInfos;
                found = operationInput.TryGetJsonObject("checkFeatures", out jsonCheckInfos);
                if (!found || jsonCheckInfos == null)
                    throw new ArgumentNullException("checkFeatures");

                //输入检测图层号,不输入时，默认检测全部
                object layersValue;
                found = operationInput.TryGetObject("layers", out layersValue);
                if (!found || layersValue == null)
                {
                    string layerString = "";
                    foreach (Dictionary<string, string> d in Utils.layerFieldNameDicList)
                    {
                        string layerid = "";
                        d.TryGetValue("layerNo", out layerid);
                        layerString += layerid + ",";
                    }
                    layersValue = layerString.Substring(0, layerString.Length - 1);
                }
                //throw new ArgumentNullException("layers");
                //图层集合
                string[] layers = layersValue.ToString().Split(',');

                //返回字段名称
                object servicesfieldsValue;
                found = operationInput.TryGetObject("servicesfields", out servicesfieldsValue);


                //输入返回数据空间参考
                object spatialRef;
                found = operationInput.TryGetObject("sr", out spatialRef);
                if (!found)
                {
                    //throw new ArgumentNullException("sr");
                    toSpatialReference = "4490";
                }
                else
                {
                    toSpatialReference = spatialRef.ToString();
                }

                //获取输入面要素
                object obj = null;
                jsonCheckInfos.TryGetObject("polygons", out obj);

                object[] featuresObj = (object[])obj;

                //获取项目信息
                object pro_obj = null;
                jsonProInfos.TryGetObject("projectInfos", out pro_obj);
                if (pro_obj!=null)
                {
                    proInfos = (object[])pro_obj;
                }
               


                //字段集合(允许为空，不为空时要与图层数对应)
                string[] fieldsArr = null;
                if (null != servicesfieldsValue)
                {
                    fieldsArr = servicesfieldsValue.ToString().Split(';');
                }

                //结果集合
                List<JsonObject> resultArray = new List<JsonObject>();


                IFeatureClass inputFeatureClass = FeatureOperations.OpenFeatureClass(Utils.fileGDB, "inputs", "input");

                if (null == inputFeatureClass)
                {
                    return null;
                }
                else
                {
                    FeatureOperations.DeleteAllFeatureFromFeatureClass(inputFeatureClass);
                }
                //计算输入面要素面积
                double prjArea = 0;
                for (int i = 0; i < featuresObj.Length; i++)
                {
                    JsonObject fsJson = (JsonObject)featuresObj[i];
                    object[] geos;
                    fsJson.TryGetArray("geometries", out geos);

                    IPolygon input_pg = Conversion.ToGeometry((JsonObject)geos[0], esriGeometryType.esriGeometryPolygon) as IPolygon;
                    //这样创建的polygon不闭合，而且查找不到外环的个数，执行以下操作闭合面
                    //而且多边形的面积为负数时，该多边形进行空间检索时，结果也很不可预测，应当空间查询到的对象有时候会查询不到。

                    //推测可能是前端传递的Geometry对象存在某种几何错误。

                    //所以，在对Geometry进行处理之前，增加了如下代码：
                    input_pg.Close();
                    ITopologicalOperator pBoundaryTop = input_pg as ITopologicalOperator;
                    pBoundaryTop.Simplify();

                    IFeature pFeature = inputFeatureClass.CreateFeature();
                    IMAware mAware = input_pg as IMAware;
                    mAware.MAware = false;
                    IZAware zAware = input_pg as IZAware;
                    zAware.ZAware = false;
                    pFeature.Shape = input_pg;

                    if (i == 0)
                    {
                        Utils.pfwEnv = input_pg.Envelope;
                    }
                    else
                    {
                        Utils.pfwEnv.Union(input_pg.Envelope);
                    }

                    pFeature.Store();

                    IPolygon polygon_pro = Conversion.ToGeometry((JsonObject)geos[0], esriGeometryType.esriGeometryPolygon) as IPolygon;

                    if (input_pg.SpatialReference is ProjectedCoordinateSystem)
                    {
                        IArea pArea = input_pg as IArea;
                        prjArea += pArea.Area;
                    }
                    else
                    {
                        //4545 CGCS2000_3_Degree_GK_CM_108E
                        int wkid = 4545;
                        if (polygon_pro.Envelope.Envelope.MMax < 106.5)
                        {
                            wkid = 4544;
                        }

                        IGeometry input_pg_pro = SpatialOperations.projectGeometry(polygon_pro, wkid);
                        IArea pArea = input_pg_pro as IArea;
                        prjArea += pArea.Area;



                    }


                }
                prjArea = Math.Abs(prjArea);
                //删除结果图层
                for (int i = 0; i < layers.Length; i++)
                {
                    IFeatureClass kzxFs;

                    kzxFs = Utils.featureClsArray[int.Parse(layers[i])];


                    IFeatureClass IntersectResultFeatureClass = FeatureOperations.OpenFeatureClass(Utils.fileGDB, "results", kzxFs.AliasName + @"IntersectResult");
                    if (IntersectResultFeatureClass != null)
                    {
                        FeatureOperations.DeleteAllFeatureFromFeatureClass(IntersectResultFeatureClass);
                    }


                }

                //返回结果
                Dictionary<string, resultClass> returnDic = new Dictionary<string, resultClass>();

                Utils.picStr = "";

                tbl.Rows.Clear();

                //按规划类型组织返回feature
                Dictionary<int, List<JsonObject>> featureJsonDic = new Dictionary<int, List<JsonObject>>();
                //规划面积统计
                ////土地规划
                //Dictionary<string, double> tdghAreas = new Dictionary<string, double>();
                ////生态规划
                //Dictionary<string, double> stghAreas = new Dictionary<string, double>();
                ////矿产规划
                //Dictionary<string, double> kcghAreas = new Dictionary<string, double>();

                List<Dictionary<string, double>> areaDicList = new List<Dictionary<string, double>>();
                List<Dictionary<string, int>> geoCountDicList = new List<Dictionary<string, int>>();
                string[] picList = new string[Utils.featureClsArray.Count];
                //初始化返回结果
                for (int i = 0; i < Utils.featureClsArray.Count; i++)
                {
                    Dictionary<string, double> areasDic = new Dictionary<string, double>();
                    Dictionary<string, int> geoCountDic = new Dictionary<string, int>();
                    areaDicList.Add(areasDic);
                    geoCountDicList.Add(geoCountDic);
                }


                //遍历图层
                for (int i = 0; i < layers.Length; i++)
                {
                    IFeatureClass outFS = null;
                    string fieldName = "";
                    Utils.layerFieldNameDicList[int.Parse(layers[i])].TryGetValue("propName", out fieldName);
                    outFS = SpatialOperations.spatialIntersect(Utils.fileGDB, inputFeatureClass, int.Parse(layers[i]));
                    //绘制图片
                    string picPath = ExportMap.exportMap(int.Parse(layers[i]),null);
                    picList[int.Parse(layers[i])] = picPath;
                    if (outFS != null)
                    {

                        int fieldIndex = outFS.FindField(fieldName);

                        Dictionary<string, double> areaDic = areaDicList[int.Parse(layers[i])];
                        Dictionary<string, int> geoCountDic = geoCountDicList[int.Parse(layers[i])];
                        double dArea = 0;
                        int geoCount = 0;
                        List<JsonObject> geosArray = new List<JsonObject>();

                        IFeatureCursor pfeatureCursor = outFS.Search(null, true);

                        IFeature pfeature = null;

                        string fieldValue = "";
                        while (null != (pfeature = pfeatureCursor.NextFeature()))
                        {

                            #region  面类型数据
                            if (pfeature.Shape.GeometryType == esriGeometryType.esriGeometryPolygon)
                            {
                                IPolygon4 pFGeo = pfeature.Shape as IPolygon4;
                                IPolygon[] polygons = GeometryOpt.MultiPartToSinglePart(pFGeo);
                                for (int p = 0; p < polygons.Length; p++)
                                {
                                    IPolygon singleGeo = polygons[p];
                                    if (singleGeo.IsEmpty || !singleGeo.IsClosed)
                                    {
                                        continue;
                                    }
                                    //需要将多面分解为单面

                                    JsonObject featureJson;

                                    if (!String.IsNullOrEmpty(toSpatialReference))
                                    {
                                        featureJson = Conversion.ToJsonObject(SpatialOperations.projectGeometry(singleGeo, int.Parse(toSpatialReference)));
                                    }
                                    else
                                    {
                                        featureJson = Conversion.ToJsonObject(singleGeo);
                                    }
                                    //string json = featureJson.ToJson();
                                    //子类型属性值
                                    if (-1 != fieldIndex)
                                    {
                                        fieldValue = pfeature.get_Value(fieldIndex).ToString();
                                        fieldValue = Utils.getKeyType(fieldValue, int.Parse(layers[i]));
                                    }
                                    else
                                    {//不分类型时，直接取配置的名称
                                        fieldValue = fieldName;
                                    }


                                    IArea pArea;
                                    if (singleGeo.SpatialReference is ProjectedCoordinateSystem)
                                    {
                                        pArea = singleGeo as IArea;

                                    }
                                    else
                                    {
                                        int wkid = 4545;
                                        if (singleGeo.Envelope.Envelope.MMax < 106.5)
                                        {
                                            wkid = 4544;
                                        }
                                        IGeometry input_pg_pro = SpatialOperations.projectGeometry(singleGeo, wkid);
                                        pArea = input_pg_pro as IArea;

                                    }
                                    if (pArea.Area < 0.000000000001)
                                    {
                                        continue;
                                    }
                                    dArea = 0;
                                    geoCount = 0;
                                    if (areaDic.ContainsKey(fieldValue))
                                    {
                                        areaDic.TryGetValue(fieldValue, out dArea);
                                        areaDic.Remove(fieldValue);
                                        areaDic.Add(fieldValue, dArea + pArea.Area);

                                        geoCountDic.TryGetValue(fieldValue, out geoCount);
                                        geoCountDic.Remove(fieldValue);
                                        geoCountDic.Add(fieldValue, geoCount + 1);
                                    }
                                    else
                                    {
                                        areaDic.Add(fieldValue, pArea.Area);
                                        geoCountDic.Add(fieldValue, 1);
                                    }

                                    if (null != fieldsArr)
                                    {
                                        string fields = fieldsArr[i];

                                        string[] field = fields.ToString().Split(',');

                                        for (int k = 0; k < field.Length; k++)
                                        {

                                            if (field[k] != "" && pfeature.Fields.FindField(field[k]) != -1)
                                            {
                                                if (field[k] == "MJ")
                                                {
                                                    featureJson.AddObject("MJ", Convert.ToDouble(Math.Round(pArea.Area, 2).ToString("0.00")));
                                                }
                                                else
                                                {

                                                    if (pfeature.get_Value(pfeature.Fields.FindField(field[k])) != null)
                                                    {
                                                        featureJson.AddObject(field[k], pfeature.get_Value(pfeature.Fields.FindField(field[k])).ToString());
                                                    }

                                                }

                                            }

                                        }
                                    }

                                    geosArray.Add(featureJson);
                                }
                            }
                            #endregion

                            #region 点类型

                            else if (pfeature.Shape.GeometryType == esriGeometryType.esriGeometryPoint || pfeature.Shape.GeometryType == esriGeometryType.esriGeometryMultipoint)
                            {
                                List<IPoint> points = new List<IPoint>();
                                if (pfeature.Shape.GeometryType == esriGeometryType.esriGeometryMultipoint)
                                {
                                    points = GeometryOpt.GetPointList(pfeature.Shape);
                                }
                                else
                                {
                                    points.Add(pfeature.Shape as IPoint);
                                }

                                for (int p = 0; p < points.Count; p++)
                                {
                                    IPoint pPoint = points[p];
                                    JsonObject featureJson;

                                    if (!String.IsNullOrEmpty(toSpatialReference))
                                    {
                                        featureJson = Conversion.ToJsonObject(SpatialOperations.projectGeometry(pPoint, int.Parse(toSpatialReference)));
                                    }
                                    else
                                    {
                                        featureJson = Conversion.ToJsonObject(pPoint);
                                    }
                                    //子类型属性值
                                    if (-1 != fieldIndex)
                                    {
                                        object value = pfeature.get_Value(fieldIndex);
                                        if (value == null)
                                        {
                                            value = "其他";
                                        }
                                        else
                                        {
                                            fieldValue = value.ToString();
                                            fieldValue = Utils.getKeyType(fieldValue, int.Parse(layers[i]));
                                        }

                                    }
                                    else
                                    {//不分类型时，直接取配置的名称
                                        fieldValue = fieldName;
                                    }
                                    double pointCount = 0;

                                    if ("".Equals(fieldValue) || fieldValue == null)
                                    {
                                        fieldValue = "其他";
                                    }
                                    dArea = 0;
                                    geoCount = 0;
                                    if (areaDic.ContainsKey(fieldValue))
                                    {
                                        areaDic.TryGetValue(fieldValue, out pointCount);
                                        areaDic.Remove(fieldValue);
                                        areaDic.Add(fieldValue, pointCount + 1);

                                        geoCountDic.TryGetValue(fieldValue, out geoCount);
                                        geoCountDic.Remove(fieldValue);
                                        geoCountDic.Add(fieldValue, geoCount + 1);
                                    }
                                    else
                                    {
                                        areaDic.Add(fieldValue, 1);
                                        geoCountDic.Add(fieldValue, 1);
                                    }

                                    if (null != fieldsArr)
                                    {
                                        string fields = fieldsArr[i];

                                        string[] field = fields.ToString().Split(',');

                                        for (int k = 0; k < field.Length; k++)
                                        {

                                            if (field[k] != "" && pfeature.Fields.FindField(field[k]) != -1)
                                            {
                                                if (pfeature.get_Value(pfeature.Fields.FindField(field[k])) != null)
                                                {
                                                    featureJson.AddObject(field[k], pfeature.get_Value(pfeature.Fields.FindField(field[k])).ToString());
                                                }
                                            }

                                        }
                                    }

                                    geosArray.Add(featureJson);
                                }

                            }//end if
                            #endregion
                            //这块强烈建议feature释放掉，如果做大数据操作的时候这块不释放 内存会撑爆的。
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pfeature);
                        }//end while
                        Marshal.ReleaseComObject(pfeatureCursor);

                        JsonObject layerJson = new JsonObject();
                        layerJson.AddString("layer", layers[i]);
                        layerJson.AddArray("geometrys", geosArray.ToArray());
                        Dictionary<string, double> newAreaDic = new Dictionary<string, double>();
                        Dictionary<string, int> newGeoCountDic = new Dictionary<string, int>();

                        List<JsonObject> statisticList = new List<JsonObject>();
                        foreach (string key in areaDic.Keys)
                        {
                            JsonObject statisticObj = new JsonObject();

                            double areaTem;
                            areaDic.TryGetValue(key, out areaTem);
                            newAreaDic.Add(Utils.getKeyType(key, int.Parse(layers[i])), areaTem);

                            int count1;
                            geoCountDic.TryGetValue(key, out count1);
                            newGeoCountDic.Add(Utils.getKeyType(key, int.Parse(layers[i])), count1);
                            statisticObj.AddString("name", Utils.getKeyType(key, i));
                           // 区分小数和整数
                            if (areaTem.ToString().IndexOf(".")==-1)
                            {
                                statisticObj.AddString("area", areaTem.ToString());
                            }
                            else
                            {
                                statisticObj.AddString("area",  (areaTem / 667.0).ToString("0.00"));
                            }
                            
                            statisticObj.AddLong("count", count1);
                            statisticList.Add(statisticObj);
                        }
                        areaDicList[int.Parse(layers[i])] = newAreaDic;
                        geoCountDicList[int.Parse(layers[i])] = newGeoCountDic;

                        layerJson.AddArray("statistic", statisticList.ToArray());
                        layerJson.AddString("pic",System.IO.Path.GetFileName(picPath));
                        resultArray.Add(layerJson);
                    }

                    

                }//endfor结束图层遍历

                string wordName = CreateYZTHGJCWord2(proInfos, layers, picList, prjArea, areaDicList);

                JsonObject fsJson1 = (JsonObject)proInfos[0];
                object obj_Area = null;
                fsJson1.TryGetObject("area", out obj_Area);
                if (obj_Area == null)
                {
                    obj_Area = prjArea;
                }
                JsonObject result = new JsonObject();
                result.AddArray("Results", resultArray.ToArray());
                result.AddString("WordName", wordName);
                result.AddDouble("RegionArea", Double.Parse(obj_Area.ToString()));

                ////删除锁文件
                //File.Delete(Utils.workSpace + "lock.lock");
                return Encoding.UTF8.GetBytes(result.ToJson());
            }
            catch (System.Exception ex)
            {
                File.Delete(Utils.workSpace + "lock.lock");
                throw ex;
            }

        }
        /// <summary>
        /// 一张图多种规划检测自主实现有问题（汉中20210303）
        /// </summary>
        /// <param name="boundVariables"></param>
        /// <param name="operationInput"></param>
        /// <param name="outputFormat"></param>
        /// <param name="requestProperties"></param>
        /// <param name="responseProperties"></param>
        /// <returns></returns>
        private byte[] YZTHgjcOperHandler2(NameValueCollection boundVariables,
                                                JsonObject operationInput,
                                                    string outputFormat,
                                                    string requestProperties,
                                                out string responseProperties)
        {
            responseProperties = null;

            //读取配置文件
            List<ydxzClass> ydxzList = new List<ydxzClass>();

            Utils.getConfigDatas(@"C:\soetemp\config.xml");


            //输入项目信息
            JsonObject jsonProInfos;
            bool found = operationInput.TryGetJsonObject("projectInfos", out jsonProInfos);
            if (!found || jsonProInfos == null)
                throw new ArgumentNullException("projectInfos");

            //输入检测要素
            JsonObject jsonCheckInfos;
            found = operationInput.TryGetJsonObject("checkFeatures", out jsonCheckInfos);
            if (!found || jsonCheckInfos == null)
                throw new ArgumentNullException("checkFeatures");

            //输入检测图层号,不输入时，默认检测全部
            object layersValue;
            found = operationInput.TryGetObject("layers", out layersValue);
            if (!found || layersValue == null)
            {
                string layerString = "";
                foreach (Dictionary<string, string> d in Utils.layerFieldNameDicList)
                {
                    string layerid = "";
                    d.TryGetValue("layerNo", out layerid);
                    layerString += layerid + ",";
                }
                layersValue = layerString.Substring(0, layerString.Length - 1);
            }
            //throw new ArgumentNullException("layers");

            //返回字段名称
            object servicesfieldsValue;
            found = operationInput.TryGetObject("servicesfields", out servicesfieldsValue);


            //输入返回数据空间参考
            object spatialRef;
            found = operationInput.TryGetObject("sr", out spatialRef);
            if (!found)
                throw new ArgumentNullException("sr");

            //获取输入面要素
            object obj = null;
            jsonCheckInfos.TryGetObject("polygons", out obj);

            object[] featuresObj = (object[])obj;

            //获取项目信息
            object pro_obj = null;
            jsonProInfos.TryGetObject("projectInfos", out pro_obj);

            proInfos = (object[])pro_obj;
            bool isReturnGeometry = true;
            if (null != proInfos)
            {
                JsonObject fsJson = (JsonObject)proInfos[0];
                object isreturn;
                fsJson.TryGetObject("returnGeometry", out isreturn);
                isReturnGeometry = isreturn == null ? true : (bool)isreturn;
            }
            //图层集合
            string[] layers = layersValue.ToString().Split(',');
            //字段集合(允许为空，不为空时要与图层数对应)
            string[] fieldsArr = null;
            if (null != servicesfieldsValue)
            {
                fieldsArr = servicesfieldsValue.ToString().Split(';');
            }


            toSpatialReference = spatialRef.ToString();
            //结果集合
            List<JsonObject> resultArray = new List<JsonObject>();


            //IFeatureClass inputFeatureClass = FeatureOperations.OpenFeatureClass(Utils.fileGDB, "inputs", "input");

            //if (null == inputFeatureClass)
            //{
            //    return null;
            //}
            //else
            //{
            //    FeatureOperations.DeleteAllFeatureFromFeatureClass(inputFeatureClass);
            //}
            //计算输入面要素面积
            double prjArea = 0;
            List<IGeometry> inputGeometryList = new List<IGeometry>();
           
           
            for (int i = 0; i < featuresObj.Length; i++)
            {
                JsonObject fsJson = (JsonObject)featuresObj[i];
                object[] geos;
                fsJson.TryGetArray("geometries", out geos);

                IPolygon input_pg = Conversion.ToGeometry((JsonObject)geos[0], esriGeometryType.esriGeometryPolygon) as IPolygon;
                input_pg.Close();
                ITopologicalOperator pBoundaryTop = input_pg as ITopologicalOperator;
                pBoundaryTop.Simplify();
                //IFeature pFeature = inputFeatureClass.CreateFeature();
                //IMAware mAware = input_pg as IMAware;
                //mAware.MAware = false;
                //IZAware zAware = input_pg as IZAware;
                //zAware.ZAware = false;
                //pFeature.Shape = input_pg;
                //所以，在对Geometry进行处理之前，增加了如下代码：
                IGeometry prjGeo = SpatialOperations.projectGeometry(input_pg, 4490);
                inputGeometryList.Add(prjGeo);

                if (i == 0)
                {
                    Utils.pfwEnv = input_pg.Envelope;
                }
                else
                {
                    Utils.pfwEnv.Union(input_pg.Envelope);
                }

               // pFeature.Store();

               
                if (input_pg.SpatialReference is ProjectedCoordinateSystem)
                {
                    IArea pArea = input_pg as IArea;
                    prjArea += pArea.Area;
                }
                else
                {
                    int wkid = 4545;
                    if (input_pg.Envelope.Envelope.MMax < 106.5)
                    {
                        wkid = 4544;
                    }
                    IClone pgeoColone = input_pg as IClone;
                    IGeometry input_pg_pro = SpatialOperations.projectGeometry(pgeoColone.Clone() as IGeometry, wkid);
                    IArea pArea = input_pg_pro as IArea;
                    prjArea += pArea.Area;
                }
                

            }
            prjArea = Math.Abs(prjArea);
            ////删除结果图层
            //for (int i = 0; i < layers.Length; i++)
            //{
            //    IFeatureClass kzxFs;

            //    kzxFs = Utils.featureClsArray[int.Parse(layers[i])];


            //    IFeatureClass IntersectResultFeatureClass = FeatureOperations.OpenFeatureClass(Utils.fileGDB, "results", kzxFs.AliasName + @"IntersectResult");
            //    if (IntersectResultFeatureClass != null)
            //    {
            //        FeatureOperations.DeleteAllFeatureFromFeatureClass(IntersectResultFeatureClass);
            //    }


            //}

            //返回结果
            Dictionary<string, resultClass> returnDic = new Dictionary<string, resultClass>();

            Utils.picStr = "";

            tbl.Rows.Clear();

            //按规划类型组织返回feature
            Dictionary<int, List<JsonObject>> featureJsonDic = new Dictionary<int, List<JsonObject>>();
            //规划面积统计
            ////土地规划
            //Dictionary<string, double> tdghAreas = new Dictionary<string, double>();
            ////生态规划
            //Dictionary<string, double> stghAreas = new Dictionary<string, double>();
            ////矿产规划
            //Dictionary<string, double> kcghAreas = new Dictionary<string, double>();

            List<Dictionary<string, double>> areaDicList = new List<Dictionary<string, double>>();
            List<Dictionary<string, int>> geoCountDicList = new List<Dictionary<string, int>>();
            string[] picList = new string[Utils.featureClsArray.Count];
            //初始化返回结果
            for (int i = 0; i < Utils.featureClsArray.Count; i++)
            {
                Dictionary<string, double> areasDic = new Dictionary<string, double>();
                Dictionary<string, int> geoCountDic = new Dictionary<string, int>();
                areaDicList.Add(areasDic);
                geoCountDicList.Add(geoCountDic);
            }

            //遍历图层
            for (int i = 0; i < layers.Length; i++)
            {
                // IFeatureClass outFS = null;
                string fieldName = "";
                Utils.layerFieldNameDicList[int.Parse(layers[i])].TryGetValue("propName", out fieldName);
                //出图
                string picPath = ExportMap.exportMap(int.Parse(layers[i]), inputGeometryList);
                picList[int.Parse(layers[i])] = picPath;
                //outFS = SpatialOperations.spatialIntersect(Utils.fileGDB, inputFeatureClass, int.Parse(layers[i]));
                IFeatureClass kzxFs = Utils.featureClsArray[int.Parse(layers[i])];
                List<IFeature> searchFeatureList = new List<IFeature>();
                List<int> featureId = new List<int>();
                ISpatialFilter spatialFilter = new SpatialFilterClass();
                spatialFilter.GeometryField = kzxFs.ShapeFieldName;
                spatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
                foreach (IGeometry geo in inputGeometryList)
                {

                    spatialFilter.Geometry = geo;

                    IFeatureCursor featureCursor = kzxFs.Search(spatialFilter, false);
                    IFeature searchF = null;
                    while ((searchF = featureCursor.NextFeature()) != null)
                    {
                        if (featureId.Contains(searchF.OID))
                        {
                            continue;
                        }
                        featureId.Add(searchF.OID);
                        searchFeatureList.Add(searchF);
                    }
                    Marshal.ReleaseComObject(featureCursor);
                }




                int fieldIndex = kzxFs.FindField(fieldName);

                Dictionary<string, double> areaDic = areaDicList[int.Parse(layers[i])];
                Dictionary<string, int> geoCountDic = geoCountDicList[int.Parse(layers[i])];

                //IFeatureCursor pfeatureCursor = outFS.Search(null, false);

                double dArea = 0;
                int geoCount = 0;
                List<JsonObject> geosArray = new List<JsonObject>();

                string fieldValue = "";
                foreach (IFeature pfeature in searchFeatureList)
                {
                    //inputFeature 与pfeature空间求交集
                    List<IGeometry> intersectGeoList = SpatialOperations.spatialIntersect(pfeature.ShapeCopy, inputGeometryList);
                    foreach (IGeometry intersectGeo in intersectGeoList)
                    {
                        if (intersectGeo.GeometryType == esriGeometryType.esriGeometryPolygon)
                        {

                            JsonObject featureJson;

                            if (!String.IsNullOrEmpty(toSpatialReference))
                            {
                                featureJson = Conversion.ToJsonObject(SpatialOperations.projectGeometry(intersectGeo, int.Parse(toSpatialReference)));
                            }
                            else
                            {
                                featureJson = Conversion.ToJsonObject(intersectGeo);
                            }
                            //子类型属性值
                            if (-1 != fieldIndex)
                            {
                                fieldValue = pfeature.get_Value(fieldIndex).ToString();
                            }
                            else
                            {//不分类型时，直接取配置的名称
                                fieldValue = fieldName;
                            }


                            IArea pArea;
                            if (intersectGeo.SpatialReference is ProjectedCoordinateSystem)
                            {
                                pArea = pfeature as IArea;

                            }
                            else
                            {
                                int wkid = 4545;
                                if (intersectGeo.Envelope.Envelope.MMax < 106.5)
                                {
                                    wkid = 4544;
                                }
                                IGeometry input_pg_pro = SpatialOperations.projectGeometry(intersectGeo, wkid);
                                pArea = input_pg_pro as IArea;

                            }

                            dArea = 0;
                            geoCount = 0;
                            if (areaDic.ContainsKey(fieldValue))
                            {
                                areaDic.TryGetValue(fieldValue, out dArea);
                                areaDic.Remove(fieldValue);
                                areaDic.Add(fieldValue, dArea + pArea.Area);

                                geoCountDic.TryGetValue(fieldValue, out geoCount);
                                geoCountDic.Remove(fieldValue);
                                geoCountDic.Add(fieldValue, geoCount + 1);
                            }
                            else
                            {
                                areaDic.Add(fieldValue, pArea.Area);
                                geoCountDic.Add(fieldValue, 1);
                            }

                            if (null != fieldsArr)
                            {
                                string fields = fieldsArr[i];

                                string[] field = fields.ToString().Split(',');

                                for (int k = 0; k < field.Length; k++)
                                {

                                    if (field[k] != "" && pfeature.Fields.FindField(field[k]) != -1)
                                    {
                                        if (field[k] == "MJ")
                                        {
                                            featureJson.AddObject("MJ", Convert.ToDouble(Math.Round(pArea.Area, 2).ToString("0.0")));
                                        }
                                        else
                                        {

                                            if (pfeature.get_Value(pfeature.Fields.FindField(field[k])) != null)
                                            {
                                                featureJson.AddObject(field[k], pfeature.get_Value(pfeature.Fields.FindField(field[k])).ToString());
                                            }

                                        }

                                    }

                                }
                            }
                            if (isReturnGeometry)
                            {
                                geosArray.Add(featureJson);
                            }
                            
                        }
                        else if (pfeature.Shape.GeometryType == esriGeometryType.esriGeometryPoint)
                        {
                            JsonObject featureJson;

                            if (!String.IsNullOrEmpty(toSpatialReference))
                            {
                                featureJson = Conversion.ToJsonObject(SpatialOperations.projectGeometry(pfeature.Shape, int.Parse(toSpatialReference)));
                            }
                            else
                            {
                                featureJson = Conversion.ToJsonObject(pfeature.Shape);
                            }
                            //子类型属性值
                            if (-1 != fieldIndex)
                            {
                                object value = pfeature.get_Value(fieldIndex);
                                if (value == null)
                                {
                                    value = "其他";
                                }
                                else
                                {
                                    fieldValue = value.ToString();
                                    fieldValue = Utils.getKeyType(fieldValue, int.Parse(layers[i]));
                                }

                            }
                            else
                            {//不分类型时，直接取配置的名称
                                fieldValue = fieldName;
                            }
                            double pointCount = 0;

                            if ("".Equals(fieldValue) || fieldValue == null)
                            {
                                fieldValue = "其他";
                            }
                            dArea = 0;
                            geoCount = 0;
                            if (areaDic.ContainsKey(fieldValue))
                            {
                                areaDic.TryGetValue(fieldValue, out pointCount);
                                areaDic.Remove(fieldValue);
                                areaDic.Add(fieldValue, pointCount + 1);

                                geoCountDic.TryGetValue(fieldValue, out geoCount);
                                geoCountDic.Remove(fieldValue);
                                geoCountDic.Add(fieldValue, geoCount + 1);
                            }
                            else
                            {
                                areaDic.Add(fieldValue, 1);
                                geoCountDic.Add(fieldValue, 1);
                            }

                            if (null != fieldsArr)
                            {
                                string fields = fieldsArr[i];

                                string[] field = fields.ToString().Split(',');

                                for (int k = 0; k < field.Length; k++)
                                {

                                    if (field[k] != "" && pfeature.Fields.FindField(field[k]) != -1)
                                    {
                                        if (pfeature.get_Value(pfeature.Fields.FindField(field[k])) != null)
                                        {
                                            featureJson.AddObject(field[k], pfeature.get_Value(pfeature.Fields.FindField(field[k])).ToString());
                                        }
                                    }

                                }
                            }

                            if (isReturnGeometry)
                            {
                                geosArray.Add(featureJson);
                            }
                        }//end if
                    }

                }//end for

                //Marshal.ReleaseComObject(pfeatureCursor);

                JsonObject layerJson = new JsonObject();
                layerJson.AddString("layer", layers[i]);
                layerJson.AddArray("geometrys", geosArray.ToArray());
                Dictionary<string, double> newAreaDic = new Dictionary<string, double>();
                Dictionary<string, int> newGeoCountDic = new Dictionary<string, int>();

                List<JsonObject> statisticList = new List<JsonObject>();
                foreach (string key in areaDic.Keys)
                {
                    JsonObject statisticObj = new JsonObject();

                    double areaTem;
                    areaDic.TryGetValue(key, out areaTem);
                    newAreaDic.Add(Utils.getKeyType(key, int.Parse(layers[i])), areaTem);

                    int count1;
                    geoCountDic.TryGetValue(key, out count1);
                    newGeoCountDic.Add(Utils.getKeyType(key, int.Parse(layers[i])), count1);
                    statisticObj.AddString("name", Utils.getKeyType(key, i));
                    statisticObj.AddDouble("area", areaTem);
                    statisticObj.AddLong("count", count1);
                    statisticList.Add(statisticObj);
                }
                areaDicList[int.Parse(layers[i])] = newAreaDic;
                geoCountDicList[int.Parse(layers[i])] = newGeoCountDic;

                layerJson.AddArray("statistic", statisticList.ToArray());
                layerJson.AddString("pic", System.IO.Path.GetFileName(picPath));
                resultArray.Add(layerJson);


                

            }//endfor结束图层遍历

            string wordName = CreateYZTHGJCWord(proInfos, layers, picList, prjArea, areaDicList);

            JsonObject fsJson1 = (JsonObject)proInfos[0];
            object obj_Area = null;
            fsJson1.TryGetObject("area", out obj_Area);
            if (obj_Area == null)
            {
                obj_Area = prjArea;
            }
            JsonObject result = new JsonObject();
            result.AddArray("Results", resultArray.ToArray());
            result.AddString("WordName", wordName);
            result.AddDouble("RegionArea", Double.Parse(obj_Area.ToString()));

            return Encoding.UTF8.GetBytes(result.ToJson());
        }

        /// <summary>
        /// 生成合规检测word文档
        /// </summary>
        /// <param name="proInfo">项目信息</param>
        /// <param name="area">检测面积</param>
        /// <param name="resultArray">分项检测结果</param>
        /// <param name="jlArray">综合检测结论</param>
        /// <returns></returns>
        private string CreateHgjcWord(object[] proInfo, double area, List<JsonObject> resultArray, string[] jlArray)
        {
            try
            {
                string wordName = "";

                WordOpr wo = new WordOpr();
                wo.killWinWordProcess();
                string outPath = Utils.workSpace + @"template_hgjc.doc";
                wo.OpenWithTemplate(outPath);
                wo.Builder();

                DocumentBuilder builder = wo.WordBuilder;

                JsonObject fsJson = (JsonObject)proInfo[0];

                object obj_name = null;
                fsJson.TryGetObject("name", out obj_name);

                object obj_type = null;
                fsJson.TryGetObject("type", out obj_type);

                object obj_department = null;
                fsJson.TryGetObject("department", out obj_department);

                object obj_Area = null;

                fsJson.TryGetObject("area", out obj_Area);
                if (obj_Area == null)
                {
                    obj_Area = area;
                }

                string prjName = (String)obj_name;
                //string prjNum = "sn232424234";
                string ydxz = (String)obj_type;
                string department = (String)obj_department;

                builder.MoveToMergeField("xmmc");
                builder.Write(prjName);
                builder.MoveToMergeField("xmdw");
                builder.Write(department);
                builder.MoveToMergeField("ydmj");
                builder.Write(Double.Parse(obj_Area.ToString()).ToString());
                builder.MoveToMergeField("sqydxz");
                builder.Write(ydxz);
                builder.MoveToMergeField("date");
                builder.Write(System.DateTime.Now.ToLongDateString());
                //三线检测结论
                double sxjc = 0;
                //其他控制线检测结论
                double qtkzxjc = 0;
                //专项规划检测结论
                double zxghjc = 0;

                //三线检测表格填写 h0-h7
                for (int i = 0; i < 8; i++)
                {
                    bool flag = false;
                    double? areaT = null;
                    long? code = null;
                    foreach (JsonObject obj in resultArray)
                    {
                        long? layerid = null;
                        obj.TryGetAsLong("layer", out layerid);
                        if (layerid.Value == i)
                        {
                            flag = true;
                        }
                        else
                        {
                            continue;
                        }
                        obj.TryGetAsDouble("area", out areaT);

                        obj.TryGetAsLong("code", out code);
                    }
                    builder.MoveToMergeField("h" + i);
                    if (flag)
                    {

                        builder.Write(Math.Round(areaT.Value, 2).ToString("0.0"));
                    }
                    else
                    {
                        builder.Write("无检测");
                    }
                }

                bool jcjl = true;
                //合规性检查结论
                builder.MoveToMergeField("sxjcjl");
                builder.Write(jlArray[0]);
                builder.MoveToMergeField("qtkzxjcjl");
                builder.Write(jlArray[1]);
                builder.MoveToMergeField("zxghjcjl");
                builder.Write(jlArray[2]);
                builder.MoveToMergeField("zhjl");
                builder.Write(jlArray[3]);
                wordName = System.DateTime.Now.ToFileTime().ToString() + ".doc";
                wo.SaveAs(Utils.wordDir + wordName);
                return wordName;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        /// <summary>
        /// 生成控制线word文档第一版
        /// </summary>
        private string CreateKzxWord(object[] proInfo, string imageFiles, double area)
        {
            try
            {
                string wordName = "";

                WordOpr wo = new WordOpr();
                wo.killWinWordProcess();
                string outPath = Utils.workSpace + @"template_kzx.doc";
                wo.OpenWithTemplate(outPath);
                wo.Builder();

                DocumentBuilder builder = wo.WordBuilder;

                JsonObject fsJson = (JsonObject)proInfo[0];

                object obj_name = null;
                fsJson.TryGetObject("name", out obj_name);

                object obj_type = null;
                fsJson.TryGetObject("type", out obj_type);

                object obj_department = null;
                fsJson.TryGetObject("department", out obj_department);

                string prjName = (String)obj_name;
                string prjNum = "sn232424234";
                string businessType = (String)obj_type;
                string department = (String)obj_department;

                builder.MoveToMergeField("xmmc1");
                builder.Write(prjName);
                builder.MoveToMergeField("xmbh");
                builder.Write(prjNum);
                builder.MoveToMergeField("xmmc2");
                builder.Write(prjName);

                builder.MoveToMergeField("ywlx");
                builder.Write(businessType);
                builder.MoveToMergeField("jsdw");
                builder.Write(department);
                builder.MoveToBookmark("ydfw");
                builder.Write(Convert.ToDouble(Math.Round(area, 2).ToString("0.0")).ToString());
                builder.MoveToMergeField("date");
                builder.Write(System.DateTime.Now.ToLongDateString());


                List<double> widthLst = new List<double>();
                for (int i = 0; i < 4; i++)
                {
                    builder.MoveToCell(1, 0, i, 0);
                    widthLst.Add(builder.CellFormat.Width);
                }
                builder.MoveToDocumentStart();
                builder.MoveToMergeField("jcxq");
                wo.InsertTable2(tbl, true, widthLst);


                builder.MoveToDocumentStart();
                builder.MoveToMergeField("pic");
                string[] picsArr = imageFiles.Split((','));
                for (int i = 0; i < picsArr.Length - 1; i++)
                {
                    builder.InsertImage(picsArr[i]);
                }


                wordName = System.DateTime.Now.ToFileTime().ToString() + ".doc";
                wo.SaveAs(Utils.wordDir + wordName);
                return wordName;
            }
            catch (Exception ex)
            {
                return null;
            }


        }

        /// <summary>
        /// 创建一张图合规检测文档（20210303）第二版
        /// </summary>
        /// <param name="proInfo"></param>
        /// <param name="layers"></param>
        /// <param name="imageFiles"></param>
        /// <param name="area"></param>
        /// <param name="areaDicList"></param>
        /// <returns></returns>
        private string CreateYZTHGJCWord(object[] proInfo, string[] layers, string[] imageFiles, double area, List<Dictionary<string, double>> areaDicList)
        {
            try
            {
                string wordName = "";

                WordOpr wo = new WordOpr();
                wo.killWinWordProcess();
                string outPath = Utils.workSpace + @"template_yzthgjc.doc";
                wo.OpenWithTemplate(outPath);
                wo.Builder();

                DocumentBuilder builder = wo.WordBuilder;
                object obj_name = null;
                object obj_type = null;
                object obj_prjNum = null;
                object obj_department = null;
                if (null != proInfo)
                {
                    JsonObject fsJson = (JsonObject)proInfo[0];

                    fsJson.TryGetObject("name", out obj_name);
                    fsJson.TryGetObject("type", out obj_type);
                    fsJson.TryGetObject("department", out obj_department);
                    fsJson.TryGetObject("bh", out obj_prjNum);
                }


                string prjName = obj_name == null ? "" : (String)obj_name;
                string prjNum = obj_prjNum == null ? "" : (String)obj_prjNum;
                string businessType = obj_type == null ? "" : (String)obj_type;
                string department = obj_department == null ? "" : (String)obj_department;

                builder.MoveToMergeField("xmmc1");
                builder.Write(prjName);
                builder.MoveToMergeField("xmbh");
                builder.Write(prjNum);
                builder.MoveToMergeField("xmmc2");
                builder.Write(prjName);

                builder.MoveToMergeField("ywlx");
                builder.Write(businessType);
                builder.MoveToMergeField("jsdw");
                builder.Write(department);
                builder.MoveToMergeField("ydfw");
                builder.Write((area * 0.0001).ToString("0.0000") + "/" + (area/666.7).ToString("0.00"));
                builder.MoveToMergeField("date");
                builder.Write(System.DateTime.Now.ToLongDateString());
                /// <summary>
                /// word文档输出表格
                /// </summary>
                DataTable tb = new DataTable();
                //构建输出表格
                //构建一张图合规检测输出表格
                tb.Columns.Add("num", typeof(string));
                tb.Columns.Add("layer", typeof(string));
                tb.Columns.Add("pic", typeof(string));
                tb.Columns.Add("typeName", typeof(string));
                tb.Columns.Add("area", typeof(string));
                tb.Columns.Add("conclusion", typeof(string));


                //获取表格宽度，记录后待填充完成，重新设置表格大小
                List<double> widthLst = new List<double>();
                for (int i = 0; i < 6; i++)
                {
                    builder.MoveToCell(0, 7, i, 0);
                    widthLst.Add(builder.CellFormat.Width);
                }
                //标记需要插入图片的行
                int[] rowPicFlag = new int[imageFiles.Length];
                int rowNum = 0;
                int rowText = 8;
                for (int i = 0; i < areaDicList.Count; i++)
                {

                    rowNum++;

                    rowPicFlag[i] = rowText;
                    //if (areaDicList[i].Count == 0)
                    //{
                    //    continue;
                    //}
                    Dictionary<string, double> areaDic = areaDicList[i];
                    //土地类型对应的名称
                    string tdlx = "";
                    //标识图层统计是否是第一条数据
                    bool keyCount = false;
                    string unitName = "面积(公顷/亩)";
                    string outputFormat = "0.00";
                    string outputFormatGongQin = "0.0000";
                    double fenMu = 667;
                    double gongqin = 0.0001;
                    bool isPoint = false;
                    if (Utils.layerTypeArray[i].ToLower().Contains("point"))
                    {
                        unitName = "个数";
                        outputFormat = "0";
                        outputFormatGongQin = "0";
                        fenMu = 1;
                        gongqin = 1;
                        isPoint = true;
                    }
                    //默认审查结果填写内容
                    string scjg = "无检测";
                    string fillNull = "/";
                    if (layers.Contains(i.ToString()))
                    {
                        if (isPoint)
                        {
                            scjg = "不包含";
                            fillNull = "无";
                        }
                        else
                        {
                            scjg = "未占用";
                            fillNull = "无";
                        }

                    }

                    if (areaDicList[i].Count == 0)
                    {
                        string layerName = "";
                        Utils.layerFieldNameDicList[i].TryGetValue("layerName", out layerName);
                        string type = "";
                        Utils.layerFieldNameDicList[i].TryGetValue("type", out type);
                        //序号|规划类型(图层)|图片||分类|面积|审查结果
                        //添加""时，下一行自动与上一行合并单元格
                        tb.Rows.Add(rowNum, layerName, " ", type, unitName, scjg);
                        rowText++;
                        tb.Rows.Add("", "", "", fillNull, fillNull, "");
                        rowText++;
                    }
                    else
                    {
                        foreach (string key in areaDic.Keys)
                        {

                            tdlx = key;
                            double areaValue;
                            areaDic.TryGetValue(key, out areaValue);
                            if (!keyCount)
                            {

                                string layerName = "";
                                Utils.layerFieldNameDicList[i].TryGetValue("layerName", out layerName);
                                string type = "";
                                Utils.layerFieldNameDicList[i].TryGetValue("type", out type);
                                //序号|规划类型(图层)|图片||分类|面积|审查结果

                                tb.Rows.Add(rowNum, layerName, " ", type, unitName, isPoint ? "包含" : "占用");
                                rowText++;
                                tb.Rows.Add("", "", "", tdlx, (areaValue *gongqin).ToString(outputFormatGongQin)+"/"+(areaValue / fenMu).ToString(outputFormat), "");
                                rowText++;
                            }
                            else
                            {
                                tb.Rows.Add("", "", "", tdlx, (areaValue * gongqin).ToString(outputFormatGongQin) + "/" + (areaValue / fenMu).ToString(outputFormat), "");
                                rowText++;
                            }
                            keyCount = true;
                        }
                    }


                }

                builder.MoveToDocumentStart();

                //表格追加tb内容行
                wo.AddTable(0, tb, widthLst);


                //得到文档中的第一个表格
                Aspose.Words.Tables.Table table = (Aspose.Words.Tables.Table)wo.Doc.GetChild(NodeType.Table, 0, true);
                //第一行第一列单元格
                for (int i = 0; i < imageFiles.Length; i++)
                {
                    if (String.IsNullOrEmpty(imageFiles[i]))
                    {
                        continue;
                    }
                    builder.MoveToDocumentStart();
                    builder.MoveToCell(0, rowPicFlag[i], 2, 0);
                    Aspose.Words.Drawing.Shape sp = builder.InsertImage(imageFiles[i]);
                    sp.Height = widthLst[2] * 0.6;
                    sp.Width = widthLst[2] * 0.6;
                    sp.VerticalAlignment = Aspose.Words.Drawing.VerticalAlignment.Center;
                    builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐
                    //会全部加粗
                    //Cell c = table.Rows[rowPicFlag[i]].Cells[3];
                    //c.FirstParagraph.ParagraphFormat.Style.Font.Bold=true;
                }


                //移动至第一行第一列(合并检测结果单元格)
                builder.MoveToCell(0, 7, 3, 0);
                builder.CellFormat.HorizontalMerge = CellMerge.First;//合并的第一个单元格
                //移动至第一行第二列
                builder.MoveToCell(0, 7, 4, 0);
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;//被合并的单元格
                string time = System.DateTime.Now.ToFileTime().ToString();
                wordName = time + ".pdf";
                //wo.SaveAs(Utils.wordDir + time + ".doc");
                wo.SaveAsPdf(Utils.wordDir+time + ".pdf");//导出pdf文件
                return wordName;
            }
            catch (Exception ex)
            {
                return null;
            }


        }

        /// <summary>
        /// 创建一张图合规检测文档（20210324）第三版 未使用
        /// </summary>
        /// <param name="proInfo"></param>
        /// <param name="layers"></param>
        /// <param name="imageFiles"></param>
        /// <param name="area"></param>
        /// <param name="areaDicList"></param>
        /// <returns></returns>
        private string CreateYZTHGJCWord2(object[] proInfo, string[] layers, string[] imageFiles, double area, List<Dictionary<string, double>> areaDicList)
        {
            try
            {
                string wordName = "";

                WordOpr wo = new WordOpr();
                wo.killWinWordProcess();
                string outPath = Utils.workSpace + @"template_yzthgjc.doc";
                wo.OpenWithTemplate(outPath);
                wo.Builder();

                DocumentBuilder builder = wo.WordBuilder;
                object obj_name = null;
                object obj_type = null;
                object obj_prjNum = null;
                object obj_department = null;
                if (null != proInfo)
                {
                    JsonObject fsJson = (JsonObject)proInfo[0];

                    fsJson.TryGetObject("name", out obj_name);
                    fsJson.TryGetObject("type", out obj_type);
                    fsJson.TryGetObject("department", out obj_department);
                    fsJson.TryGetObject("bh", out obj_prjNum);
                }


                string prjName = obj_name == null ? "" : (String)obj_name;
                string prjNum = obj_prjNum == null ? "" : (String)obj_prjNum;
                string businessType = obj_type == null ? "" : (String)obj_type;
                string department = obj_department == null ? "" : (String)obj_department;

                builder.MoveToMergeField("xmmc1");
                builder.Write(prjName);
                builder.MoveToMergeField("xmbh");
                builder.Write(prjNum);
                builder.MoveToMergeField("xmmc2");
                builder.Write(prjName);

                builder.MoveToMergeField("ywlx");
                builder.Write(businessType);
                builder.MoveToMergeField("jsdw");
                builder.Write(department);
                builder.MoveToMergeField("ydfw");
                builder.Write(Convert.ToDouble(Math.Round(area, 2).ToString("0.0")).ToString());
                builder.MoveToMergeField("date");
                builder.Write(System.DateTime.Now.ToLongDateString());
                /// <summary>
                /// word文档输出表格
                /// </summary>
                DataTable tb = new DataTable();
                //构建输出表格
                //构建一张图合规检测输出表格
                tb.Columns.Add("num", typeof(string));
                tb.Columns.Add("layer", typeof(string));
                tb.Columns.Add("typeName", typeof(string));
                tb.Columns.Add("area", typeof(string));
                tb.Columns.Add("result", typeof(string));


                //获取表格宽度，记录后待填充完成，重新设置表格大小
                List<double> widthLst = new List<double>();
                for (int i = 0; i < 5; i++)
                {
                    builder.MoveToCell(0, 7, i, 0);
                    widthLst.Add(builder.CellFormat.Width);
                }
                //标记需要插入图片的行
                int[] rowPicFlag = new int[imageFiles.Length];
                int rowNum = 0;
                int rowText = 7;
                for (int i = 0; i < areaDicList.Count; i++)
                //for (int i = 0; i < Utils.layerFieldNameDicList.Count; i++)
                {

                    
                    rowNum++;
                    //if (areaDicList[i].Count == 0)
                    //{
                    //    continue;
                    //}
                    Dictionary<string, double> areaDic = areaDicList[i];
                    //土地类型对应的名称
                    string tdlx = "";
                    //标识图层统计是否是第一条数据
                    bool keyCount = false;
                    string unitName = "面积（亩）";
                    string outputFormat = "0.00";
                    double fenMu = 667;
                    bool isPoint = false;
                    if (Utils.layerTypeArray[i].ToLower().Contains("point"))
                    {
                        unitName = "个数";
                        outputFormat = "0";
                        fenMu = 1;
                        isPoint = true;
                    }
                    //默认审查结果填写内容
                    string scjg = "无检测";
                    string fillNull = "/";
                    if (layers.Contains(i.ToString()))
                    {
                        if (isPoint)
                        {
                            scjg = "不包含";
                            fillNull = "无";
                        }
                        else
                        {
                            scjg = "未占用";
                            fillNull = "无";
                        }

                    }

                    if (areaDicList[i].Count == 0)
                    {
                        string layerName = "";
                        Utils.layerFieldNameDicList[i].TryGetValue("layerName", out layerName);
                        string type = "";
                        Utils.layerFieldNameDicList[i].TryGetValue("type", out type);
                        //序号|规划类型(图层)|图片||分类|面积|审查结果
                        //添加""时，下一行自动与上一行合并单元格
                        tb.Rows.Add(rowNum, layerName, type, unitName, scjg);
                        rowText++;
                        tb.Rows.Add("", "",  fillNull, fillNull, "");
                        rowText++;
                    }
                    else
                    {
                        foreach (string key in areaDic.Keys)
                        {

                            tdlx = key;
                            double areaValue;
                            areaDic.TryGetValue(key, out areaValue);
                            if (!keyCount)
                            {

                                string layerName = "";
                                Utils.layerFieldNameDicList[i].TryGetValue("layerName", out layerName);
                                string type = "";
                                Utils.layerFieldNameDicList[i].TryGetValue("type", out type);
                                //序号|规划类型(图层)|图片||分类|面积|审查结果

                                tb.Rows.Add(rowNum, layerName,  type, unitName, isPoint ? "包含" : "占用");
                                rowText++;
                                tb.Rows.Add("", "", tdlx, (areaValue / fenMu).ToString(outputFormat), "");
                                rowText++;
                            }
                            else
                            {
                                tb.Rows.Add("", "", tdlx, (areaValue / fenMu).ToString(outputFormat), "");
                                rowText++;
                            }
                            keyCount = true;
                        }
                        //当前图层末尾追加一行存放图片
                        
                    }
                    if (!String.IsNullOrEmpty(imageFiles[i]))
                    {
                        tb.Rows.Add("", "", " ", " ", "");
                        rowText++;
                    }
                    rowPicFlag[i] = rowText;
                }

                builder.MoveToDocumentStart();

                //表格追加tb内容行
                wo.AddTable(0, tb, widthLst);


                //得到文档中的第一个表格
                Aspose.Words.Tables.Table table = (Aspose.Words.Tables.Table)wo.Doc.GetChild(NodeType.Table, 0, true);
                //第一行第一列单元格
                for (int i = 0; i < imageFiles.Length; i++)
                {
                    if (String.IsNullOrEmpty(imageFiles[i]))
                    {
                        continue;
                    }
                    builder.MoveToDocumentStart();
                    builder.MoveToCell(0, rowPicFlag[i], 2, 0);
                    Aspose.Words.Drawing.Shape sp = builder.InsertImage(imageFiles[i]);
                    sp.Height = widthLst[2] * 0.6;
                    sp.Width = widthLst[2] * 0.6;
                    sp.VerticalAlignment = Aspose.Words.Drawing.VerticalAlignment.Center;
                    builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐
                    //合并图片所在及右边单元格
                    builder.MoveToCell(0, rowPicFlag[i], 2, 0);
                    builder.CellFormat.HorizontalMerge = CellMerge.First;//合并的第一个单元格
                    //移动至第一行第二列
                    builder.MoveToCell(0, rowPicFlag[i], 3, 0);
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;//被合并的单元格
                    //会全部加粗
                    //Cell c = table.Rows[rowPicFlag[i]].Cells[3];
                    //c.FirstParagraph.ParagraphFormat.Style.Font.Bold=true;
                }


                //移动至第一行第一列(合并检测结果单元格)
                builder.MoveToCell(0, 7, 2, 0);
                builder.CellFormat.HorizontalMerge = CellMerge.First;//合并的第一个单元格
                //移动至第一行第二列
                builder.MoveToCell(0, 7, 3, 0);
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;//被合并的单元格
                string time = System.DateTime.Now.ToFileTime().ToString();
                wordName = time + ".pdf";
                //wo.SaveAs(Utils.wordDir + time + ".doc");
                wo.SaveAsPdf(Utils.wordDir + time + ".pdf");//导出pdf文件
                return wordName;
            }
            catch (Exception ex)
            {
                return null;
            }


        }

        /// <summary>
        /// 填充输出表格数据
        /// </summary>
        /// <param name="layerId">相交分析的图层</param>
        /// <param name="dArea">分析结果面积</param>
        public void FillDataTable(int layerId, double dArea, string type)
        {
            try
            {
                //组织输出表格数据
                //是否预警
                string warnStr;
                if (dArea > 0.001)
                {
                    warnStr = "预警，该范围" + type + Utils.featureClsArray[layerId].AliasName;
                }
                else
                {
                    warnStr = "无预警";
                }

                //tbl.Rows.Add(tbl.Rows.Count + 1, Utils.featureClsArray[layerId].AliasName, warnStr, dArea.ToString());
                tbl.Rows.Add(tbl.Rows.Count + 1, Utils.featureClsArray[layerId].AliasName, dArea.ToString());
            }
            catch (System.Exception ex)
            {

            }
        }
        /// <summary>
        /// 生成差异分析word文档
        /// </summary>
        /// <param name="count">个数</param>
        /// <param name="dArea">面积</param>
        /// <param name="proportion">占比</param>
        /// <returns></returns>
        private string CreateCYFXWord(int[] count, double[] dArea, string[] proportion, string imageFile)
        {
            try
            {
                string wordName = "";

                WordOpr wo = new WordOpr();
                wo.killWinWordProcess();
                string outPath = Utils.workSpace + @"template_kzx.doc";
                wo.OpenWithTemplate(outPath);
                wo.Builder();

                DocumentBuilder builder = wo.WordBuilder;

                builder.MoveToMergeField("date");
                builder.Write(System.DateTime.Now.ToLongDateString());

                for (int i = 1; i < 5; i++)
                {

                    builder.MoveToCell(0, i, 2, 0);
                    builder.Write(count[i - 1].ToString());
                    builder.MoveToCell(0, i, 3, 0);
                    builder.Write(dArea[i - 1].ToString());
                    builder.MoveToCell(0, i, 4, 0);
                    builder.Write(proportion[i - 1].ToString());
                }


                builder.MoveToDocumentStart();
                builder.MoveToMergeField("pic");
                string[] picsArr = imageFile.Split((','));
                for (int i = 0; i < picsArr.Length - 1; i++)
                {
                    builder.InsertImage(picsArr[i]);
                }

                wordName = System.DateTime.Now.ToFileTime().ToString() + ".doc";
                wo.SaveAs(Utils.wordDir + wordName);
                return wordName;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

    }
}
