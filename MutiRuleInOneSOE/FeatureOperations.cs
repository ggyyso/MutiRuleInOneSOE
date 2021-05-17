using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.SOESupport;

namespace MultiRuleInOneSOE
{
    class FeatureOperations
    {
        /// <summary>
        /// 清空要素类中的所有要素
        /// </summary>
        /// 输入要素类
        /// <param name="fs"></param>
        public static void DeleteAllFeatureFromFeatureClass(IFeatureClass fs)
        {
            IFeatureCursor cursor = fs.Search(new QueryFilterClass
            {
                WhereClause = "1=1"
            }, false);
            IFeature feat = cursor.NextFeature();
            int i = 0;
            while (feat != null)
            {
                feat.Delete();
                feat = cursor.NextFeature();
                i++;
            }
            Marshal.ReleaseComObject(cursor);

        }
        /// <summary>
        /// 将要素类序列化成json格式对象
        /// </summary>
        /// <param name="inputFeaClass">输入要素类</param>
        /// <returns></returns>
        public static JsonObject FclassToJsonObj(IFeatureClass inputFeaClass)
        {
            //获取要素数目
            IQueryFilter pQueryFilter = new QueryFilterClass();
            pQueryFilter.WhereClause = null;
            int count = inputFeaClass.FeatureCount(pQueryFilter);

            //将每一个要素序列化成json数据
            IFeature pFeature = null;
            List<JsonObject> jsonGeometries = new List<JsonObject>();
            for (int i = 1; i < count; i++)//OBJECTID从1开始
            {
                pFeature = inputFeaClass.GetFeature(i);
                IGeometry pGeometry = pFeature.Shape;
                JsonObject featureJson = new JsonObject();
                JsonObject feaGeoJson = null;//几何对象
                if (pGeometry != null)
                {
                    feaGeoJson = Conversion.ToJsonObject(pGeometry);
                    featureJson.AddJsonObject("geometry", feaGeoJson);//加入几何对象
                }

                jsonGeometries.Add(featureJson);
            }

            JsonObject resultJson = new JsonObject();
            resultJson.AddArray("geometries", jsonGeometries.ToArray());
            return resultJson;
        }
        /// <summary>
        /// 打开数据库中FeatureClasss
        /// </summary>
        /// <param name="FileGdbName"></param>
        /// <param name="featureDataset"></param>
        /// <param name="featureclassname"></param>
        /// <returns></returns>
        public static IFeatureClass OpenFeatureClass(string FileGdbName, string featureDataset, string featureclassname)
        {
            IWorkspaceFactory pworkspF = new FileGDBWorkspaceFactoryClass();
            IWorkspace workspace = pworkspF.OpenFromFile(FileGdbName, 0);
            IFeatureWorkspace featureWorkspace = workspace as IFeatureWorkspace;

            IFeatureClass featureClass = null;
            IFeatureDataset pFeatureDataset = featureWorkspace.OpenFeatureDataset(featureDataset);
            if (pFeatureDataset == null)
                return null;
            IEnumDataset pEnumDataset = pFeatureDataset.Subsets;
            pEnumDataset.Reset();
            IDataset dataset = pEnumDataset.Next();
            while (dataset != null)
            {
                if (dataset.Name == featureclassname)
                {
                    featureClass = dataset as IFeatureClass;
                    break;
                }
                dataset = pEnumDataset.Next();
            }


            return featureClass;
        }

        /// <summary>
        /// 创建FeatureClasss文件
        /// </summary>
        /// <param name="_strFullPath">文件名</param>
        /// <param name="spatial_reference">空间参考</param>
        /// <returns></returns>
        public static IFeatureClass CreateFeatureClass(string _strFullPath, string spatial_reference)
        {
            int index = _strFullPath.LastIndexOf("\\");
            string strShapeFolder = _strFullPath.Substring(0, index);
            string strShapeFile = _strFullPath.Substring(index + 1);
            DirectoryInfo di = new DirectoryInfo(strShapeFolder);
            if (!di.Exists)
            {
                di.Create();
            }
            Geoprocessor gp = new Geoprocessor();
            ESRI.ArcGIS.DataManagementTools.CreateFeatureclass createFs = new ESRI.ArcGIS.DataManagementTools.CreateFeatureclass();
            createFs.geometry_type = "POLYGON";
            createFs.out_name = strShapeFile;
            createFs.out_path = strShapeFolder;
            createFs.spatial_reference = spatial_reference;
            try
            {
                gp.AddOutputsToMap = false;
                gp.OverwriteOutput = true;
                gp.Execute(createFs, null);
            }
            catch
            {
                string error = "";
                for (int i = 0; i < gp.MessageCount; i++)
                {
                    error += gp.GetMessage(i);
                }
            }
            gp.ResetEnvironments();

            IWorkspaceFactory pWorkspaceFactory = new ShapefileWorkspaceFactoryClass();
            IFeatureWorkspace pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strShapeFolder, 0) as IFeatureWorkspace;
            IFeatureClass pFeatureClass = pFeatureWorkspace.OpenFeatureClass(strShapeFile);
            return pFeatureClass;
        }
        /// <summary>
        /// 获取Shape数据的FeatureClass
        /// </summary>
        /// <param name="_strFullPath"></param>
        /// <returns></returns>
        public static IFeatureClass OpenFeatureClass(string _strFullPath)
        {

            try
            {
                int index = _strFullPath.LastIndexOf("\\");
                string strShapeFolder = _strFullPath.Substring(0, index);
                string strShapeFile = _strFullPath.Substring(index + 1);
                IWorkspaceFactory pWorkspaceFactory = new ShapefileWorkspaceFactoryClass();
                IFeatureWorkspace pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strShapeFolder, 0) as IFeatureWorkspace;
                IFeatureClass pFeatureClass = pFeatureWorkspace.OpenFeatureClass(strShapeFile);
                return pFeatureClass;
            }
            catch (System.Exception ex)
            {

            }
            return null;
        }

    }
}
