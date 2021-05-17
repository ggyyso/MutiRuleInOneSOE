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
using ESRI.ArcGIS.esriSystem;

namespace MultiRuleInOneSOE
{
    class SpatialOperations
    {
        public static IGeometry projectGeometry(IGeometry polygon, int pcs)
        {
            ISpatialReferenceFactory spatialReferenceFactory = new SpatialReferenceEnvironmentClass();
            ISpatialReference pFromSpatialReference = spatialReferenceFactory.CreateGeographicCoordinateSystem(4490);
            if (polygon.SpatialReference!=null)
            {
                pFromSpatialReference = polygon.SpatialReference;
            }
            ISpatialReference pToSpatialReference;
            if(pcs==4326 || pcs==4490)
            {
                pToSpatialReference = spatialReferenceFactory.CreateGeographicCoordinateSystem(4490);
            }
            else
            {
                pToSpatialReference = spatialReferenceFactory.CreateProjectedCoordinateSystem(pcs);
            }

            polygon.SpatialReference = pFromSpatialReference;
            polygon.Project(pToSpatialReference);

            return polygon;
        }

        /// <summary>
        /// 空间求交
        /// </summary>
        /// <param name="featureClassPath"></param>
        /// <param name="layerId"></param>
        /// <returns></returns>
        public static IFeatureClass spatialIntersect(string gdb,IFeatureClass features, int layerId)
        {

            IFeatureClass kzxFs = Utils.featureClsArray[layerId];

            string layer_name = kzxFs.AliasName;

            Geoprocessor gp = new Geoprocessor();
            Intersect intersect = new Intersect();

            string fs1 = gdb + @"\inputs\input";
            string fs2 = gdb + @"\kzx\" + layer_name;
            intersect.in_features = fs1 + ";" + fs2;
            intersect.out_feature_class = gdb + @"\results\" + layer_name + @"IntersectResult";

            try
            {
                gp.AddOutputsToMap = false;
                gp.OverwriteOutput = true;
                gp.Execute(intersect, null);

                gp.ResetEnvironments();
                IFeatureClass outFs = FeatureOperations.OpenFeatureClass(gdb, "results", layer_name + @"IntersectResult");

                return outFs;
            }
            catch (System.Exception ex)
            {
                string error = "";
                for (int i = 0; i < gp.MessageCount; i++)
                {
                    error += gp.GetMessage(i);
                }
                return null;
            }

        }

        /// <summary>
        /// 空间擦除
        /// </summary>
        /// <param name="featureClassPath"></param>
        /// <param name="layerId"></param>
        /// <param name="outFsName"></param>
        /// <returns></returns>
        public static IFeatureClass spatialErase(string gdb,IFeatureClass features, int layerId)
        {
            IFeatureClass kzxFs = Utils.featureClsArray[layerId];

            string layer_name = kzxFs.AliasName;

            Geoprocessor gp = new Geoprocessor();
            Erase erase = new Erase();

            erase.in_features = features;
            erase.erase_features = kzxFs;

            erase.out_feature_class = gdb + @"\results\" + layer_name + @"EraseResult";

            try
            {
                gp.AddOutputsToMap = false;
                gp.OverwriteOutput = true;
                gp.Execute(erase, null);

                gp.ResetEnvironments();
                IFeatureClass outFs = FeatureOperations.OpenFeatureClass(gdb, "results", layer_name + @"EraseResult");

                return outFs;
            }
            catch (System.Exception ex)
            {
                string error = "";
                for (int i = 0; i < gp.MessageCount; i++)
                {
                    error += gp.GetMessage(i);
                }
                return null;
            }

        }

       


        internal static List<IGeometry> spatialIntersect(IGeometry pGeometry, List<IGeometry> inputGeometryList)
        {
            
            List<IGeometry> resultGeoList = new List<IGeometry>();
            if (pGeometry.GeometryType == esriGeometryType.esriGeometryPoint || pGeometry.GeometryType == esriGeometryType.esriGeometryMultipoint)
            {
                resultGeoList.Add(pGeometry);
                return resultGeoList;
            }
            
            foreach (IGeometry geo in inputGeometryList)
            {
                ITopologicalOperator2 topoOper = pGeometry as ITopologicalOperator2;
                IGeometry resultGeo = topoOper.Intersect(geo, esriGeometryDimension.esriGeometry2Dimension);
                if (resultGeo != null)
                {
                    resultGeoList.Add(resultGeo);
                }
            }

            return resultGeoList;
        }
    }
}
