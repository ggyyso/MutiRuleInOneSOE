using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

namespace MultiRuleInOneSOE
{
    class ExportMap
    {
        /// <summary>
        /// 根据MXD模板动态设置信息后出专题图
        /// </summary>
        /// <param name="themeName">专题名称</param>
        /// <param name="layerId">图层id</param>
        public static string exportMap(int layerId, List<IGeometry> inputGeometryList)
        {
            IFeatureClass kzxFs = Utils.featureClsArray[layerId];
            string themeName = kzxFs.AliasName;

            string mxdPath = Utils.workSpace + Utils.name_MXD + ".mxd";

            if (!File.Exists(mxdPath))
            {
                return null;
            }
            else
            {

                //打开当前Mxd文档
                IMapDocument mapDoc = new MapDocumentClass();
                mapDoc.Open(mxdPath, "");

                string mxdFileName = mapDoc.DocumentFilename;

                IPageLayout pageLayout = mapDoc.PageLayout;

                IMap map = mapDoc.ActiveView.FocusMap;
                //修改当前图层为可见，其他图层不可见
                for (int L = 0; L < map.LayerCount; L++)
                {
                    map.get_Layer(L).Visible = false;
                }
                //只显示第0个图层input和当前图层
                // map.get_Layer(0).Visible = true;
                map.get_Layer(layerId + 1).Visible = true;
                try
                {
                    //mapDoc.Save();
                }
                catch (System.Exception ex)
                {
                    mapDoc.Close();
                    throw ex;
                }
                IActiveView mapActiveView = map as IActiveView;
                IActiveView pageLayoutActiveView = pageLayout as IActiveView;

                //设置图名
                //IElement tmElement = GetPageLayoutElementByName(pageLayout, "Title");
                //if (tmElement != null)
                //{
                //    ITextElement textElement = tmElement as ITextElement;

                //    textElement.Text = themeName + "检测结果图";

                //}
                //绘制输入范围的Element
                IRgbColor pRgbColor;
                IActiveView pActiveView;


                IGraphicsContainer pGraphicsContainer = mapActiveView.FocusMap as IGraphicsContainer; ;//起到容器的作用，可使画的内容不消失
                pRgbColor = new RgbColorClass();
                pRgbColor.Red = 255;
                pRgbColor.Green = 150;
                pRgbColor.Blue = 0;
                IRgbColor pRgbColor2 = new RgbColorClass();
                pRgbColor2.Red = 255;
                pRgbColor2.Green = 0;
                pRgbColor2.Blue = 0;

                pActiveView = mapActiveView;
                ILineFillSymbol pLineFillSymbol = new LineFillSymbolClass();
                pLineFillSymbol.LineSymbol.Color = pRgbColor;
                pLineFillSymbol.Color = pRgbColor;
                pLineFillSymbol.Angle = 45;
                pLineFillSymbol.Separation = 20;
                pLineFillSymbol.Outline.Width = 10;
                pLineFillSymbol.Outline.Color = pRgbColor2;

                // 多边形符号
                ISimpleFillSymbol pSimpleFillSymbol = new SimpleFillSymbol();
                pSimpleFillSymbol.Color = pRgbColor;
                pSimpleFillSymbol.Style = esriSimpleFillStyle.esriSFSNull;
                pSimpleFillSymbol.Outline.Width = 10;
                pSimpleFillSymbol.Outline.Color = pRgbColor2;

                
                // 线符号
                ISimpleLineSymbol pSimpleLineSymbol = new SimpleLineSymbol();
                pSimpleLineSymbol.Color = pRgbColor2;
                pSimpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
                pSimpleLineSymbol.Width = 10;
                pLineFillSymbol.LineSymbol = pSimpleLineSymbol;


                foreach (IGeometry pgeo in inputGeometryList)
                {
                    // 绘制多边形元素
                    IElement pElement = new PolygonElement();
                    pElement.Geometry = pgeo;
                    //pFillShapeElement.Symbol.Outline = pSimpleLineSymbol as ILineSymbol;
                    IFillShapeElement pFillShapeElement = pElement as IFillShapeElement;
                    pFillShapeElement.Symbol = (IFillSymbol)CreateSimpleFillSymbol(Color.Red, 1, esriSimpleFillStyle.esriSFSDiagonalCross);

                    pGraphicsContainer.AddElement(pElement as IElement, 0);
                }
                pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);

                mapActiveView.Extent = Utils.pfwEnv;

                Guid guid = Guid.NewGuid();
                string fileName = guid.ToString().Replace("-", "_");

                string tempImage = Utils.workSpace + @"wordFiles\" + themeName + "检测结果图" + fileName + ".png";

                ExportRasterFile(pageLayoutActiveView, tempImage, "PNG", 140);
                mapDoc.Close();

                return tempImage;
            }
        }

        public static ISymbol CreateSimpleLineSymbol(Color color, int width, esriSimpleLineStyle style)
        {
            ISimpleLineSymbol pSimpleLineSymbol;
            pSimpleLineSymbol = new SimpleLineSymbol();
            pSimpleLineSymbol.Width = width;
            pSimpleLineSymbol.Color = GetColor(color.R, color.G, color.B);
            pSimpleLineSymbol.Style = style;
            return (ISymbol)pSimpleLineSymbol;

        }

        public static IRgbColor GetColor(int r, int g, int b)
        {
            RgbColor color = new RgbColor();
            color.Red = r;
            color.Green = g;
            color.Blue = b;
            return color;
        }

        public static ISymbol CreateSimpleFillSymbol(Color fillColor, int oLineWidth, esriSimpleFillStyle fillStyle)
        {
            ISimpleFillSymbol pSimpleFillSymbol;
            pSimpleFillSymbol = new SimpleFillSymbol();
            pSimpleFillSymbol.Style = fillStyle;
            pSimpleFillSymbol.Color = GetColor(fillColor.R, fillColor.G, fillColor.B);
            pSimpleFillSymbol.Outline = (ILineSymbol)CreateSimpleLineSymbol(fillColor, 2, esriSimpleLineStyle.esriSLSSolid);
           
            return (ISymbol)pSimpleFillSymbol;

        }

        public static ISymbol CreateLineFillSymbol(Color fillColor, int oLineWidth, esriSimpleFillStyle fillStyle)
        {
            ILineFillSymbol pLineFillSymbol;
            pLineFillSymbol = new LineFillSymbol();
            pLineFillSymbol.LineSymbol = (ILineSymbol)CreateSimpleLineSymbol(Color.Yellow, 1, esriSimpleLineStyle.esriSLSSolid);
            pLineFillSymbol.Separation = 5;
            pLineFillSymbol.Color = GetColor(fillColor.R, fillColor.G, fillColor.B);
            pLineFillSymbol.Outline = (ILineSymbol)CreateSimpleLineSymbol(fillColor, 2, esriSimpleLineStyle.esriSLSSolid);

            return (ISymbol)pLineFillSymbol;

        }
        /// <summary>
        /// 根据MXD模板动态设置信息后出规划检测专题图
        /// </summary>
        /// <param name="themeName">专题名称</param>
        /// <param name="layerId">图层id</param>
        public static string exportGHJCMap(int layerId)
        {
            IFeatureClass kzxFs = Utils.featureClsArray[layerId];
            string themeName = kzxFs.AliasName;

            string mxdPath = Utils.workSpace + Utils.GHX_MXD + layerId.ToString() + ".mxd";

            if (!File.Exists(mxdPath))
            {
                return null;
            }
            else
            {

                //打开当前Mxd文档
                IMapDocument mapDoc = new MapDocumentClass();
                mapDoc.Open(mxdPath, "");

                string mxdFileName = mapDoc.DocumentFilename;

                IPageLayout pageLayout = mapDoc.PageLayout;

                IMap map = mapDoc.ActiveView.FocusMap;
                IActiveView mapActiveView = map as IActiveView;
                IActiveView pageLayoutActiveView = pageLayout as IActiveView;

                //设置图名
                IElement tmElement = GetPageLayoutElementByName(pageLayout, "Title");
                if (tmElement != null)
                {
                    ITextElement textElement = tmElement as ITextElement;

                    textElement.Text = themeName + "检测结果图";

                }

                mapActiveView.Extent = Utils.pfwEnv;

                Guid guid = Guid.NewGuid();
                string fileName = guid.ToString().Replace("-", "_");

                string tempImage = Utils.workSpace + @"tempPicFiles\" + themeName + "检测结果图" + fileName + ".png";

                ExportRasterFile(pageLayoutActiveView, tempImage, "PNG", 140);

                return tempImage;
            }
        }

        /// <summary>
        /// 输出图片
        /// </summary>
        /// <param name="activeView"></param>
        /// <param name="filename"></param>
        /// <param name="format"></param>
        /// <param name="resolution"></param>
        public static void ExportRasterFile(IActiveView activeView, string filename, string format, double resolution)
        {
            if (activeView == null || filename.Length == 0)
                return;

            double screenResolution = resolution;
            double outputResolution = resolution;

            IExport export = null;
            switch (format)
            {
                case "PDF":
                    export = new ExportPDFClass();

                    IExportVectorOptions exportVectorOptions = export as IExportVectorOptions;
                    exportVectorOptions.PolygonizeMarkers = true;

                    IExportPDF exportPDF = export as IExportPDF;
                    exportPDF.EmbedFonts = true;
                    break;
                case "BMP":
                    export = new ExportBMPClass();
                    break;
                case "JPG":
                    export = new ExportJPEGClass();
                    if (filename.IndexOf("彩信") != -1)
                    {
                        IExportJPEG exportJPEG = export as IExportJPEG;
                        exportJPEG.Quality = 70;
                        outputResolution = 70.0;
                    }
                    break;
                case "PNG":
                    export = new ExportPNGClass();
                    break;
                case "TIF":
                    export = new ExportTIFFClass();
                    break;
                case "GIF":
                    export = new ExportGIFClass();
                    break;
                case "EMF":
                    export = new ExportEMFClass();
                    break;
                case "SVG":
                    export = new ExportSVGClass();
                    break;
                case "AI":
                    export = new ExportAIClass();
                    break;
                case "EPS":
                    export = new ExportPSClass();
                    break;
            }

            IGraphicsContainer docGraphicsContainer;
            IElement docElement;
            IOutputRasterSettings docOutputRasterSettings;
            IMapFrame docMapFrame;
            IActiveView tmpActiveView;
            IOutputRasterSettings doOutputRasterSettings;
            if (activeView is IMap)
            {
                doOutputRasterSettings = activeView.ScreenDisplay.DisplayTransformation as IOutputRasterSettings;
                doOutputRasterSettings.ResampleRatio = 3;
            }
            else if (activeView is IPageLayout)
            {
                doOutputRasterSettings = activeView.ScreenDisplay.DisplayTransformation as IOutputRasterSettings;
                doOutputRasterSettings.ResampleRatio = 4;
                //and assign ResampleRatio to the maps in the PageLayout.
                docGraphicsContainer = activeView as IGraphicsContainer;
                docGraphicsContainer.Reset();
                docElement = docGraphicsContainer.Next();
                int c = 0;

                while (docElement != null)
                {
                    c += 1;
                    if (docElement is IMapFrame)
                    {
                        docMapFrame = docElement as IMapFrame;
                        tmpActiveView = docMapFrame.Map as IActiveView;
                        docOutputRasterSettings = tmpActiveView.ScreenDisplay.DisplayTransformation as IOutputRasterSettings;
                        docOutputRasterSettings.ResampleRatio = 1;
                    }
                    docElement = docGraphicsContainer.Next();
                }


                docMapFrame = null;
                docGraphicsContainer = null;
                tmpActiveView = null;

            }
            else
                docOutputRasterSettings = null;
            //end




            export.ExportFileName = filename;
            export.Resolution = outputResolution;


            tagRECT exportRECT = activeView.ExportFrame;
            exportRECT.right = (int)(exportRECT.right * (outputResolution / screenResolution));
            exportRECT.bottom = (int)(exportRECT.bottom * (outputResolution / screenResolution));

            IEnvelope envelope = new EnvelopeClass();
            envelope.PutCoords(exportRECT.left, exportRECT.top, exportRECT.right, exportRECT.bottom);
            export.PixelBounds = envelope;

            int hDC = export.StartExporting();
            activeView.Output(hDC, Convert.ToInt32(export.Resolution), ref exportRECT, null, null);
            export.FinishExporting();
            export.Cleanup();

        }
        /// <summary>
        /// 根据指定名称获取元素
        /// </summary>
        /// <param name="pageLayout"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static IElement GetPageLayoutElementByName(IPageLayout pageLayout, string name)
        {
            IElement element = null;
            IGraphicsContainer graphicsContainer = pageLayout as IGraphicsContainer;
            graphicsContainer.Reset();
            element = graphicsContainer.Next();
            while (element != null)
            {
                IElementProperties elementProperties = element as IElementProperties;
                if (elementProperties.Name.ToLower().Contains(name.ToLower()))
                    return element;

                element = graphicsContainer.Next();
            }
            return null;
        }
    }
}
