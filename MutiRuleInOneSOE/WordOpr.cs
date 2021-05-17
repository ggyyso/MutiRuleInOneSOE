using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using System.IO;
using Aspose.Words.Tables;

namespace MultiRuleInOneSOE
{
    class WordOpr
    {
        private DocumentBuilder _WordBuilder; //   a   reference   to   Word   application 

        public DocumentBuilder WordBuilder
        {
            get { return _WordBuilder; }
            set { _WordBuilder = value; }
        }

        private Aspose.Words.Document _Doc; //   a   reference   to   the   document   

        public Aspose.Words.Document Doc
        {
            get { return _Doc; }
            set { _Doc = value; }
        }
        /// <summary>
        /// word版本
        /// </summary>

        private int _Docversion = 2003;

        public int Docversion
        {
            get { return _Docversion; }
            set { _Docversion = value; }
        }

        public void OpenWithTemplate(string strFileName)
        {
            if (!string.IsNullOrEmpty(strFileName))
            {
                _Doc = new Aspose.Words.Document(strFileName);
            }
        }

        public void Open()
        {
            _Doc = new Aspose.Words.Document();
        }

        public void Builder()
        {
            _WordBuilder = new DocumentBuilder(_Doc);


        }
        /// <summary>  
        /// 保存文件  
        /// </summary>  
        /// <param name="strFileName"></param>  
        public void SaveAs(string strFileName)
        {

            if (this.Docversion == 2007)
            {
                _Doc.Save(strFileName, SaveFormat.Docx);
            }
            else
            {
                _Doc.Save(strFileName, SaveFormat.Doc);
            }

        }

        /// <summary>  
        /// 添加内容  
        /// </summary>  
        /// <param name="strText"></param>  
        public void InsertText(string strText, double conSize, bool conBold, string conAlign)
        {
            _WordBuilder.Bold = conBold;
            _WordBuilder.Font.Size = conSize;
            switch (conAlign)
            {
                case "left":
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    break;
                case "center":
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    break;
                case "right":
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    break;
                default:
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    break;
            }
            _WordBuilder.Writeln(strText);

        }

        /// <summary>  
        /// 添加内容  
        /// </summary>  
        /// <param name="strText"></param>  
        public void WriteText(string strText, double conSize, bool conBold, string conAlign)
        {
            _WordBuilder.Bold = conBold;
            _WordBuilder.Font.Size = conSize;
            switch (conAlign)
            {
                case "left":
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    break;
                case "center":
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    break;
                case "right":
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    break;
                default:
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    break;
            }
            _WordBuilder.Write(strText);

        }


        #region 设置纸张
        public void setPaperSize(string papersize)
        {

            switch (papersize)
            {
                case "A4":
                    foreach (Aspose.Words.Section section in _Doc)
                    {
                        section.PageSetup.PaperSize = PaperSize.A4;
                        section.PageSetup.Orientation = Orientation.Portrait;
                        section.PageSetup.VerticalAlignment = Aspose.Words.PageVerticalAlignment.Top;
                    }
                    break;
                case "A4H"://A4横向  
                    foreach (Aspose.Words.Section section in _Doc)
                    {
                        section.PageSetup.PaperSize = PaperSize.A4;
                        section.PageSetup.Orientation = Orientation.Landscape;
                        section.PageSetup.TextColumns.SetCount(2);
                        section.PageSetup.TextColumns.EvenlySpaced = true;
                        section.PageSetup.TextColumns.LineBetween = true;
                        //section.PageSetup.LeftMargin = double.Parse("3.35");  
                        //section.PageSetup.RightMargin =double.Parse("0.99");  
                    }
                    break;
                case "A3":
                    foreach (Aspose.Words.Section section in _Doc)
                    {
                        section.PageSetup.PaperSize = PaperSize.A3;
                        section.PageSetup.Orientation = Orientation.Portrait;

                    }

                    break;
                case "A3H"://A3横向  

                    foreach (Aspose.Words.Section section in _Doc)
                    {
                        section.PageSetup.PaperSize = PaperSize.A3;
                        section.PageSetup.Orientation = Orientation.Landscape;
                        section.PageSetup.TextColumns.SetCount(2);
                        section.PageSetup.TextColumns.EvenlySpaced = true;
                        section.PageSetup.TextColumns.LineBetween = true;

                    }

                    break;

                case "16K":

                    foreach (Aspose.Words.Section section in _Doc)
                    {
                        section.PageSetup.PaperSize = PaperSize.B5;
                        section.PageSetup.Orientation = Orientation.Portrait;

                    }

                    break;

                case "8KH":

                    foreach (Aspose.Words.Section section in _Doc)
                    {

                        section.PageSetup.PageWidth = double.Parse("36.4 ");//纸张宽度  
                        section.PageSetup.PageHeight = double.Parse("25.7");//纸张高度  
                        section.PageSetup.Orientation = Orientation.Landscape;
                        section.PageSetup.TextColumns.SetCount(2);
                        section.PageSetup.TextColumns.EvenlySpaced = true;
                        section.PageSetup.TextColumns.LineBetween = true;
                        //section.PageSetup.LeftMargin = double.Parse("3.35");  
                        //section.PageSetup.RightMargin = double.Parse("0.99");  
                    }



                    break;
            }
        }
        #endregion

        public void SetHeade(string strBookmarkName, string text)
        {
            if (_Doc.Range.Bookmarks[strBookmarkName] != null)
            {
                Aspose.Words.Bookmark mark = _Doc.Range.Bookmarks[strBookmarkName];
                mark.Text = text;
            }
        }
        public void InsertFile(string vfilename)
        {
            Aspose.Words.Document srcDoc = new Aspose.Words.Document(vfilename);
            Node insertAfterNode = _WordBuilder.CurrentParagraph.PreviousSibling;
            InsertDocument(insertAfterNode, _Doc, srcDoc);

        }

        public void InsertFile(string vfilename, string strBookmarkName, int pNum)
        {
            //Aspose.Words.Document srcDoc = new Aspose.Words.Document(vfilename);  
            //Aspose.Words.Bookmark bookmark = oDoc.Range.Bookmarks[strBookmarkName];  
            //InsertDocument(bookmark.BookmarkStart.ParentNode, srcDoc);  
            //替换插入word内容  
            _WordBuilder.Document.Range.Replace(new System.Text.RegularExpressions.Regex(strBookmarkName),
                new InsertDocumentAtReplaceHandler(vfilename, pNum), false);


        }
        /// <summary>  
        /// 插入word内容  
        /// </summary>  
        /// <param name="insertAfterNode"></param>  
        /// <param name="mainDoc"></param>  
        /// <param name="srcDoc"></param>  
        public static void InsertDocument(Node insertAfterNode, Aspose.Words.Document mainDoc, Aspose.Words.Document srcDoc)
        {
            // Make sure that the node is either a pargraph or table.  
            if ((insertAfterNode.NodeType != NodeType.Paragraph)
                & (insertAfterNode.NodeType != NodeType.Table))
                throw new Exception("The destination node should be either a paragraph or table.");

            //We will be inserting into the parent of the destination paragraph.  

            CompositeNode dstStory = insertAfterNode.ParentNode;

            //Remove empty paragraphs from the end of document  

            while (null != srcDoc.LastSection.Body.LastParagraph && !srcDoc.LastSection.Body.LastParagraph.HasChildNodes)
            {
                srcDoc.LastSection.Body.LastParagraph.Remove();
            }
            NodeImporter importer = new NodeImporter(srcDoc, mainDoc, ImportFormatMode.KeepSourceFormatting);

            //Loop through all sections in the source document.  

            int sectCount = srcDoc.Sections.Count;

            for (int sectIndex = 0; sectIndex < sectCount; sectIndex++)
            {
                Aspose.Words.Section srcSection = srcDoc.Sections[sectIndex];
                //Loop through all block level nodes (paragraphs and tables) in the body of the section.  
                int nodeCount = srcSection.Body.ChildNodes.Count;
                for (int nodeIndex = 0; nodeIndex < nodeCount; nodeIndex++)
                {
                    Node srcNode = srcSection.Body.ChildNodes[nodeIndex];
                    Node newNode = importer.ImportNode(srcNode, true);
                    dstStory.InsertAfter(newNode, insertAfterNode);
                    insertAfterNode = newNode;
                }
            }


        }

        static void InsertDocument(Node insertAfterNode, Aspose.Words.Document srcDoc)
        {
            // Make sure that the node is either a paragraph or table.  
            if ((!insertAfterNode.NodeType.Equals(NodeType.Paragraph)) &
              (!insertAfterNode.NodeType.Equals(NodeType.Table)))
                throw new ArgumentException("The destination node should be either a paragraph or table.");

            // We will be inserting into the parent of the destination paragraph.  
            CompositeNode dstStory = insertAfterNode.ParentNode;

            // This object will be translating styles and lists during the import.  
            NodeImporter importer = new NodeImporter(srcDoc, insertAfterNode.Document, ImportFormatMode.KeepSourceFormatting);

            // Loop through all sections in the source document.  
            foreach (Aspose.Words.Section srcSection in srcDoc.Sections)
            {
                // Loop through all block level nodes (paragraphs and tables) in the body of the section.  
                foreach (Node srcNode in srcSection.Body)
                {
                    // Let's skip the node if it is a last empty paragraph in a section.  
                    if (srcNode.NodeType.Equals(NodeType.Paragraph))
                    {
                        Aspose.Words.Paragraph para = (Aspose.Words.Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document.  
                    Node newNode = importer.ImportNode(srcNode, true);

                    // Insert new node after the reference node.  
                    dstStory.InsertAfter(newNode, insertAfterNode);
                    insertAfterNode = newNode;
                }
            }
        }
        /// <summary>  
        /// 换行  
        /// </summary>  
        public void InsertLineBreak()
        {
            _WordBuilder.InsertBreak(BreakType.LineBreak);
        }
        /// <summary>  
        /// 换多行  
        /// </summary>  
        /// <param name="nline"></param>  
        public void InsertLineBreak(int nline)
        {
            for (int i = 0; i < nline; i++)
                _WordBuilder.InsertBreak(BreakType.LineBreak);
        }

        #region InsertScoreTable
        public bool InsertScoreTable(bool dishand, bool distab, string handText)
        {
            try
            {


                _WordBuilder.StartTable();//开始画Table  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                //添加Word表格  
                _WordBuilder.InsertCell();
                _WordBuilder.CellFormat.Width = 115.0;
                _WordBuilder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(115);
                _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.None;

                _WordBuilder.StartTable();//开始画Table  
                _WordBuilder.RowFormat.Height = 20.2;
                _WordBuilder.InsertCell();
                _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.Single;
                _WordBuilder.Font.Size = 10.5;
                _WordBuilder.Bold = false;
                _WordBuilder.Write("评卷人");

                _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                _WordBuilder.CellFormat.Width = 50.0;
                _WordBuilder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                _WordBuilder.RowFormat.Height = 20.0;
                _WordBuilder.InsertCell();
                _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.Single;
                _WordBuilder.Font.Size = 10.5;
                _WordBuilder.Bold = false;
                _WordBuilder.Write("得分");
                _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                _WordBuilder.CellFormat.Width = 50.0;
                _WordBuilder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                _WordBuilder.EndRow();
                _WordBuilder.RowFormat.Height = 25.0;
                _WordBuilder.InsertCell();
                _WordBuilder.RowFormat.Height = 25.0;
                _WordBuilder.InsertCell();
                _WordBuilder.EndRow();
                _WordBuilder.EndTable();

                _WordBuilder.InsertCell();
                _WordBuilder.CellFormat.Width = 300.0;
                _WordBuilder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.Auto;
                _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.None;


                _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                _WordBuilder.Font.Size = 11;
                _WordBuilder.Bold = true;
                _WordBuilder.Write(handText);
                _WordBuilder.EndRow();
                _WordBuilder.RowFormat.Height = 28;
                _WordBuilder.EndTable();
                return true;
            }
            catch
            {

                return false;
            }

        }
        #endregion
        #region 插入表格
        public bool InsertTable(System.Data.DataTable dt, bool haveBorder)
        {
            Aspose.Words.Tables.Table table = _WordBuilder.StartTable();//开始画Table  
            ParagraphAlignment paragraphAlignmentValue = _WordBuilder.ParagraphFormat.Alignment;
            _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            //添加Word表格  
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                _WordBuilder.RowFormat.Height = 5;
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    _WordBuilder.InsertCell();
                    _WordBuilder.Font.Size = 10.0;
                    _WordBuilder.Font.Name = "宋体";
                    _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                    _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐  
                    _WordBuilder.CellFormat.Width = 100.0;
                    _WordBuilder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(100);
                    if (haveBorder == true)
                    {
                        //设置外框样式     
                        _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.Single;
                        //样式设置结束     
                    }

                    _WordBuilder.Write(dt.Rows[row][col].ToString());
                }

                _WordBuilder.EndRow();

            }
            _WordBuilder.EndTable();
            _WordBuilder.ParagraphFormat.Alignment = paragraphAlignmentValue;
            table.Alignment = Aspose.Words.Tables.TableAlignment.Center;
            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.Auto;



            return true;
        }
        #endregion


        #region 插入表格
        public bool InsertTable2(System.Data.DataTable dt, bool haveBorder, List<double> widthLst)
        {

            Aspose.Words.Tables.Table table = _WordBuilder.StartTable();//开始画Table  
            ParagraphAlignment paragraphAlignmentValue = _WordBuilder.ParagraphFormat.Alignment;
            _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            //添加Word表格  
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                _WordBuilder.RowFormat.Height = 50;
                //第一列
                _WordBuilder.InsertCell();
                _WordBuilder.Font.Size = 10.0;
                _WordBuilder.Font.Name = "宋体";
                _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐  
                _WordBuilder.CellFormat.Width = widthLst[0];
                _WordBuilder.RowFormat.Height = 10;
                _WordBuilder.CellFormat.PreferredWidth = PreferredWidth.Auto;// Aspose.Words.Tables.PreferredWidth.FromPoints(70);
                if (haveBorder == true)
                {
                    //设置外框样式     
                    _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束     
                }
                _WordBuilder.Write(dt.Rows[row][0].ToString());

                //第二列
                _WordBuilder.InsertCell();
                _WordBuilder.Font.Size = 10.0;
                _WordBuilder.Font.Name = "宋体";
                _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;//水平居中对齐  
                _WordBuilder.CellFormat.Width = widthLst[1];
                _WordBuilder.CellFormat.PreferredWidth = PreferredWidth.Auto;// Aspose.Words.Tables.PreferredWidth.FromPoints(100);
                if (haveBorder == true)
                {
                    //设置外框样式     
                    _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束     
                }
                _WordBuilder.Write(dt.Rows[row][1].ToString());
                //第三列
                _WordBuilder.InsertCell();
                _WordBuilder.Font.Size = 10.0;
                _WordBuilder.Font.Name = "宋体";
                _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left;//水平居中对齐  
                _WordBuilder.CellFormat.Width = widthLst[2];
                _WordBuilder.CellFormat.PreferredWidth = PreferredWidth.Auto;//Aspose.Words.Tables.PreferredWidth.FromPoints(100);
                if (haveBorder == true)
                {
                    //设置外框样式     
                    _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束     
                }
                _WordBuilder.Write(dt.Rows[row][2].ToString());

                //第四列
                _WordBuilder.InsertCell();
                _WordBuilder.Font.Size = 10.0;
                _WordBuilder.Font.Name = "宋体";
                _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐  
                _WordBuilder.CellFormat.Width = widthLst[3];
                _WordBuilder.CellFormat.PreferredWidth = PreferredWidth.Auto;// Aspose.Words.Tables.PreferredWidth.FromPoints(100);
                if (haveBorder == true)
                {
                    //设置外框样式     
                    _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束     
                }
                _WordBuilder.Write(dt.Rows[row][3].ToString());

                /*/第五列图片
                _WordBuilder.InsertCell();
                _WordBuilder.Font.Size = 10.0;
                _WordBuilder.Font.Name = "宋体";
                _WordBuilder.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐  
                _WordBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐  
                _WordBuilder.CellFormat.Width = widthLst[4];
                _WordBuilder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(100);
                if (haveBorder == true)
                {
                    //设置外框样式     
                    _WordBuilder.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束     
                }
                FileInfo file=new FileInfo(dt.Rows[row][4].ToString());
                if (file.Exists)
                {
                    _WordBuilder.InsertImage(dt.Rows[row][4].ToString());
                }*/

                _WordBuilder.EndRow();

            }
            _WordBuilder.EndTable();
            _WordBuilder.ParagraphFormat.Alignment = paragraphAlignmentValue;
            table.Alignment = Aspose.Words.Tables.TableAlignment.Center;
            table.AllowAutoFit = true;
            table.AutoFit(Aspose.Words.Tables.AutoFitBehavior.FixedColumnWidths);



            return true;
        }
        #endregion
        /// <summary>
        /// 追加表格
        /// </summary>
        /// <param name="tableIndex"></param>
        /// <param name="dt"></param>
        /// <param name="widthLst"></param>
        /// <returns></returns>
        public bool AddTable(int tableIndex, System.Data.DataTable dt, List<double> widthLst)
        {
            NodeCollection allTables =_Doc.GetChildNodes(NodeType.Table, true); //拿到所有表格
            Aspose.Words.Tables.Table table = allTables[tableIndex] as Aspose.Words.Tables.Table; //拿到第tableIndex个表格
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                Aspose.Words.Tables.Row r = new Aspose.Words.Tables.Row(_Doc);
                var newRow = CreateRow(dt.Columns.Count, dt.Rows[row].ItemArray, _Doc, widthLst); //创建一行

               table.Rows.Add(newRow); //添加一行
            }

            return true;
        }
    
        Aspose.Words.Tables.Row CreateRow(int columnCount, object[] columnValues, Document doc, List<double> widthLst)
        {
            Aspose.Words.Tables.Row r2 = new Aspose.Words.Tables.Row(doc);
            for (int i = 0; i < columnCount; i++)
            {
                if (columnValues.Length > i)
                {
                    var cell = CreateCell(columnValues[i].ToString(), doc);
                    cell.CellFormat.Width = widthLst[i];
                    cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    if(String.IsNullOrEmpty(columnValues[i].ToString())){
                        cell.CellFormat.VerticalMerge=CellMerge.Previous;
                    }else{
                        cell.CellFormat.VerticalMerge=CellMerge.First;
                    }
                    r2.Cells.Add(cell);
                }
                else
                {
                    var cell = CreateCell("", doc);
                    r2.Cells.Add(cell);
                }

            }
            return r2;

        }
       
        Aspose.Words.Tables.Cell CreateCell(string value, Document doc)
        {
            Aspose.Words.Tables.Cell c1 = new Aspose.Words.Tables.Cell(doc);
            Aspose.Words.Paragraph p = new Paragraph(doc);
            p.AppendChild(new Run(doc, value));
            c1.AppendChild(p);
            return c1;
        }
        public void InsertPagebreak()
        {
            _WordBuilder.InsertBreak(BreakType.PageBreak);

        }
        public void InsertBookMark(string BookMark)
        {
            _WordBuilder.StartBookmark(BookMark);
            _WordBuilder.EndBookmark(BookMark);

        }
        public void GotoBookMark(string strBookMarkName)
        {
            _WordBuilder.MoveToBookmark(strBookMarkName);
        }
        public void ClearBookMark()
        {
            _Doc.Range.Bookmarks.Clear();
        }

        public void ReplaceText(string oleText, string newText)
        {
            //System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(oleText);  
            _Doc.Range.Replace(oleText, newText, false, false);

        }
        private class InsertDocumentAtReplaceHandler : IReplacingCallback
        {
            private string vfilename;
            private int pNum;

            public InsertDocumentAtReplaceHandler(string filename, int _pNum)
            {
                this.vfilename = filename;
                this.pNum = _pNum;
            }
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Document subDoc = new Document(this.vfilename);
                subDoc.FirstSection.Body.FirstParagraph.InsertAfter(new Run(subDoc, this.pNum + "."), null);

                // Insert a document after the paragraph, containing the match text.  
                Node currentNode = e.MatchNode;
                Paragraph para = (Paragraph)e.MatchNode.ParentNode;
                InsertDocument(para, subDoc);
                // Remove the paragraph with the match text.  
                e.MatchNode.Remove();
                e.MatchNode.Range.Delete();



                return ReplaceAction.Skip;
            }
        }

        /// <summary>
        /// 杀掉winword.exe进程
        /// </summary>
        public void killWinWordProcess()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes)
            {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle != "")
                {
                    process.Kill();
                }
            }
        }

         public  void SaveAsPdf(string pdfFile)
        {
            _Doc.Save(pdfFile, SaveFormat.Pdf);
        }
    }
}
