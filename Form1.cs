using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NPOITest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void tsbOpen_Click(object sender, EventArgs e)
        {
            lblStatus.Text = string.Empty;

            OpenFileDialog dia = new OpenFileDialog();
            dia.Filter = "Word2007文档(*.docx)|*.docx";

            if (DialogResult.OK == dia.ShowDialog())
            {
                lblStatus.Text = "请稍候，处理中...";

                //NPOITestMerge(dia.FileName);

                //NPOITestBookMark(dia.FileName);

                //NPOITestColumn(dia.FileName);

                NPOITestFillData(dia.FileName);

                NPOITestColumn(dia.FileName);

                #region //
                //XWPFDocument doc;
                //using (FileStream stream = File.OpenRead(dia.FileName))
                //{
                //    doc = new XWPFDocument(stream);
                //}
                //FileStream file = new FileStream(dia.FileName, FileMode.Create, FileAccess.Write);
                //doc.Write(file);
                //file.Close();
                #endregion

                lblStatus.Text = "DONE";
            }
        }

        private void NPOITestFillData(string file)
        {
            #region 读取Word
            XWPFDocument doc;
            using (FileStream fileread = File.OpenRead(file))
            {
                doc = new XWPFDocument(fileread);
            }
            #endregion

            List<string>[] data = new List<string>[3];
            #region 组织填充数据
            List<string> a = new List<string>();
            List<string> b = new List<string>();
            List<string> c = new List<string>();
            a.Add("1.1");
            a.Add("1.2");
            a.Add("1.3");
            a.Add("1.4");
            a.Add("1.5");
            a.Add("1.6");
            a.Add("1.7");
            b.Add("2.1");
            b.Add("2.2");
            b.Add("2.3");
            b.Add("2.4");
            b.Add("2.5");
            b.Add("2.6");
            c.Add("3.1");
            c.Add("3.2");
            c.Add("3.3");
            c.Add("3.4");
            c.Add("3.5");
            c.Add("3.6");
            c.Add("3.7");
            c.Add("3.8");
            c.Add("3.9");
            data[0] = a;
            data[1] = b;
            data[2] = c;
            #endregion

            WordTable wt = new WordTable(data);
            
            wt.CaptionLineCount = 2;    //标题行数

            XWPFTable table = LocationTable(doc, "本年发生的非同一控制下企业合并情况");

            if (wt.ObjectData != null)
            {
                for (int i = 0; i < wt.ObjectData.Length + wt.CaptionLineCount; i++)
                {
                    if (i >= wt.CaptionLineCount)
                    {
                        XWPFTableRow row;
                        string[] rowdata = wt.ObjectData[i - wt.CaptionLineCount].ToArray<string>();
                        if (i < table.Rows.Count)
                        {
                            row = table.GetRow(i);
                            row.GetCTRow().AddNewTrPr().AddNewTrHeight().val = 397;
                            row.GetCTRow().trPr.GetTrHeightArray(0);
                            for (int n = 0; n < rowdata.Length; n++)
                            {
                                XWPFTableCell cell = row.GetCell(n);
                                //模板中的单元格少时，接收数据将会部分丢失
                                //也可以在下边if后添加else在该行后补充单元格
                                //按接收数据循环，所以单元格数多于接收数据时不需要另做处理，该行后边的部分单元格无法补填充
                                if (cell != null)
                                {
                                    //SetText是追加内容，所以要先删除单元格内容（删除单元格内所有段落）再写入
                                    for (int p = 0; p < cell.Paragraphs.Count; p++)
                                    {
                                        cell.RemoveParagraph(p);
                                    }
                                    for (int t = 0; t < cell.Tables.Count; t++)
                                    {
                                        //表格删除
                                        //cell.RemoveTable(t);
                                    }
                                    cell.SetText(rowdata[n]);
                                }
                            }
                        }
                        else
                        {
                            //添加新行
                            //row = table.InsertNewTableRow(table.Rows.Count - 1);

                            row = new XWPFTableRow(new CT_Row(), table);
                            row.GetCTRow().AddNewTrPr().AddNewTrHeight().val = 100;
                            table.AddRow(row);

                            for (int n = 0; n < rowdata.Length; n++)
                            {
                                XWPFTableCell cell = row.CreateCell();
                                CT_Tc tc = cell.GetCTTc();
                                CT_TcPr pr = tc.AddNewTcPr();
                                tc.GetPList()[0].AddNewR().AddNewT().Value = rowdata[n];
                            }
                        }
                    }
                }
            }

            #region 保存Word
            FileStream filewrite = new FileStream(file, FileMode.Create, FileAccess.Write);
            try
            {
                doc.Write(filewrite);
            }
            catch (Exception x)
            {
                lblStatus.Text = x.Message;
            }
            finally
            {
                filewrite.Close();
            }
            #endregion
        }

        /// <summary>
        /// 列测试
        /// 未完成
        /// </summary>
        /// <param name="p"></param>
        private void NPOITestColumn(string file)
        {
            #region 读取Word
            XWPFDocument doc;
            using (FileStream fileread = File.OpenRead(file))
            {
                doc = new XWPFDocument(fileread);
            }
            #endregion

            doc.Tables[0].GetRow(0).GetTableCells()[0].GetCTTc().tcPr.AddNewTcW().w = "100";
            //doc.Tables[0].GetRow(0).

            #region 保存Word
            FileStream filewrite = new FileStream(file, FileMode.Create, FileAccess.Write);
            try
            {
                doc.Write(filewrite);
            }
            catch (Exception x)
            {
                lblStatus.Text = x.Message;
            }
            finally
            {
                filewrite.Close();
            }
            #endregion
        }

        /// <summary>
        /// 测试书签
        /// 不知道书签咋用。用其他方法查找指定表
        /// </summary>
        /// <param name="file"></param>
        private void NPOITestBookMark(string file)
        {
            #region 读取Word
            XWPFDocument doc;
            using (FileStream fileread = File.OpenRead(file))
            {
                doc = new XWPFDocument(fileread);
            }
            #endregion

            //XWPFTable table = LocationTable(doc, "本年发生的非同一控制下企业合并情况");
            XWPFTable table = LocationTableByBookMark(doc, "本年发生的非同一控制下企业合并情况");

            //CT_Bookmark bm = new CT_Bookmark();
            //CT_P p = doc.Document.body.AddNewP();

            //List<ParagraphItemsChoiceType> x = new List<ParagraphItemsChoiceType>();
            //x = p.ItemsElementName;
            //CT_TrPr tp = new CT_TrPr();

            //CT_OnOff oo = new CT_OnOff();
        }

        /// <summary>
        /// 按表名，从word文档中取出相应表
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="name">表名</param>
        /// <returns></returns>
        private XWPFTable LocationTable(XWPFDocument doc, string name)
        {
            try
            {
                for (int i = 0; i < doc.BodyElements.Count - 1; i++)
                {
                    if (doc.BodyElements[i].GetType() == typeof(XWPFParagraph) && doc.BodyElements[i + 1].GetType() == typeof(XWPFTable))
                    {
                        string pg = ((XWPFParagraph)doc.BodyElements[i]).Text;
                        if (pg.IndexOf('）') > 0)
                        {
                            pg = pg.Substring(pg.IndexOf('）') + 1);
                        }
                        else if (pg.IndexOf(')') > 0)
                        {
                            pg = pg.Substring(pg.IndexOf(')') + 1);
                        }

                        if (pg == name)
                        {
                            return ((XWPFTable)doc.BodyElements[i + 1]);
                        }
                    }
                }
                return null;
            }
            catch (Exception x)
            {
                lblStatus.Text = x.Message;
                return null;
            }
        }

        /// <summary>
        /// 按表名，从word文档中取出相应表
        /// 按书签（取书签所在段落下的表格）
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="name">表名</param>
        /// <returns></returns>
        private XWPFTable LocationTableByBookMark(XWPFDocument doc, string bm)
        {
            try
            {
                for (int i = 0; i < doc.Document.body.Items.Count - 1; i++)
                {
                    if (doc.Document.body.Items[i].GetType() == typeof(CT_P) && doc.Document.body.Items[i + 1].GetType() == typeof(CT_Tbl))
                    {
                        CT_P p = (CT_P)doc.Document.body.Items[i];
                        if (p != null /*&& p.GetBookmarkStartArray(0).name == bm*/)
                        {
                            //段落中有一个书签符合报表编码，则返回该段落下的表格
                            foreach (CT_Bookmark e in p.GetBookmarkStartList())
                            {
                                if (e.name == bm)
                                {
                                    return ((XWPFTable)doc.BodyElements[i + 1]);
                                }
                            }
                        }
                    }
                }
                ((CT_P)doc.Document.body.Items[0]).GetBookmarkStartArray(0);
                doc.Document.body.GetPArray(0);
                return null;
            }
            catch (Exception x)
            {
                lblStatus.Text = x.Message;
                return null;
            }
        }

        /// <summary>
        /// 测试合并单元格
        /// </summary>
        /// <param name="file"></param>
        private void NPOITestMerge(string file)
        {
            #region
            XWPFDocument doc;
            using (FileStream fileread = File.OpenRead(file))
            {
                doc = new XWPFDocument(fileread);
            }
            #endregion

            foreach (XWPFTable table in doc.Tables)
            {
                XWPFTableRow row = table.GetRow(0);

                #region 老版NPOI(2.0)行合并【所有行】
                for (int c = 0; c < row.GetTableCells().Count; c++)
                {
                    XWPFTableCell cell = row.GetTableCells()[c];
                    CT_Tc tc = cell.GetCTTc();
                    if (tc.tcPr == null)
                    {
                        tc.AddNewTcPr();
                    }
                    if (c == 0)
                    {
                        tc.tcPr.AddNewHMerge().val = ST_Merge.restart;
                    }
                    else
                    {
                        tc.tcPr.AddNewHMerge().val = ST_Merge.@continue;
                    }
                }
                #endregion

                #region //行合并【0～2列】（有错？）
                //row.MergeCells(0, 2);
                #endregion

                #region //列合并【0～2行】（有错？）
                //for (int r = 0; r < 2; r++)
                //{
                //    XWPFTableCell cell = table.GetRow(r).GetTableCells()[0];
                //    CT_Tc tc = cell.GetCTTc();
                //    if (tc.tcPr == null)
                //    {
                //        tc.AddNewTcPr();
                //    }
                //    if (r == 0)
                //    {
                //        tc.tcPr.AddNewVMerge().val = ST_Merge.restart;
                //    }
                //    else
                //    {
                //        tc.tcPr.AddNewVMerge().val = ST_Merge.@continue;
                //    }
                //}
                #endregion
            }

            #region 保存Word
            FileStream filewrite = new FileStream(file, FileMode.Create, FileAccess.Write);
            try
            {
                doc.Write(filewrite);
            }
            catch (Exception x)
            {
                lblStatus.Text = x.Message;
            }
            finally
            {
                filewrite.Close();
            }
            #endregion
        }

        private void tsbExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

