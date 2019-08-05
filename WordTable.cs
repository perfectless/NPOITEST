using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOITest
{
    public class WordTable
    {
        /// <summary>
        /// 标题行数
        /// </summary>
        public int CaptionLineCount
        {
            get;
            set;
        }

        List<string>[] data = null;
        /// <summary>
        /// 表格数据（不包含标题）
        /// 数据类型为不等长二维数据
        /// </summary>
        public List<string>[] ObjectData
        {
            get { return data; }
        }

        /// <summary>
        /// 浮动行号集合
        /// </summary>
        public List<int>[] FloatLine
        {
            get;
            set;
        }

        string[] hMergeInfo;
        /// <summary>
        /// 水平合并信息
        /// </summary>
        public string[] HMergeInfo
        {
            get { return hMergeInfo; }
            set { hMergeInfo = value; }
        }

        string[] vMergeInfo;
        /// <summary>
        /// 垂直合并信息
        /// </summary>
        public string[] VMergeInfo
        {
            get { return vMergeInfo; }
            set { vMergeInfo = value; }
        }

        Hashtable cellMergeInfo;
        /// <summary>
        /// 单元格合并信息
        /// Hashtable.Key=单元格坐标
        /// Hashtable.Value=CellMergeEnum
        /// </summary>
        public Hashtable CellMergeInfo
        {
            get { return cellMergeInfo; }
            set { cellMergeInfo = value; }
        }

        public WordTable(List<string>[] _data)
        {
            data = _data;
        }
    }

    /// <summary>
    /// 单元格合并枚举值
    /// </summary>
    public enum CellMergeEnum
    {
        上=1,
        下=2,
        左=3,
        右=4,
        无=5,

        /// <summary>
        /// 水平合并
        /// </summary>
        H=6,
        /// <summary>
        /// 垂直合并
        /// </summary>
        V=7,
        /// <summary>
        /// 双向合并
        /// </summary>
        HV=8,
        /// <summary>
        /// 无合并
        /// </summary>
        NONE=9
    };
}
