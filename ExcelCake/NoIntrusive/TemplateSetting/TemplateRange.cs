using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    [Serializable]
    internal abstract class TemplateSettingBase
    {
        internal string Content { set; get; }

        [NonSerialized]
        private ExcelRangeBase _CurrentCell;
        /// <summary>
        /// 配置所在单元格
        /// </summary>

        internal ExcelRangeBase CurrentCell { set { this._CurrentCell = value; } get {return this._CurrentCell; } }

        [NonSerialized]
        private ExcelRangeBase _ScopeRange;
        /// <summary>
        /// 作用范围
        /// </summary>
        internal ExcelRangeBase ScopeRange { set { this._ScopeRange = value; } get { return this._ScopeRange; } }

        protected abstract void AnalyseSetting();
    }
}
