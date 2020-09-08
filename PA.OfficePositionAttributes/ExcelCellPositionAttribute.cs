using System;
using System.Collections.Generic;
using System.Text;

namespace OfficePositionAttributes
{
    //
    // 概要:
    //     Excleのセルの行と列を表します。
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class ExcelCellPositionAttribute: Attribute
    {
        
        public ExcelCellPositionAttribute()
        {
        }
        //
        // 概要:
        //     行の 1 から始まる順序を設定します。
        //
        // 戻り値:
        //     列の順序。
        public int Row { get; set; }
        
        //
        // 概要:
        //     列の 1 から始まる順序を設定します。
        //
        // 戻り値:
        //     列の順序。
        public int Col { get; set; }
    }
}
