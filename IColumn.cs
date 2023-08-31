using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace ExtensionNPOI
{
    public interface IColumn
    {
        /// <summary>
        /// 0-Based Index
        /// </summary>
        int Index { get; }
        string AlphabeticIndex { get; }
        ISheet sheet { get; }
        List<ICell> Cells { get;}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index">0-Based Index</param>
        /// <returns></returns>
        ICell GetCell(int index);
    }
}
