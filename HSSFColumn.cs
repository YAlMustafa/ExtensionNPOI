using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace ExtensionNPOI
{
    public class HSSFColumn : IColumn
    {
        public List<ICell> Cells { get; set; }

        public ISheet sheet { get; }

        public int Index { get; }

        public string AlphabeticIndex { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index">0-Based Index</param>
        /// <returns></returns>
        public virtual ICell GetCell(int index)
        {
            return Cells[index];
        }
    }
}
