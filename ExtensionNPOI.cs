using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;

namespace ExtensionNPOI
{
    public static class ExtensionNPOI
    {
        #region ISheet
        /// <summary>
        /// Get the specified cell.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">0-Based Index</param>
        /// <param name="column">Alphabetic Index</param>
        /// <returns></returns>
        public static ICell GetCell(this ISheet sheet, int row, string column)
        {
            return sheet.GetRow(row).GetCell(ConvertAlphabetToNumaric0(column));
        }
        /// <summary>
        /// Get the specified cell.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">0-Based Index</param>
        /// <param name="column">0-Based Index</param>
        /// <returns></returns>
        public static ICell GetCell(this ISheet sheet, int row, int column)
        {
            return sheet.GetRow(row).GetCell(column);
        }

        /// <summary>
        /// Returns the specified column
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index">0-Based Index</param>
        /// <returns></returns>
        public static IColumn GetColumn(this ISheet sheet, int index)
        {
            // Create a list of cells
            List<ICell> cells = new List<ICell>();
            for (int i = 0; i <= sheet.LastRowNum; i++) cells.Add(sheet.GetRow(i).GetCell(index));

            // Return the created list
            if (sheet.GetType() == typeof(XSSFSheet))
                return new XSSFColumn() { Cells = cells };
            else if (sheet.GetType() == typeof(HSSFSheet))
                return new HSSFColumn() { Cells = cells };
            else return null;
        }
        /// <summary>
        /// Returns the specified column
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index">Alphabetic Index</param>
        /// <returns></returns>
        public static IColumn GetColumn(this ISheet sheet, string index)
        {
            return GetColumn(sheet, ConvertAlphabetToNumaric0(index));
        }
        #endregion

        #region XSSFSheet
        /// <summary>
        /// Get the specified cell.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">0-Based Index</param>
        /// <param name="column">Alphabetic Index</param>
        /// <returns></returns>
        public static ICell GetCell(this XSSFSheet sheet, int row, string column)
        {
            return sheet.GetRow(row).GetCell(ConvertAlphabetToNumaric0(column));
        }
        /// <summary>
        /// Get the specified cell.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">0-Based Index</param>
        /// <param name="column">0-Based Index</param>
        /// <returns></returns>
        public static ICell GetCell(this XSSFSheet sheet, int row, int column)
        {
            return sheet.GetRow(row).GetCell(column);
        }

        /// <summary>
        /// Returns the specified column
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index">0-Based Index</param>
        /// <returns></returns>
        public static XSSFColumn GetColumn(this XSSFSheet sheet, int index)
        {
            // Create a list of cells
            List<ICell> cells = new List<ICell>();
            for (int i = 0; i <= sheet.LastRowNum; i++) cells.Add(sheet.GetRow(i).GetCell(index));

            // Return the created list
            return new XSSFColumn() { Cells = cells };
        }
        /// <summary>
        /// Returns the specified column
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index">Alphabetic Index</param>
        /// <returns></returns>
        public static XSSFColumn GetColumn(this XSSFSheet sheet, string index)
        {
            return GetColumn(sheet, ConvertAlphabetToNumaric0(index));
        }
        #endregion

        #region HSSFSheet
        /// <summary>
        /// Get the specified cell.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">0-Based Index</param>
        /// <param name="column">Alphabetic Index</param>
        /// <returns></returns>
        public static ICell GetCell(this HSSFSheet sheet, int row, string column)
        {
            return sheet.GetRow(row).GetCell(ConvertAlphabetToNumaric0(column));
        }
        /// <summary>
        /// Get the specified cell.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">0-Based Index</param>
        /// <param name="column">0-Based Index</param>
        /// <returns></returns>
        public static ICell GetCell(this HSSFSheet sheet, int row, int column)
        {
            return sheet.GetRow(row).GetCell(column);
        }

        /// <summary>
        /// Returns the specified column
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index">0-Based Index</param>
        /// <returns></returns>
        public static HSSFColumn GetColumn(this HSSFSheet sheet, int index)
        {
            // Create a list of cells
            List<ICell> cells = new List<ICell>();
            for (int i = 0; i <= sheet.LastRowNum; i++) cells.Add(sheet.GetRow(i).GetCell(index));

            // Return the created list
            return new HSSFColumn() { Cells = cells };
        }
        /// <summary>
        /// Returns the specified column
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index">Alphabetic Index</param>
        /// <returns></returns>
        public static HSSFColumn GetColumn(this HSSFSheet sheet, string index)
        {
            return GetColumn(sheet, ConvertAlphabetToNumaric0(index));
        }
        #endregion


        #region Helper Methods
        /// <summary>
        /// Convert alphabetic index to one-based numaric index.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int ConvertAlphabetToNumaric(string str)
        {
            int index = 0;
            string letters = " abcdefghijklmnopqrstuvwxyz";
            string letters_c = letters.ToUpper();

            int n = str.Length - 1;
            foreach (char c in str)
            {
                // Get the index of the character
                int d = letters.IndexOf(c);
                if (d == -1) d = letters_c.IndexOf(c);
                // Add to the result
                index += d * Convert.ToInt32(Math.Pow(26, n));
                // Decrease to the next digit
                n--;
            }

            return index;
        }
        /// <summary>
        /// Convert alphabetic index to zero-based numaric index.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int ConvertAlphabetToNumaric0(string str) => ConvertAlphabetToNumaric(str) - 1;
        #endregion
    }
}
