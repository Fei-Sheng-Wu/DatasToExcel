using System;

namespace DatasToExcel
{
    internal static class Internal
    {
        /// <summary>
        /// Get column name from column number.
        /// <para>https://stackoverflow.com/a/182924/12949439</para>
        /// </summary>
        /// <param name="columnNumber">The column number from 1.</param>
        /// <returns>The column name in letters</returns>
        internal static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
