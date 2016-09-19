using System;
using DocumentFormat.OpenXml;
using FSCProLib;

namespace FVMacros
{
    /// <summary>
    /// Represents a conversion between an Excel column and a multi-row (screen) column.
    /// </summary>
    public class ScreenToExcelColumnConversion
    {
        /// <summary>
        /// Ctor.
        /// </summary>
        /// <param name="screenColumnName"></param>
        /// <param name="excelColumnHeader"></param>
        /// <param name="excelColumnWidth"></param>
        /// <param name="excelColumnFormat"></param>
        /// <param name="formatFunc"></param>
        public ScreenToExcelColumnConversion( string screenColumnName, string excelColumnHeader, bool isNumber, int excelColumnWidth = -1, string excelColumnFormat = null, Func<string, string> formatFunc = null )
        {
            ScreenColumnName = screenColumnName;
            ExcelColumnHeader = excelColumnHeader;
            IsNumber = isNumber;
            ExcelColumnWidth = excelColumnWidth;
            ExcelColumnFormat = excelColumnFormat;
            FormatFunc = formatFunc;
        }

        /// <summary>
        /// Gets and sets the index of the cells style that applies to all cells in the column.
        /// Null = use the default style.
        /// </summary>
        public UInt32Value CellStyleIndex { get; set; }

        /// <summary>
        /// Gets the header (name) of this column in Excel.
        /// </summary>
        public string ExcelColumnHeader { get; private set; }

        /// <summary>
        /// Gets the Excel column format string (if any) for this column.
        /// Any Excel format should work but you may need to escape quotes etc.
        /// </summary>
        public string ExcelColumnFormat { get; private set; }

        /// <summary>
        /// Gets and sets the Excel column ref (A, B etc.) for this conversion.
        /// Will need more than one char if more than 26 cols are in use.
        /// </summary>
        public char ExcelColumnRef { get; set; }

        /// <summary>
        /// Get the width to set for this column in Excel (if any).
        /// If the column is narrower than expected, try adding 3 to the width.
        /// </summary>
        public int ExcelColumnWidth { get; private set; }

        /// <summary>
        /// Gets the format function (if any) to call in GetText().
        /// </summary>
        public Func<string, string> FormatFunc { get; private set; }

        /// <summary>
        /// Gets a flag indicating if tis conversion specifies an Excel column format.
        /// </summary>
        public bool HasExcelColumnFormat
        {
            get
            {
                return !string.IsNullOrWhiteSpace( ExcelColumnFormat );
            }
        }

        /// <summary>
        /// Gets a flag indicating if tis conversion specifies an Excel column width.
        /// </summary>
        public bool HasExcelColumnWidth
        {
            get
            {
                return ExcelColumnWidth != -1;
            }
        }

        /// <summary>
        /// Gets a flag that indicates if values in this column are numbers.
        /// </summary>
        public bool IsNumber { get; private set; }

        /// <summary>
        /// Gets the name of the column in the screen map.
        /// </summary>
        public string ScreenColumnName { get; private set; }

        /// <summary>
        /// Gets the text for this column from the screen, passing it through the
        /// format function if one was specified.
        /// </summary>
        /// <param name="screen"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public string GetText( HostScreen screen, int row )
        {
            var text = screen.mappedRowGet( ScreenColumnName, row );

            if ( FormatFunc != null )
            {
                text = FormatFunc( text );
            }

            return text;
        }
    }
}
