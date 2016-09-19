using System.Collections.Generic;
using FSCProLib;

namespace FVMacros
{
    /// <summary>
    /// Class which represents a multi-row on a screen and its conversion to Excel.
    /// </summary>
    public class MultiRow
    {
        /// <summary>
        /// The available column conversions.
        /// </summary>
        private Dictionary<string, ScreenToExcelColumnConversion> Columns = new Dictionary<string, ScreenToExcelColumnConversion>();

        public MultiRow( string mapName, int dataRows, string pageDownKey, string pageUpKey, MoreIndicator moreIndicator )
        {
            MapName = mapName;
            DataRows = dataRows;
            PageDownKey = pageDownKey;
            PageUpKey = pageUpKey;

            MoreIndicator = moreIndicator;
        }

        /// <summary>
        /// Gets the number of rows of DATA in the multi-row.
        /// </summary>
        public int DataRows { get; private set; }

        /// <summary>
        /// Gets the name of the screen map to use.
        /// </summary>
        public string MapName { get; private set; }

        /// <summary>
        /// Gets the MoreIndicator for the screen.
        /// </summary>
        public MoreIndicator MoreIndicator { get; private set; }

        /// <summary>
        /// Gets the page down key for the screen.
        /// </summary>
        public string PageDownKey { get; private set; }

        /// <summary>
        /// Gets the page up key for the screen.
        /// </summary>
        public string PageUpKey { get; private set; }

        /// <summary>
        /// Add a column and its Excel conversion to this multi-row.
        /// </summary>
        /// <param name="column"></param>
        public void AddColumn( ScreenToExcelColumnConversion column )
        {
            Columns.Add( column.ExcelColumnHeader, column );
        }

        /// <summary>
        /// Gets a flag indicating if the screen has more rows of this multi-row available.
        /// </summary>
        /// <param name="screen"></param>
        /// <returns></returns>
        public bool HasMoreRows( HostScreen screen )
        {
            return MoreIndicator.ScreenHasMoreRows( screen );
        }

        /// <summary>
        /// Pages the given screen down.
        /// </summary>
        /// <param name="screen"></param>
        public void PageDown( HostScreen screen )
        {
            screen.putCommand( PageDownKey );
        }

        /// <summary>
        /// Pages the given screen up.
        /// </summary>
        /// <param name="screen"></param>
        public void PageUp( HostScreen screen )
        {
            screen.putCommand( PageUpKey );
        }

        /// <summary>
        /// Trys to get a column conversion for the given Excel column header.
        /// </summary>
        /// <param name="excelColumnHeader"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public bool TryGetColumn( string excelColumnHeader, out ScreenToExcelColumnConversion column )
        {
            column = null;

            return Columns.TryGetValue( excelColumnHeader, out column );
        }
    }
}
