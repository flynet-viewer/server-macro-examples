using FSCProLib;

namespace FVMacros
{
    /// <summary>
    /// Class that represents the more indicator for a multi-row.
    /// </summary>
    public class MoreIndicator
    {
        /// <summary>
        /// Ctor.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="text"></param>
        public MoreIndicator( int row, int column, string text )
        {
            Row = row;
            Column = column;
            Length = text.Length;
            Text = text;
        }

        /// <summary>
        /// Gets the column at which the excpected text starts.
        /// </summary>
        public int Column { get; private set; }

        /// <summary>
        /// Gets the length of the expected text on the screen.
        /// </summary>
        public int Length { get; private set; }

        /// <summary>
        /// Gets the row on which the expected text appears.
        /// </summary>
        public int Row { get; private set; }

        /// <summary>
        /// Gets the text which appears on the screen if more rows are avaialble.
        /// </summary>
        public string Text { get; private set; }

        /// <summary>
        /// Returns a flag indicating if the given screen has more rows available.
        /// </summary>
        /// <param name="screen"></param>
        /// <returns></returns>
        public bool ScreenHasMoreRows( HostScreen screen )
        {
            return screen.getText( Row, Column, Length ).Trim() == Text;
        }
    }
}
