using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FVMacros
{
    /// <summary>
    ///  A class to manage adding number formats to a spreadsheet.
    /// </summary>
    public class NumberingFormats
    {
        /// <summary>
        /// Next available index for custom formats. Lowest allowable value is 164.
        /// </summary>
        private uint NextCustomIndex = 164;

        /// <summary>
        /// The Stylesheet for the spreadsheet.
        /// </summary>
        private Stylesheet Stylesheet;

        /// <summary>
        /// Lookup of available formats by their value.
        /// </summary>
        private Dictionary<string, uint> LookupByFormat = new Dictionary<string, uint>();

        /// <summary>
        /// Ctor. Reads any formats in the stylesheet into the lookup.
        /// </summary>
        /// <param name="stylesheet"></param>
        public NumberingFormats( Stylesheet stylesheet )
        {
            Stylesheet = stylesheet;

            var numberingFormats = Stylesheet.NumberingFormats;

            if ( numberingFormats == null )
            {
                // Existing XML has no element to hold the formats. Add one.
                Stylesheet.NumberingFormats = new DocumentFormat.OpenXml.Spreadsheet.NumberingFormats();

                Stylesheet.Save();
            }
            else
            {
                // Loop through all existing formats adding them to the lookup.
                foreach ( var numberingFormat in numberingFormats.OfType<NumberingFormat>() )
                {
                    var index = numberingFormat.NumberFormatId.Value;

                    LookupByFormat.Add( numberingFormat.FormatCode, index );

                    if ( index > NextCustomIndex )
                    {
                        // Custom format, update NextCustomIndex.
                        NextCustomIndex = index + 1;
                    }
                }
            }
        }

        /// <summary>
        /// Given a format string get its index adding it to the stylesheet if it's not already in it.
        /// </summary>
        /// <param name="format"></param>
        /// <returns></returns>
        /// <remarks> A complete implementation would look for formats that match the std ones and return the right index. </remarks>
        public uint GetIndex( string format )
        {
            uint index;

            if ( !LookupByFormat.TryGetValue( format, out index ) )
            {
                // New format so add it to the stylesheet.
                index = NextCustomIndex++;

                Stylesheet.NumberingFormats.AppendChild( new NumberingFormat() { FormatCode = format, NumberFormatId = index } );

                Stylesheet.Save();
            }

            return index;
        }
    }
}
