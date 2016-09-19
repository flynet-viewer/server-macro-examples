using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FVMacros
{
    /// <summary>
    /// Class that manages a spreadsheets shared strings table
    /// </summary>
    public class SharedStrings
    {
        /// <summary>
        /// Lookup to conver known strings to their index.
        /// </summary>
        private Dictionary<string, int> LookupByText = new Dictionary<string, int>();

        /// <summary>
        /// Lookup to convert an index to its shared string.
        /// </summary>
        private List<string> LookupByIndex = new List<string>();

        /// <summary>
        /// The spreadsheets shared string table.
        /// </summary>
        private SharedStringTable Table;

        /// <summary>
        /// Ctor. Reads the shared string table into the lookups.
        /// </summary>
        /// <param name="workbookPart"></param>
        public SharedStrings( WorkbookPart workbookPart )
        {
            var sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            if ( sharedStringPart == null )
            {
                throw new OpenXmlPackageException( "No SharedStringTablePart Found." );
            }

            if ( sharedStringPart.SharedStringTable == null )
            {
                // Spreadsheet has no shared strings table - add one.
                Table = new SharedStringTable();

                sharedStringPart.SharedStringTable = Table;
            }
            else
            {
                // Read the existing strings into the lookups.
                Table = sharedStringPart.SharedStringTable;

                int index = 0;

                foreach ( var item in Table.Elements<SharedStringItem>() )
                {
                    LookupByIndex.Add( item.InnerText );

                    LookupByText.Add( item.InnerText, index++ );
                }
            }
        }

        /// <summary>
        /// Given a string, get its index in the shared string table adding it if needed.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public int GetIndex( string text )
        {
            int index;

            if ( !LookupByText.TryGetValue( text, out index ) )
            {
                // String not in table - add it and update the lookups.
                Table.AppendChild( new SharedStringItem( new Text( text ) ) );

                Table.Save();

                LookupByIndex.Add( text );

                index = LookupByIndex.Count - 1;

                LookupByText.Add( text, index );
            }

            return index;
        }

        /// <summary>
        /// Trys to get the text for the given index.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public bool TryGetText( int index, out string text )
        {
            bool found = false;

            text = string.Empty;

            if ( index >= 0 && index < LookupByIndex.Count )
            {
                text = LookupByIndex[ index ];

                found = true;
            }

            return found;
        }
    }
}
