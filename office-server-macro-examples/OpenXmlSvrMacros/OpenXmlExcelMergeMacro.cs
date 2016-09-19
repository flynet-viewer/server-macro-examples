using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FVMacros
{
    /// <summary>
    /// Server macro that merges data from the Account Transtactions screen into an Excel spreadsheet.
    /// </summary>
    public sealed class OpenXmlExcelMergeMacro : ServerMacro
    {
        private NumberingFormats NumberingFormats;

        private SharedStrings SharedStrings;

        public OpenXmlExcelMergeMacro() { }

        public OpenXmlExcelMergeMacro( MacroHost host ) : base( host ) { }

        public override string name
        {
            get
            {
                return "OpenXmlExcelMergeMacro";
            }
        }

        public override string description
        {
            get
            {
                return "Open Xml Excel Merge Macro";
            }
        }

        public override string runFromScreenID
        {
            get
            {
                return "AcctTrans";
            }
        }

        public override MacroRunResult run()
        {
            try
            {
                // Get the .xslsx from the user.
                var result = GetClientFile( "Excel.xlsx", "Please Upload a Suitable Excel Template", "Template Upload", false );

                if ( !result )
                {
                    this.ClientMessageBox( "No Template Uploaded", "Word Merge Macro", MacroMessageType.info, MacroButtons.ok );

                    return MacroRunResult.ok;
                }

                if ( !File.Exists( xferReceiveFile ) )
                {
                    this.ClientMessageBox( "Error Opening Uploaded Template - File Not Found", "Error", MacroMessageType.error, MacroButtons.ok );

                    return MacroRunResult.error;
                }

                // TODO: Check that the file is an Excel file.

                // New filename to avoid overwriting the template.
                var pathAndNewFilename = Path.Combine( Path.GetDirectoryName( xferReceiveFile ), string.Format( "{0}_Merged{1}", Path.GetFileNameWithoutExtension( xferReceiveName ), Path.GetExtension( xferReceiveName ) ) );

                var mergedDoc = new FileInfo( pathAndNewFilename );

                if ( File.Exists( mergedDoc.FullName ) )
                {
                    File.Delete( mergedDoc.FullName );
                }

                File.Copy( xferReceiveFile, mergedDoc.FullName );

                using ( var doc = SpreadsheetDocument.Open( mergedDoc.FullName, true ) )
                {
                    var workbookPart = doc.WorkbookPart;

                    if ( workbookPart == null )
                    {
                        this.ClientMessageBox( "Error Opening Uploaded Template - No Workbook Part Found", "Error", MacroMessageType.error, MacroButtons.ok );

                        return MacroRunResult.error;
                    }

                    var worksheetPart = doc.WorkbookPart.WorksheetParts.FirstOrDefault();

                    if ( worksheetPart == null || worksheetPart.Worksheet == null )
                    {
                        this.ClientMessageBox( "Error Opening Uploaded Template - No Worksheet Part or Worksheet Found", "Error", MacroMessageType.error, MacroButtons.ok );

                        return MacroRunResult.error;
                    }

                    var firstRow = worksheetPart.Worksheet.Descendants<Row>().FirstOrDefault();

                    if ( firstRow == null )
                    {
                        this.ClientMessageBox( "Error Opening Uploaded Template - No First (Header) Row Found", "Error", MacroMessageType.error, MacroButtons.ok );

                        return MacroRunResult.error;
                    }

                    // Populate shared strings and number formats.
                    SharedStrings = new SharedStrings( workbookPart );
                    NumberingFormats = new NumberingFormats( workbookPart.WorkbookStylesPart.Stylesheet );

                    // Get the object that represents the multi-row on the screen.
                    var multiRow = GetMultiRow();

                    // Need to remember the mapping between screen column and Excel column.
                    // They could be in a different order.
                    List<ScreenToExcelColumnConversion> includedColumns = new List<ScreenToExcelColumnConversion>();

                    // Note: this will fails if you have more than 26 columns.
                    char excelColumnRef = 'A';

                    // TODO: There could already be a Columns element.
                    var columns = new Columns();

                    // Go through the headers one at a time looking for matching columns in the multi-row. 
                    foreach ( var cell in firstRow.Descendants<Cell>() )
                    {
                        string columnName;

                        // The text of the cell might be in the shared strings.
                        if ( cell.DataType == CellValues.SharedString )
                        {
                            var index = int.Parse( cell.CellValue.InnerText );

                            SharedStrings.TryGetText( index, out columnName );
                        }
                        else
                        {
                            columnName = cell.CellValue.Text;
                        }

                        if ( string.IsNullOrWhiteSpace( columnName ) )
                        {
                            // 1st empty column = no more columns.
                            break;
                        }

                        ScreenToExcelColumnConversion includedColumn = null;

                        if ( multiRow.TryGetColumn( columnName, out includedColumn ) )
                        {
                            // This column in Excel is a match for one in the multi-row.
                            includedColumn.ExcelColumnRef = excelColumnRef;

                            if ( includedColumn.HasExcelColumnFormat )
                            {
                                // Create a style for the cell based on any exisitng style but setting the number format.
                                includedColumn.CellStyleIndex = CreateCellStyle( workbookPart.WorkbookStylesPart.Stylesheet, cell.StyleIndex, includedColumn.ExcelColumnFormat );
                            }

                            if ( includedColumn.HasExcelColumnWidth )
                            {
                                // Add a column to the spreadsheet which sets the column width.
                                var columnIndex = (uint) excelColumnRef - 'A' + 1;

                                columns.Append( new Column() { CustomWidth = true, Min = columnIndex, Max = columnIndex, Width = includedColumn.ExcelColumnWidth } );
                            }

                            includedColumns.Add( includedColumn );
                        }

                        // Moving on to the next column.
                        ++excelColumnRef;
                    }

                    if ( columns.Any() )
                    {
                        // Order is important here, insert the Columns element AFTER SheetFormatProperties.
                        // A crash here probably indicates that the template already had a Columns element.
                        worksheetPart.Worksheet.InsertAfter( columns, worksheetPart.Worksheet.GetFirstChild<SheetFormatProperties>() );
                    }

                    if ( includedColumns.Any() )
                    {
                        // Must use the multi-row's screen map to access the columns by name.
                        // If you wanted to interact more with the screen after this, you'd need to switch the map back.
                        oScreen.activeMapName = multiRow.MapName;

                        // Start with the 2nd row of the spreadsheet.
                        uint sheetRow = 2;

                        // Count the pages so we can page back up.
                        // Note: if the user has already paged around the list, then the spreadsheet will only contain
                        //       data from the current location onward. To be sure of getting all rows, page up untill
                        //       there is no change to the multi-row.
                        int pages = 0;

                        bool finished = false;

                        do
                        {
                            // Loop through the rows on the current page.
                            for ( var row = 0; row < multiRow.DataRows; ++row )
                            {
                                // Excell columns start at 1.
                                int columnIndex = 1;

                                foreach ( var includedColumn in includedColumns )
                                {
                                    // Get text from the screen and put it in the spreadsheet.
                                    InsertCellAndText( worksheetPart.Worksheet, sheetRow, includedColumn, includedColumn.GetText( oScreen, row ) );

                                    ++columnIndex;
                                }

                                ++sheetRow;
                            }

                            // Is there another page?
                            if ( multiRow.HasMoreRows( oScreen ) )
                            {
                                // Yes, page down and repeat.
                                multiRow.PageDown( oScreen );

                                ++pages;
                            }
                            else
                            {
                                // No, we're done.
                                finished = true;
                            }

                        } while ( !finished );

                        if ( pages > 0 )
                        {
                            for ( int i = 0; i < pages; ++i )
                            {
                                // Page back up to where we were.
                                multiRow.PageUp( oScreen );
                            }
                        }
                    }

                    //worksheetPart.Worksheet.Save();

                    //doc.WorkbookPart.Workbook.Save();

                    // Save the changes.
                    doc.Save();
                }

                // Let the user save the file on their machine.
                var saveResult = PutClientFile( pathAndNewFilename, "Save Merged File", "Template Merged" );

                if ( saveResult == MacroDialogRC.ok )
                {
                    // Clear up.
                    File.Delete( xferReceiveFile );

                    File.Delete( pathAndNewFilename );
                }
            }
            catch ( Exception ex )
            {
                this.ClientMessageBox( ex.Message, "Exception Caught", MacroMessageType.error, MacroButtons.ok );
            }

            return MacroRunResult.ok;
        }

        /// <summary>
        /// Create a cell format based on the one referenced by styleIndex that sets the number format of a cell to the given format.
        /// </summary>
        /// <param name="stylesheet"></param>
        /// <param name="styleIndex"></param>
        /// <param name="excelColumnFormat"></param>
        /// <returns></returns>
        private UInt32Value CreateCellStyle( Stylesheet stylesheet, UInt32Value styleIndex, string excelColumnFormat )
        {
            CellFormat cellFormat;

            if ( styleIndex != null && styleIndex.HasValue )
            {
                // Get the existing format and clone it.
                var existingFormat = stylesheet.CellFormats.ToList()[ (int) styleIndex.Value ];

                cellFormat = new CellFormat( existingFormat.OuterXml );
            }
            else
            {
                // Just create a new format.
                cellFormat = new CellFormat();
            }

            // This will add the format to the spreadsheet (if not already there) and give us its index.
            var index = NumberingFormats.GetIndex( excelColumnFormat );

            cellFormat.ApplyNumberFormat = true;
            cellFormat.NumberFormatId = index;

            // Add the format to the spreadsheet.
            stylesheet.CellFormats.Append( cellFormat );

            var newStyleIndex = stylesheet.CellFormats.Count++;

            // Save the changes to the stylesheet.
            stylesheet.Save();

            return (uint) newStyleIndex;
        }

        /// <summary>
        /// Creates a MultiRow that reperesents the multi-row on the account transactions screen.
        /// All this data is hard coded but could be loaded from an XML file chosen by the user.
        /// </summary>
        /// <returns></returns>
        private MultiRow GetMultiRow()
        {
            var multiRow = new MultiRow( "rowData", 14, "[PF8]", "[PF7]", new MoreIndicator( 23, 73, "MORE..." ) );

            multiRow.AddColumn( new ScreenToExcelColumnConversion( "Date", "Date", false, 11, "dd/mm/yyyy", FormatDateText ) );
            multiRow.AddColumn( new ScreenToExcelColumnConversion( "Action", "Action", false, 19 ) );
            multiRow.AddColumn( new ScreenToExcelColumnConversion( "EffDue", "Eff\\Due", false, 11, "dd/mm/yyyy", FormatDateText ) );
            multiRow.AddColumn( new ScreenToExcelColumnConversion( "Amount", "Amount", true, -1, "£#,##0.00" ) );
            multiRow.AddColumn( new ScreenToExcelColumnConversion( "MinDue", "Min Due", true, -1, "£#,##0.00" ) );
            multiRow.AddColumn( new ScreenToExcelColumnConversion( "Description", "Description", false, 15 ) );
            multiRow.AddColumn( new ScreenToExcelColumnConversion( "LOB", "LOB", false ) );

            return multiRow;
        }

        /// <summary>
        /// Formats 6 digit dates from the account transaction screen into something more readable.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string FormatDateText( string text )
        {
            if ( string.IsNullOrWhiteSpace( text ) ||
                 text.Length != 6 ||
                 text == "999999" )
            {
                // Not a date.
                return string.Empty;
            }

            // Get the day, month and year.
            var day = text.Substring( 0, 2 );
            var month = text.Substring( 2, 2 );
            var year = text.Substring( 4, 2 );

            // Add the /s and assume a 2000+ date..
            return day + "/" + month + "/20" + year;
        }

        /// <summary>
        /// Inserts a cell into the given row with the given text.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="includedColumn"></param>
        /// <param name="text"></param>
        private void InsertCellAndText( Worksheet worksheet, uint rowIndex, ScreenToExcelColumnConversion includedColumn, string text )
        {
            // Rows must be added to the SheetData element or they won't appear in the spreadsheet.
            var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();

            if ( sheetData == null )
            {
                throw new OpenXmlPackageException( "No SheetData Element Found." );
            }

            // Look for an existing row with the right index.
            var row = sheetData.Elements<Row>().Where( r => r.RowIndex == rowIndex ).SingleOrDefault();

            if ( row == null )
            {
                // No existing row, so create one and append it to the SheetData.
                row = new Row() { RowIndex = rowIndex };

                sheetData.Append( row );
            }

            // Create an Excel cell ref, A1, B2 etc.
            var cellRef = string.Format( "{0}{1}", includedColumn.ExcelColumnRef, rowIndex );

            var cells = row.Elements<Cell>();

            // Look for an existing cell with the right ref.
            var cell = cells.Where( c => c.CellReference.Value == cellRef ).SingleOrDefault();

            if ( cell == null )
            {
                // No cell, so create one with the right ref and style index.
                cell = new Cell() { CellReference = cellRef, StyleIndex = includedColumn.CellStyleIndex };

                Cell insertBefore = null;

                // Cells need to be in the right order.
                // Look for the 1st cell with a ref greater than that of the cell.
                foreach ( var existingCell in cells )
                {
                    if ( string.Compare( existingCell.CellReference, cellRef, true ) > 0 )
                    {
                        insertBefore = existingCell;

                        break;
                    }
                }

                // Append or Insert the cell.
                if ( insertBefore == null )
                {
                    row.Append( cell );
                }
                else
                {
                    row.InsertBefore( insertBefore, cell );
                }
            }
            else
            {
                // Existing cell, just set its style index.
                cell.StyleIndex = includedColumn.CellStyleIndex;
            }

            if ( includedColumn.IsNumber )
            {
                // Values in this column are numbers, set the value and type.
                cell.CellValue = new CellValue( text );

                cell.DataType = new EnumValue<CellValues>( CellValues.Number );
            }
            else
            {
                // Values in this column are strings.
                // Get the index of the string in the shared strings table.
                // This will add text to the table if it's not already there.
                var index = SharedStrings.GetIndex( text );

                // Set the index and the type on the cell.
                cell.CellValue = new CellValue( index.ToString() );

                cell.DataType = new EnumValue<CellValues>( CellValues.SharedString );
            }

            worksheet.Save();
        }
    }
}
