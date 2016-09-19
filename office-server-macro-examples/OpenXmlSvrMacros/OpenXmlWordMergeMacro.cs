using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FVMacros
{
    /// <summary>
    /// Macro that merges data from the screen into a Word template.
    /// </summary>
    public sealed class OpenXmlWordMergeMacro : ServerMacro
    {
        public OpenXmlWordMergeMacro() { }

        public OpenXmlWordMergeMacro( MacroHost host ) : base( host ) { }

        public override string name
        {
            get
            {
                return "OpenXmlWordMergeMacro";
            }
        }

        public override string description
        {
            get
            {
                return "open Xml Word Merge Macro";
            }
        }

        public override string runFromScreenID
        {
            get
            {
                return "AcctSummary";
            }
        }

        public override MacroRunResult run()
        {
            try
            {
                // Get the .docx from the user.
                var result = GetClientFile( "Template.docx", "Please Upload a Document Template", "Template Upload", false );

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

                // TODO: check that it's a .docx.

                // New filename to avoid overwriting the template.
                var pathAndNewFilename = Path.Combine( Path.GetDirectoryName( xferReceiveFile ), string.Format( "{0}_Merged{1}", Path.GetFileNameWithoutExtension( xferReceiveName ), Path.GetExtension( xferReceiveName ) ) );

                var mergedDoc = new FileInfo( pathAndNewFilename );

                if ( File.Exists( mergedDoc.FullName ) )
                {
                    File.Delete( mergedDoc.FullName );
                }

                File.Copy( xferReceiveFile, mergedDoc.FullName );

                using ( var doc = WordprocessingDocument.Open( mergedDoc.FullName, true ) )
                {
                    var fieldCodes = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().ToList();

                    foreach ( var fieldCode in fieldCodes )
                    {
                        ReplaceMergeField( fieldCode );
                    }

                    var simpleFields = doc.MainDocumentPart.RootElement.Descendants<SimpleField>().ToList();

                    foreach ( var simpleField in simpleFields )
                    {
                        ReplaceSimpleField( simpleField );
                    }
                }

                // Let the user save the modified file.
                var saveResult = PutClientFile( pathAndNewFilename, "Save Merged File", "Template Merged" );

                if ( saveResult == MacroDialogRC.ok )
                {
                    // Tidy up.
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

        private void ReplaceMergeField( FieldCode fieldCode )
        {
            //Console.WriteLine( fieldCode.Text );

            if ( !fieldCode.Text.Contains( "MERGEFIELD" ) )
            {
                return;
            }

            var parentRun = (Run) fieldCode.Parent;

            var beginRun = parentRun.PreviousSibling<Run>();
            var separateRun = parentRun.NextSibling<Run>();
            var textRun = separateRun.NextSibling<Run>();
            var endRun = textRun.NextSibling<Run>();

            var text = textRun.GetFirstChild<Text>();

            text.Text = GetTextForField( text );

            parentRun.Remove();
            beginRun.Remove();
            separateRun.Remove();
            endRun.Remove();
        }

        private void ReplaceSimpleField( SimpleField simpleField )
        {
            var parent = simpleField.Parent;

            var run = simpleField.GetFirstChild<Run>();

            var text = run.GetFirstChild<Text>();

            //Console.WriteLine( text.Text );

            text.Text = GetTextForField( text );

            simpleField.RemoveChild<Run>( run );

            parent.ReplaceChild<SimpleField>( run, simpleField );
        }

        private string GetTextForField( Text text )
        {
            var fieldName = text.Text.Trim();

            if ( string.IsNullOrWhiteSpace( fieldName ) && fieldName.Length < 3 )
            {
                return string.Empty;
            }

            var screenFieldName = text.Text.Substring( 1, fieldName.Length - 2 );

            var screenFieldText = string.Empty;

            if ( screenFieldName == "AccountNumber" )
            {
                // Account number is a read/write field and not mapped.
                // Get its value by position.
                screenFieldText = oScreen.getField( 1, 11 ).value;
            }
            else
            {
                // Get the text from the screen for a filed with the right name.
                screenFieldText = oScreen.mappedGet( screenFieldName );
            }

            if ( string.IsNullOrWhiteSpace( screenFieldText ) )
            {
                screenFieldText = string.Empty;
            }

            return screenFieldText;
        }
    }
}
