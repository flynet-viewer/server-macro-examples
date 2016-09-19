using System.Collections.Generic;

namespace FVMacros
{
    /// <summary>
    /// Macro that logs into the Insure app and moves from Start to Account Summary.
    /// </summary>
    public sealed class LoginToAccountSummaryMacro : ServerMacro
    {
        public LoginToAccountSummaryMacro() { }

        public LoginToAccountSummaryMacro( MacroHost host ) : base( host ) { }

        public override string name
        {
            get
            {
                return "LoginToAccountSummary";
            }
        }

        public override string description
        {
            get
            {
                return "Login to Account Summary Screen";
            }
        }

        public override string runFromScreenID
        {
            get
            {
                return "Start";
            }
        }

        public override MacroRunResult run()
        {
            var result = this.ClientMessageBox( "Logon?", "Logon", MacroMessageType.info, MacroButtons.yes | MacroButtons.no | MacroButtons.cancel );

            if ( result == MacroDialogRC.yes )
            {
                // Enter keys required to move from Start to AccountSummary.
                this.oScreen.putCommand( "simmy[enter]" );
                this.oScreen.putCommand( "2[enter]" );
                this.oScreen.putCommand( "simmy[tab]host[enter]" );
                this.oScreen.putCommand( "[clear]" );
                this.oScreen.putCommand( "info[enter]" );
                this.oScreen.putCommand( "5012345678[tab][tab][tab][tab][tab][tab][tab]1[enter]" );
            }

            return MacroRunResult.ok;
        }
    }
}
