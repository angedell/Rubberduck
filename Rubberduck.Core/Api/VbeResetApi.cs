using System;

namespace Rubberduck.Api
{
    /// <summary>What a reset did: whether it pressed Run > Reset, and where the editor is now (<see cref="ParserSync.EditorMode"/>).</summary>
    public class VbeResetReport
    {
        /// <summary>True if the editor was stopped or running and Run > Reset was pressed; false if it was in design mode already, so nothing happened.</summary>
        public bool Reset { get; set; }

        public string VbeMode { get; set; }

        public string Note { get; set; }
    }

    /// <summary>
    /// Run > Reset as a call: ends a macro that is stopped or running, as after an error dialog closed with End leaves the editor on the faulty line.
    /// It runs no code, but it throws away the state of the running program (module-level variables, the call stack), so it needs the user's say like a write:
    /// the same approval the edits that would otherwise trigger the editor's own "this action will reset your project" prompt need.
    /// Never claims success it has not seen: the answer is the mode the editor reports afterwards, and anything but design is an error.
    /// </summary>
    public class VbeResetApi
    {
        public const string ResetFailed = "vbe_reset_failed";

        private readonly ApiWriteApproval _approval;
        private readonly Func<string> _editorMode;
        private readonly Func<bool> _resetIfStopped;
        private readonly IDialogWatcher _dialogs;

        /// <param name="editorMode">Where the editor is, as <see cref="ParserSync.EditorMode"/> says.</param>
        /// <param name="resetIfStopped">Presses Run > Reset on the editor's UI thread, as <see cref="ParserSync.ResetIfStopped"/> does; true if the editor is in design mode afterwards.</param>
        public VbeResetApi(ApiWriteApproval approval, Func<string> editorMode, Func<bool> resetIfStopped, IDialogWatcher dialogs = null)
        {
            _approval = approval;
            _editorMode = editorMode;
            _resetIfStopped = resetIfStopped;
            _dialogs = dialogs;
        }

        /// <exception cref="ApiException">The user said no (403), or the editor is not in design mode afterwards (409 <see cref="ResetFailed"/>).</exception>
        public VbeResetReport Reset(string client)
        {
            //Nothing to reset: no question to the user, no command pressed. An editor that does not answer is not in design mode either, so it goes on.
            var before = _editorMode();
            if (before == ParserSync.ModeDesign)
            {
                return new VbeResetReport { Reset = false, VbeMode = before, Note = "The editor was in design mode already; nothing was reset." };
            }

            _approval.Require(client);

            //Runs on the UI thread and waits up to 10 seconds for the editor to take the command. It does not wait for a parse: that follows on its own.
            var inDesign = _resetIfStopped();
            var after = _editorMode();
            if (inDesign && after == ParserSync.ModeDesign)
            {
                return new VbeResetReport { Reset = true, VbeMode = after, Note = "The editor is back in design mode. The parser reads the code again on its own; POST /api/v1/reparse waits for it." };
            }

            throw new ApiException(409, ResetFailed, $"The editor is still not in design mode (vbeMode '{after}'). {Reason(after)}");
        }

        private string Reason(string mode)
        {
            BlockingDialog dialog = null;
            try
            {
                dialog = _dialogs?.Find();
            }
            catch (Exception)
            {
                //A watcher that fails must not hide the real answer.
            }

            if (dialog != null)
            {
                return $"A dialog is open: '{dialog.Title}' - {dialog.Message}. Reset is refused while it is; answer it, then try again. Reset may still be carried out once the dialog is closed.";
            }

            return mode == ParserSync.ModeUnresponsive
                ? "The editor did not answer: it may show a dialog nobody can see, or be stuck in a call. Take a dump (POST /api/v1/dump) if it stays so."
                : "Run > Reset was pressed but the editor still reports a running or stopped macro (code may be re-entering, or the command was not available). Try again, or reset it in the editor.";
        }
    }
}
