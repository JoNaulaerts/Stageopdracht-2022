using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.LanguageGeneration;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Builder.Dialogs.Adaptive.Generators;
using Microsoft.Bot.Builder.Dialogs.Adaptive;
using Microsoft.Bot.Builder.Dialogs.Adaptive.Actions;

namespace Microsoft.BotBuilderSamples
{
    public class LogoutDialog : ComponentDialog
    {
        public bool IsLoggedIn = false;
        private static Templates _templates;
        protected string ConnectionName { get; }
        public LanguageGenerator Generator { get; set; }

        public LogoutDialog(string id, string connectionName)
            : base(id)
        {
            string[] paths = { ".", "Dialogs", "LogoutDialog.lg" };
            string fullPath = Path.Combine(paths);
            _templates = Templates.ParseFile(fullPath);

            ConnectionName = connectionName;

            Generator = new TemplateEngineLanguageGenerator(Templates.ParseFile(fullPath));
        }

        protected override async Task<DialogTurnResult> OnBeginDialogAsync(DialogContext innerDc, object options, CancellationToken cancellationToken = default(CancellationToken))
        {
            var result = await InterruptAsync(innerDc, cancellationToken);
            if (result != null)
            {
                return result;
            }

            return await base.OnBeginDialogAsync(innerDc, options, cancellationToken);
        }

        protected override async Task<DialogTurnResult> OnContinueDialogAsync(DialogContext innerDc, CancellationToken cancellationToken = default(CancellationToken))
        {
            var result = await InterruptAsync(innerDc, cancellationToken);
            if (result != null)
            {
                return result;
            }

            return await base.OnContinueDialogAsync(innerDc, cancellationToken);
        }

        private async Task<DialogTurnResult> InterruptAsync(DialogContext innerDc, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (innerDc.Context.Activity.Type == ActivityTypes.Message)
            {
                var text = innerDc.Context.Activity.Text.ToLowerInvariant();

                // Allow logout anywhere in the command
                if (text.IndexOf("logout") >= 0)
                {
                    // The UserTokenClient encapsulates the authentication processes.
                    var userTokenClient = innerDc.Context.TurnState.Get<UserTokenClient>();
                    await userTokenClient.SignOutUserAsync(innerDc.Context.Activity.From.Id, ConnectionName, innerDc.Context.Activity.ChannelId, cancellationToken).ConfigureAwait(false);

                    var replyText = ActivityFactory.FromObject(_templates.Evaluate("LogoutMessage"));
                    //var replyText = MessageFactory.Text("You have been logged out.");

                    await innerDc.Context.SendActivityAsync(replyText, cancellationToken);
                    //new SendActivity.("${LogoutMessage()}");
                    IsLoggedIn = false;
                    return await innerDc.CancelAllDialogsAsync(cancellationToken);
                }
                else if (text.IndexOf("exit") >= 0)
                {
                    // cancel the current dialog and redirect to the help function
                    var replyText = ActivityFactory.FromObject(_templates.Evaluate("ExitMessage"));
                    await innerDc.Context.SendActivityAsync(replyText, cancellationToken);
                    await innerDc.CancelAllDialogsAsync(cancellationToken);
                    return await innerDc.ReplaceDialogAsync("helpWaterfall");                    
                }
            }

            return null;
        }
    }
}