using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

using System.IO;
using Microsoft.Bot.Builder.LanguageGeneration;
using Microsoft.Bot.Builder.Dialogs.Adaptive.Generators;
using TeamsAuth.Models;
using Microsoft.Graph;
using Microsoft.Bot.Builder.Dialogs.Choices;
using System.Collections.Generic;
using TeamsAuth;

namespace Microsoft.BotBuilderSamples
{
    public class MainDialog : LogoutDialog
    {
        private static Templates _templates;
        protected readonly ILogger Logger;
        private string helpInputStr = "";
        public HelpClass helpInput;
        private string manageUserInputStr = "";
        public ManageUserClass manageUserClass = new ManageUserClass();
        private string manageGroupInputStr = "";
        public ManageGroupClass manageGroupClass = new ManageGroupClass();
        private string manageGroupMode = "create"; // create, update
        private TokenResponse logintoken;
        private User loggedInUser;
        private string webjobQueueMessage = "";

        public MainDialog(IConfiguration configuration, ILogger<MainDialog> logger)
            : base(nameof(MainDialog), configuration["ConnectionName"])
        {
            // combine path for cross platform support
            string[] paths = { ".", "Dialogs", "MainDialog.lg" };
            string fullPath = Path.Combine(paths);
            _templates = Templates.ParseFile(fullPath);

            Logger = logger;

            //define all dialogs needed in this project. line 44 - 191

            AddDialog(new OAuthPrompt(
                nameof(OAuthPrompt),
                new OAuthPromptSettings
                {
                    ConnectionName = ConnectionName,
                    Text = "You are not signed in.\n\nSign in by tapping the button.",
                    Title = "Sign In",
                    Timeout = 300000, // User has 5 minutes to login (1000 * 60 * 5)
                    EndOnInvalidMessage = true
                }));

            AddDialog(new ConfirmPrompt(nameof(ConfirmPrompt)));

            AddDialog(new TextPrompt(nameof(TextPrompt)));

            AddDialog(new ChoicePrompt(nameof(ChoicePrompt)));

            AddDialog(new WaterfallDialog( "logingWaterfall", new WaterfallStep[]
            {
                PromptStepAsync,
                LoginStepAsync,
            }));

            AddDialog(new WaterfallDialog("helpWaterfall", new WaterfallStep[]
            {
                HelpPromptStepAsync,
                HelpChoisePromptStepAsync,
                HelpRedirectToSubdomainStepAsync,
            }));

            AddDialog(new WaterfallDialog("helpUserWaterfall", new WaterfallStep[]
            {
                HelpUserFirstLayerStepAsync,
                HelpUserSecondLayerStepAsync,
                HelpUserThirdLayerStepAsync,
                HelpFinalRedirectStepAsync,
            }));

            AddDialog(new WaterfallDialog("helpGroupWaterfall", new WaterfallStep[]
            {
                HelpGroupFirstLayerStepAsync,
                HelpFinalRedirectStepAsync,
            }));

            AddDialog(new WaterfallDialog("helpOtherWaterfall", new WaterfallStep[]
            {
                HelpOtherFirstLayerStepAsync,
                HelpOtherSecondLayerStepAsync,
                HelpOtherThirdLayerStepAsync,
                HelpFinalRedirectStepAsync,
            }));

            AddDialog(new WaterfallDialog("ChangePasswordWaterfall", new WaterfallStep[]
            {
                CurrentPasswordStepAsync,
                NewPasswordStepAsync,
                RepeatNewPasswordStepAsync,
                ChangePasswordStepAsync,
            }));

            AddDialog(new WaterfallDialog("ResetPasswordWaterfall", new WaterfallStep[]
            {
                ConfirmResetPasswordStepAsync,
                ResetPasswordStepAsync
            }));

            AddDialog(new WaterfallDialog("CreateNewUserWaterfall", new WaterfallStep[]
            {
                UserDisplayNameStepAsync,
                UserMailNicknameStepAsync,
                UserPrincipalNameStepAsync,
                AccountEnabledStepAsync,
                PostNewUserStepAsync,
            }));

            //AddDialog(new WaterfallDialog("AssignLicenseWaterfall", new WaterfallStep[]
            //{
            //    AssignLicenseStepAsync,

            //}));

            AddDialog(new WaterfallDialog("CreateNewGroupWaterfall", new WaterfallStep[]
            {
                GroupDisplayNameStepAsync,
                GroupDescriptionStepAsync,
                GroupMailNicknameStepAsync,
                AddTeamStepAsync,
                AddOwnerStepAsync,
                PostNewGroupStepAsync,
            }));

            AddDialog(new WaterfallDialog("UpdateGroupWaterfall", new WaterfallStep[]
            {
                GroupIdStepAsync,
                GroupDisplayNameStepAsync,
                GroupDescriptionStepAsync,
                GroupMailNicknameStepAsync,
                AddTeamStepAsync,
                AddOwnerStepAsync,
                PostNewGroupStepAsync,
            }));

            AddDialog(new WaterfallDialog("DeleteGroupWaterfall", new WaterfallStep[]
            {
                GroupIdDeleteStepAsync,
                GroupDeleteConfirmationStepAsync,
                GroupDeleteStepAsync,
            }));

            //AddDialog(new WaterfallDialog("CreatePermissionSharePointWaterfall", new WaterfallStep[]
            //{
            //    SiteNameGetStepAsync,
            //    FollowSiteStepAsync
            //}));

            //AddDialog(new WaterfallDialog("UpdatePermissionSharePointWaterfall", new WaterfallStep[]
            //{
            //    SiteNameGetStepAsync,
            //    UnfollowSiteStepAsync

            //}));

            //AddDialog(new WaterfallDialog("DeletePermissionSharePointWaterfall", new WaterfallStep[]
            //{
            //    SiteIdGetStepAsync,
            //    GetSitePermissionsStepAsync
            //}));

            AddDialog(new WaterfallDialog("ActivateWebjobWaterfall", new WaterfallStep[]
            {
                GetMessageForQueueStepAsync,
            }));

            // The initial child Dialog to run.
            InitialDialogId = "logingWaterfall";

            Generator = new TemplateEngineLanguageGenerator(Templates.ParseFile(fullPath));
        }

        private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }
                   
        private async Task<DialogTurnResult> LoginStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Get the token from the previous step. Note that we could also have gotten the
            // token directly from the prompt itself. There is an example of this in the next method.
            var tokenResponse = (TokenResponse)stepContext.Result;
            logintoken = (TokenResponse)stepContext.Result;
            if (tokenResponse?.Token != null)
            {
                if (IsLoggedIn == false) {

                // Pull in the data from the Microsoft Graph API.
                var client = new SimpleGraphClient(tokenResponse.Token);
                var me = await client.GetMeAsync();
                var title = !string.IsNullOrEmpty(me.JobTitle) ?
                            me.JobTitle : "Unknown";

                // Show the data from the Microsoft Graph API.
                await stepContext.Context.SendActivityAsync(
                    $"You're logged in as {me.DisplayName} ({me.UserPrincipalName})."+
                    $"\n\nYour job title is: {title}");

                    loggedInUser = me;

                    IsLoggedIn = true;
                }

                // Redirect to the helpdialog.
                return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }

            await stepContext.Context.SendActivityAsync(MessageFactory.Text("Login was not successful please try again."), cancellationToken);
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }

        private async Task<DialogTurnResult> HelpPromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Reset values in case of the "exit interruption"
            helpInputStr = "";
            manageUserInputStr = "";
            manageUserClass.resetValues();
            manageGroupInputStr = "";
            manageGroupClass.resetValues();
            manageGroupMode = "create"; // create, update
            webjobQueueMessage = "";

        // Define the messageprompt for this WaterfallStep.
        var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("HelpMessage")),
            };
            
            // Define new object of HelpClass to retrieve userinput in later stages of this WaterfallDialog.
            stepContext.Values[helpInputStr] = new HelpClass();
            //await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("HelpMessage")));
            return await stepContext.NextAsync();
        }

        private async Task<DialogTurnResult> HelpChoisePromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define new object of HelpClass to retrieve userinput in later stages of this WaterfallDialog.
            stepContext.Values[helpInputStr] = new HelpClass();
            return await stepContext.PromptAsync(nameof(ChoicePrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseBroadTopicMessage")),
                        Choices = ChoiceFactory.ToChoices(new List<string> { "User", "Group", "Other" }),
                    }, cancellationToken);
        }

        private async Task<DialogTurnResult> HelpRedirectToSubdomainStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Retrieve userinput from the previous step in this WaterfallDialog.
            var helpInput = (HelpClass)stepContext.Values[helpInputStr];
            helpInput.helpInput = ((FoundChoice)stepContext.Result).Value.ToString();

            // Redirect to a new dialog based on the input of the previous step.
            switch (helpInput.helpInput.ToString())
            {
                case "User":
                    return await stepContext.ReplaceDialogAsync("helpUserWaterfall");
                case "Group":
                    return await stepContext.ReplaceDialogAsync("helpGroupWaterfall");
                case "Other":
                    return await stepContext.ReplaceDialogAsync("helpOtherWaterfall");
                default:
                    await stepContext.PromptAsync(nameof(TextPrompt),
                                       new PromptOptions
                                       {
                                           Prompt = ActivityFactory.FromObject(_templates.Evaluate("TryAgainMessage")),
                                       }, cancellationToken);
                    return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }
        }

        private async Task<DialogTurnResult> HelpUserFirstLayerStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define new object of HelpClass to retrieve userinput in later stages of this WaterfallDialog.
            stepContext.Values[helpInputStr] = new HelpClass();
            return await stepContext.PromptAsync(nameof(ChoicePrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseSpecificTopicMessage")),
                        Choices = ChoiceFactory.ToChoices(new List<string> { "Password", "User changes", "logout" }),
                    }, cancellationToken);
        }
        private async Task<DialogTurnResult> HelpUserSecondLayerStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Retrieve userinput from the previous step in this WaterfallDialog.
            var helpInput = (HelpClass)stepContext.Values[helpInputStr];
            helpInput.helpInput = ((FoundChoice)stepContext.Result).Value.ToString();

            // Redirect to The following dialog step based on the input of the previous step.
            switch (helpInput.helpInput.ToString())
            {
                case "Password":
                    helpInput.helpTopic = "password";
                    return await stepContext.NextAsync("password");
                case "User changes":
                    helpInput.helpTopic = "userCRUD";
                    return await stepContext.NextAsync("userCRUD");
                // logout is an interruptfunction so it is not used in this particular case
                default:
                    await stepContext.PromptAsync(nameof(TextPrompt),
                                       new PromptOptions
                                       {
                                           Prompt = ActivityFactory.FromObject(_templates.Evaluate("TryAgainMessage")),
                                       }, cancellationToken);
                    return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }
        }

        private async Task<DialogTurnResult> HelpUserThirdLayerStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (stepContext.Result == "password")
            {
                return await stepContext.PromptAsync(nameof(ChoicePrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseCommandMessage")),
                        Choices = ChoiceFactory.ToChoices(new List<string> { "change password", "reset password" }),
                    }, cancellationToken);
            }
            else if (stepContext.Result.ToString() == "userCRUD")
            {
                return await stepContext.PromptAsync(nameof(ChoicePrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseCommandMessage")),
                        Choices = ChoiceFactory.ToChoices(new List<string> { "create user"}),
                    }, cancellationToken);
            }
            else
            {
                await stepContext.PromptAsync(nameof(TextPrompt),
                                   new PromptOptions
                                   {
                                       Prompt = ActivityFactory.FromObject(_templates.Evaluate("TryAgainMessage")),
                                   }, cancellationToken);
                return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }
        }
        private async Task<DialogTurnResult> HelpGroupFirstLayerStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define new object of HelpClass to retrieve userinput in later stages of this WaterfallDialog.
            stepContext.Values[helpInputStr] = new HelpClass();
            return await stepContext.PromptAsync(nameof(ChoicePrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseCommandMessage")),
                        Choices = ChoiceFactory.ToChoices(new List<string> { "create group", "update group", "delete group" }),
                    }, cancellationToken);
        }
        private async Task<DialogTurnResult> HelpOtherFirstLayerStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define new object of HelpClass to retrieve userinput in later stages of this WaterfallDialog.
            stepContext.Values[helpInputStr] = new HelpClass();
            return await stepContext.PromptAsync(nameof(ChoicePrompt),
                new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseSpecificTopicMessage")),
                    Choices = ChoiceFactory.ToChoices(new List<string> { "SharePoint", "Webjobs", "help" }),
                }, cancellationToken);
        }

        private async Task<DialogTurnResult> HelpOtherSecondLayerStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Retrieve userinput from the previous step in this WaterfallDialog.
            var helpInput = (HelpClass)stepContext.Values[helpInputStr];
            helpInput.helpInput = ((FoundChoice)stepContext.Result).Value.ToString();

            // Redirect to The following dialog step based on the input of the previous step.
            switch (helpInput.helpInput.ToString())
            {
                case "SharePoint":
                    helpInput.helpTopic = "SharePoint";
                    return await stepContext.NextAsync("SharePoint");
                case "Webjobs":
                    helpInput.helpTopic = "Webjobs";
                    return await stepContext.NextAsync("webjob");
                case "help":
                    return await stepContext.ReplaceDialogAsync("helpWaterfall");
                default:
                    await stepContext.PromptAsync(nameof(TextPrompt),
                                       new PromptOptions
                                       {
                                           Prompt = ActivityFactory.FromObject(_templates.Evaluate("TryAgainMessage")),
                                       }, cancellationToken);
                    return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }
        }

        private async Task<DialogTurnResult> HelpOtherThirdLayerStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Show a choiceprompt based on the input of the input of the previous step.
            if (stepContext.Result == "SharePoint")
            {
                return await stepContext.PromptAsync(nameof(ChoicePrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseCommandMessage")),
                        Choices = ChoiceFactory.ToChoices(new List<string> { "create permission SP", "update permission SP", "delete permission SP"}),
                    }, cancellationToken);
            }
            else if (stepContext.Result.ToString() == "webjob")
            {
                return await stepContext.PromptAsync(nameof(ChoicePrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseCommandMessage")),
                        Choices = ChoiceFactory.ToChoices(new List<string> { "Webjob: Azure report" }),
                    }, cancellationToken);
            }
            else
            {
                await stepContext.PromptAsync(nameof(TextPrompt),
                                   new PromptOptions
                                   {
                                       Prompt = ActivityFactory.FromObject(_templates.Evaluate("TryAgainMessage")),
                                   }, cancellationToken);
                return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }
        }

        private async Task<DialogTurnResult> HelpFinalRedirectStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Retrieve userinput from the previous step in this WaterfallDialog.
            var helpInput = (HelpClass)stepContext.Values[helpInputStr];
            helpInput.helpInput = ((FoundChoice)stepContext.Result).Value.ToString();

            // Redirect to a fitting WaterfallDialog based on the userinput.
            switch (helpInput.helpInput.ToString())
            {
                case "change password":
                    return await stepContext.ReplaceDialogAsync("ChangePasswordWaterfall");
                case "reset password":
                    return await stepContext.ReplaceDialogAsync("ResetPasswordWaterfall");
                //case "assign license":
                //    return await stepContext.ReplaceDialogAsync("AssignLicenseWaterfall");
                case "create user":
                    return await stepContext.ReplaceDialogAsync("CreateNewUserWaterfall");
                case "create group":
                    return await stepContext.ReplaceDialogAsync("CreateNewGroupWaterfall");
                case "update group":
                    return await stepContext.ReplaceDialogAsync("UpdateGroupWaterfall");
                case "delete group":
                    return await stepContext.ReplaceDialogAsync("DeleteGroupWaterfall");
                case "help":
                case "?":
                    await stepContext.PromptAsync(nameof(ChoicePrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("ChooseBroadTopicMessage")),
                        Choices = ChoiceFactory.ToChoices(new List<string> { "User", "Group", "Other"}),
                    }, cancellationToken);
                    return await stepContext.ReplaceDialogAsync("helpWaterfall");
                //case "create permission SP":
                //    return await stepContext.ReplaceDialogAsync("CreatePermissionSharePointWaterfall");
                //case "update permission SP":
                //    return await stepContext.ReplaceDialogAsync("UpdatePermissionSharePointWaterfall");
                //case "delete permission SP":
                //    return await stepContext.ReplaceDialogAsync("DeletePermissionSharePointWaterfall");
                case "Webjob: Azure report":
                    webjobQueueMessage = "webjob " + logintoken.Token;
                    return await stepContext.ReplaceDialogAsync("ActivateWebjobWaterfall");

                default:
                    await stepContext.PromptAsync(nameof(TextPrompt),
                    new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("TryAgainMessage")),
                    }, cancellationToken);
                    return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }
        }
        
        private async Task<DialogTurnResult> CurrentPasswordStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("AskForCurrentPassword")),
            };

            // Define object of ManageUserClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageUserInputStr] = new ManageUserClass();
            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }

        private async Task<DialogTurnResult> NewPasswordStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("AskForNewPassword")),
            };

            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var currentPasswordInput = (ManageUserClass)stepContext.Values[manageUserInputStr];
            currentPasswordInput.currentPassword = (string)stepContext.Result;

            manageUserClass.currentPassword = currentPasswordInput.currentPassword;

            // Redefine object of ManageUserClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageUserInputStr] = new ManageUserClass();
            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }

        private async Task<DialogTurnResult> RepeatNewPasswordStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("AskForNewPasswordAgain")),
            };

            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var newPasswordInput = (ManageUserClass)stepContext.Values[manageUserInputStr];
            newPasswordInput.newPassword = (string)stepContext.Result;

            manageUserClass.newPassword = newPasswordInput.newPassword;

            // Redefine object of ManageUserClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageUserInputStr] = new ManageUserClass();
            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }
        private async Task<DialogTurnResult> ChangePasswordStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var controlNewPasswordInput = (ManageUserClass)stepContext.Values[manageUserInputStr];
            controlNewPasswordInput.contolNewPassword = (string)stepContext.Result;

            manageUserClass.contolNewPassword = controlNewPasswordInput.contolNewPassword;

            if (manageUserClass.newPassword == manageUserClass.contolNewPassword)
            {
                // Use your authenticationtoken and the fully filled global object to change your password 
                // with the SimpleGraphClient
                var tokenResponse = logintoken;
                var client = new SimpleGraphClient(tokenResponse.Token);
                await client.PostNewPasswordAsync(manageUserClass);

                // Reset the global object for later use and redirect to the helpdialog.
                manageUserClass.resetValues();

                await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("PasswordResetSuccessMessage")));
                return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }
            else
            {
                // Reset the global object for later use and redirect to the helpdialog.
                manageUserClass.resetValues();

                await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("PasswordResetFailMessage")));
                return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }            
        }

        private async Task<DialogTurnResult> ConfirmResetPasswordStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var promptOptions = new PromptOptions
            {
                Prompt = MessageFactory.Text("Are you sure you want to reset your password?"),
            };

            // Redefine object of ManageUserClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

            return await stepContext.PromptAsync(nameof(ConfirmPrompt), promptOptions, cancellationToken);
        }

        private async Task<DialogTurnResult> ResetPasswordStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var confirmDeleteInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
            confirmDeleteInput.confirmDeleteGroup = (bool)stepContext.Result;

            manageGroupClass.confirmDeleteGroup = confirmDeleteInput.confirmDeleteGroup;

            if (confirmDeleteInput.confirmDeleteGroup) {
                // Use your authenticationtoken and the fully filled global object to reset your password 
                // with the SimpleGraphClient
                var tokenResponse = logintoken;
                var client = new SimpleGraphClient(tokenResponse.Token);
                var newPassword = await client.ResetPasswordAsync(loggedInUser);

                var promptOptions = new PromptOptions
                {
                    Prompt = MessageFactory.Text("Your password is sucessfully reset \n\nYour temporary password is: " + newPassword),
                };
                await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            } else
            {
                var promptOptions = new PromptOptions
                {
                    Prompt = MessageFactory.Text("Canceling password reset."),
                };
                await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            }


            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }

        private async Task<DialogTurnResult> UserDisplayNameStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("AskForDisplayName")),
            };

            // Define object of ManageUserClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageUserInputStr] = new ManageUserClass();

            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }

        private async Task<DialogTurnResult> UserMailNicknameStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("AskForMailNickname")),
            };

            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var displayNameInput = (ManageUserClass)stepContext.Values[manageUserInputStr];
            displayNameInput.displayName = (string)stepContext.Result;

            manageUserClass.displayName = displayNameInput.displayName;

            // Redefine object of ManageUserClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageUserInputStr] = new ManageUserClass();

            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }

        private async Task<DialogTurnResult> UserPrincipalNameStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("AskForUserPrincipalName")),
            };

            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var mailNicknameInput = (ManageUserClass)stepContext.Values[manageUserInputStr];
            mailNicknameInput.mailNickname = (string)stepContext.Result;

            manageUserClass.mailNickname = mailNicknameInput.mailNickname;

            // Redefine object of ManageUserClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageUserInputStr] = new ManageUserClass();

            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }
        
        private async Task<DialogTurnResult> AccountEnabledStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("AccountEnabledPrompt")),
            };

            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var userPrincipalNameInput = (ManageUserClass)stepContext.Values[manageUserInputStr];
            userPrincipalNameInput.userPrincipalName = (string)stepContext.Result;

            manageUserClass.userPrincipalName = userPrincipalNameInput.userPrincipalName;

            // Redefine object of ManageUserClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageUserInputStr] = new ManageUserClass();

            return await stepContext.PromptAsync(nameof(ConfirmPrompt), promptOptions, cancellationToken);
        }

        private async Task<DialogTurnResult> PostNewUserStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("UsercreatedSuccessMessage")),
            };

            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var accountEnabledInput = (ManageUserClass)stepContext.Values[manageUserInputStr];
            accountEnabledInput.accountEnabled = (bool)stepContext.Result;

            manageUserClass.accountEnabled = accountEnabledInput.accountEnabled;

            // Use your authenticationtoken and the fully filled global object to create a new user 
            // with the SimpleGraphClient
            var tokenResponse = logintoken;
            var client = new SimpleGraphClient(tokenResponse.Token);
            var password = await client.PostCreateUserAsync(manageUserClass);

            // Reset the global object for later use and redirect to the helpdialog.
            manageUserClass.resetValues();

            await stepContext.Context.SendActivityAsync(MessageFactory.Text("The temporary password of this user is: " + password), cancellationToken);

            await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }

        //private async Task<DialogTurnResult> AssignLicenseStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        //{
        //    var promptOptions = new PromptOptions
        //    {
        //        Prompt = ActivityFactory.FromObject(_templates.Evaluate("LicenseAssignedSuccessMessage")),
        //        //Prompt = MessageFactory.Text("License assigned."),
        //    };

        //    await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        //    return await stepContext.ReplaceDialogAsync("helpWaterfall");
        //}

        private async Task<DialogTurnResult> GroupDisplayNameStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (manageGroupMode == "create")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("GroupDisplaynamePrompt")),
                };

                // Define object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            }
            else if (manageGroupMode == "update")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("UpdateGroupDisplaynamePrompt")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
                var groupDisplayNameInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                groupDisplayNameInput.groupDisplayName = (string)stepContext.Result;

                //// Use your authenticationtoken and displayName to retrieve the data from the existing group
                //// with the SimpleGraphClient
                var tokenResponse = logintoken;
                var client = new SimpleGraphClient(tokenResponse.Token);
                Group currentGroup = await client.GetGroupByDisplayNameAsync(groupDisplayNameInput.groupDisplayName);

                // Fill global object with current data of the existing group
                manageGroupClass.groupId = currentGroup.Id;
                manageGroupClass.groupDisplayName = currentGroup.DisplayName;
                manageGroupClass.groupDescription = currentGroup.Description;
                manageGroupClass.groupMailNickname = currentGroup.MailNickname;
                var testvarTeam = currentGroup.Team;

                // Define object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            }
            await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("SomethingWentWrongMessage")));
            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }

        private async Task<DialogTurnResult> GroupDescriptionStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (manageGroupMode == "create")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("GroupDescriptionPrompt")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
                var groupDisplayNameInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                groupDisplayNameInput.groupDisplayName = (string)stepContext.Result;

                manageGroupClass.groupDisplayName = groupDisplayNameInput.groupDisplayName;

                // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            }
            else if (manageGroupMode == "update")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("UpdateGroupDescriptionPrompt")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog.
                var groupDisplayNameInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                groupDisplayNameInput.groupDisplayName = (string)stepContext.Result;

                if (groupDisplayNameInput.groupDisplayName == "/")
                {
                    // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                    stepContext.Values[manageGroupInputStr] = new ManageGroupClass();
                    return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
                }
                else
                {
                    manageGroupClass.groupDisplayName = groupDisplayNameInput.groupDisplayName;

                    // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                    stepContext.Values[manageGroupInputStr] = new ManageGroupClass();
                    return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
                }                
            }
            await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("SomethingWentWrongMessage")));
            return await stepContext.ReplaceDialogAsync("helpWaterfall");

        }

        private async Task<DialogTurnResult> GroupMailNicknameStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (manageGroupMode == "create")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("GroupMailNicknamePrompt")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
                var groupDescriptionInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                groupDescriptionInput.groupDescription = (string)stepContext.Result;

                manageGroupClass.groupDescription = groupDescriptionInput.groupDescription;

                // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            }
            else if (manageGroupMode == "update")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("UpdateGroupMailNicknamePrompt")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog.
                var groupDescriptionInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                groupDescriptionInput.groupDescription = (string)stepContext.Result;

                // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                if (groupDescriptionInput.groupDescription == "/")
                {
                    // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                    stepContext.Values[manageGroupInputStr] = new ManageGroupClass();
                    return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
                }
                else
                {
                    manageGroupClass.groupDescription = groupDescriptionInput.groupDescription;

                    // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                    stepContext.Values[manageGroupInputStr] = new ManageGroupClass();
                    return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
                }
            }
            await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("SomethingWentWrongMessage")));
            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }

        private async Task<DialogTurnResult> AddTeamStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (manageGroupMode == "create")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("AddTeamMessage")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
                var groupMailNicknameInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                groupMailNicknameInput.groupMailNickname = (string)stepContext.Result;

                manageGroupClass.groupMailNickname = groupMailNicknameInput.groupMailNickname;

                // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                return await stepContext.PromptAsync(nameof(ConfirmPrompt), promptOptions, cancellationToken);
            }
            else if (manageGroupMode == "update")
            {
                // Use your authenticationtoken and the filled global object to verify if a group already has a team or not
                // with the SimpleGraphClient
                var tokenResponse = logintoken;
                var client = new SimpleGraphClient(tokenResponse.Token);
                var teamExists = await client.GetTeamVerifyAsync(manageGroupClass.groupId);

                if (teamExists)
                {
                    var promptOptions = new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("TeamAlreadyExistsMessage")),
                    };

                    manageGroupClass.teamExists = teamExists;
                    manageGroupClass.addTeam = false;

                    await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("TeamAlreadyExistsMessage")));
                    return await stepContext.NextAsync();
                }
                else
                {
                    // Define the messageprompt for this WaterfallStep.
                    var promptOptions = new PromptOptions
                    {
                        Prompt = ActivityFactory.FromObject(_templates.Evaluate("AddTeamMessage")),
                    };

                    // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
                    var groupMailNicknameInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                    groupMailNicknameInput.groupMailNickname = (string)stepContext.Result;

                    manageGroupClass.groupMailNickname = groupMailNicknameInput.groupMailNickname;

                    // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                    stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                    manageGroupClass.teamExists = teamExists;
                    manageGroupClass.addTeam = true;
                    return await stepContext.PromptAsync(nameof(ConfirmPrompt), promptOptions, cancellationToken);
                }
            }

            await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("SomethingWentWrongMessage")));
            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }

        private async Task<DialogTurnResult> AddOwnerStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (manageGroupMode == "create")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("AddOwnerMessage")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
                var addTeamInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                addTeamInput.addTeam = (bool)stepContext.Result;

                manageGroupClass.addTeam = addTeamInput.addTeam;

                // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            }
            else if (manageGroupMode == "update")
            {
                if (manageGroupClass.teamExists == false)
                {
                    // Retrieve userinput from the previous step in this WaterfallDialog.
                    var addTeamInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                    addTeamInput.addTeam = (bool)stepContext.Result;

                    manageGroupClass.addTeam = addTeamInput.addTeam;
                }
                else
                {
                    var addTeamInput = new ManageGroupClass();
                    addTeamInput.addTeam = false;

                    manageGroupClass.addTeam = addTeamInput.addTeam;
                }

                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("AddOwnerMessage")),
                };

                // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
                stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

                return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            }

            await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("SomethingWentWrongMessage")));
            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }

        private async Task<DialogTurnResult> PostNewGroupStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (manageGroupMode == "create") {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("GroupCreatedMessage")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
                var addownerInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                addownerInput.ownerPrincipalName = (string)stepContext.Result;

                manageGroupClass.ownerPrincipalName = addownerInput.ownerPrincipalName;

                // Use your authenticationtoken and the fully filled global object to create a new group
                // with the SimpleGraphClient
                var tokenResponse = logintoken;
                var client = new SimpleGraphClient(tokenResponse.Token);
                var returnGroup = await client.PostUpdateCreateGroupAsync(manageGroupClass, manageGroupMode);

                manageGroupClass.groupId = returnGroup.Id;
                if (manageGroupClass.ownerPrincipalName != "/")
                {
                    await client.AddGroupOwnerAsync(manageGroupClass.groupId, manageGroupClass.ownerPrincipalName);
                }

                // Reset the global object for later use and redirect to the helpdialog.
                manageGroupClass.resetValues();

                await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
                return await stepContext.ReplaceDialogAsync("helpWaterfall");
            }
            else if (manageGroupMode == "update")
            {
                // Define the messageprompt for this WaterfallStep.
                var promptOptions = new PromptOptions
                {
                    Prompt = ActivityFactory.FromObject(_templates.Evaluate("GroupUpdatedMessage")),
                };

                // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
                var addownerInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
                addownerInput.ownerPrincipalName = (string)stepContext.Result;

                manageGroupClass.ownerPrincipalName = addownerInput.ownerPrincipalName;

                if (manageGroupClass.addTeam == false || manageGroupClass.addTeam == null )
                {
                    // Use your authenticationtoken and the fully filled global object to update an existing group
                    // with the SimpleGraphClient
                    var tokenResponse = logintoken;
                    var client = new SimpleGraphClient(tokenResponse.Token);
                    await client.PostUpdateCreateGroupAsync(manageGroupClass, manageGroupMode);

                    if (manageGroupClass.ownerPrincipalName != "/")
                    {
                        await client.AddGroupOwnerAsync(manageGroupClass.groupId, manageGroupClass.ownerPrincipalName);
                    }

                    // Reset the global object and manageGroupMode for later use and redirect to the helpdialog.
                    manageGroupClass.resetValues();
                    manageGroupMode = "create";

                    await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
                    return await stepContext.ReplaceDialogAsync("helpWaterfall");
                }
                else
                {
                    // Use your authenticationtoken and the fully filled global object to update an existing group
                    // with the SimpleGraphClient
                    var tokenResponse = logintoken;
                    var client = new SimpleGraphClient(tokenResponse.Token);
                    await client.PostUpdateCreateGroupAsync(manageGroupClass, manageGroupMode);

                    if (manageGroupClass.ownerId != "/")
                    {
                        await client.AddGroupOwnerAsync(manageGroupClass.groupId, manageGroupClass.ownerId);
                    }

                    // Reset the global object and manageGroupMode for later use and redirect to the helpdialog.
                    manageGroupClass.resetValues();
                    manageGroupMode = "create";

                    await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
                    return await stepContext.ReplaceDialogAsync("helpWaterfall");
                }
            }
            await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("SomethingWentWrongMessage")));
            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }

        private async Task<DialogTurnResult> GroupIdStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("GroupDisplayNamePrompt")),
            };

            manageGroupMode = "update";

            // Define object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }

        private async Task<DialogTurnResult> GroupIdDeleteStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("DeleteGroupDisplayNamePrompt")),
            };

            // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }

        private async Task<DialogTurnResult> GroupDeleteConfirmationStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var groupDisplayNameInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
            groupDisplayNameInput.groupDisplayName = (string)stepContext.Result;

            //// Use your authenticationtoken and displayName to retrieve the data from the existing group
            //// with the SimpleGraphClient
            var tokenResponse = logintoken;
            var client = new SimpleGraphClient(tokenResponse.Token);
            Group currentGroup = await client.GetGroupByDisplayNameAsync(groupDisplayNameInput.groupDisplayName);

            // Fill global object with current data of the existing group
            manageGroupClass.groupId = currentGroup.Id;
            manageGroupClass.groupDisplayName = currentGroup.DisplayName;
            manageGroupClass.groupDescription = currentGroup.Description;
            manageGroupClass.groupMailNickname = currentGroup.MailNickname;

            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                //Prompt = ActivityFactory.FromObject(_templates.Evaluate("GroupConfirmDeletePrompt", "groepsnaam komt hier" + "?" )),
                Prompt = MessageFactory.Text("Are you sure you want to delete " + manageGroupClass.groupDisplayName + "?"),
            };

            // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageGroupInputStr] = new ManageGroupClass();
            return await stepContext.PromptAsync(nameof(ConfirmPrompt), promptOptions, cancellationToken);
        }
        private async Task<DialogTurnResult> GroupDeleteStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
            var confirmDeleteInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
            confirmDeleteInput.confirmDeleteGroup = (bool)stepContext.Result;

            manageGroupClass.confirmDeleteGroup = confirmDeleteInput.confirmDeleteGroup;

            if (confirmDeleteInput.confirmDeleteGroup)
            {
                // Use your authenticationtoken and the groupId to delete an existing group
                // with the SimpleGraphClient
                var tokenResponse = logintoken;
                var client = new SimpleGraphClient(tokenResponse.Token);
                await client.DeleteGroupAsync(manageGroupClass.groupId);

                // Reset the global object for later use and redirect to the helpdialog.
                manageGroupClass.resetValues();

                await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("GroupDeleteSuccessMessage")));
            } else
            {
                await stepContext.Context.SendActivityAsync(ActivityFactory.FromObject(_templates.Evaluate("GroupCancelDeleteMessage")));
            }

            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }        

        private async Task<DialogTurnResult> SiteNameGetStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("GroupDisplayNamePrompt")),
            };

            manageGroupMode = "update";

            // Define object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

            return await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
        }

        //private async Task<DialogTurnResult> FollowSiteStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        //{
        //    // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
        //    var groupNameInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
        //    groupNameInput.groupId = (string)stepContext.Result;

        //    manageGroupClass.groupId = groupNameInput.groupId;

        //    // Use your authenticationtoken and groupId to retrieve the data from the existing group
        //    // with the SimpleGraphClient
        //    var tokenResponse = logintoken;
        //    var client = new SimpleGraphClient(tokenResponse.Token);
        //    var me = await client.GetMeAsync();
        //    var currentPermissions = await client.SharePointFollowSiteAsync(groupNameInput.groupId, me);

        //    // Define the messageprompt for this WaterfallStep.
        //    var promptOptions = new PromptOptions
        //    {
        //        Prompt = MessageFactory.Text("You followed:" + groupNameInput.groupId),
        //    };

        //    // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
        //    stepContext.Values[manageGroupInputStr] = new ManageGroupClass();
        //    await stepContext.PromptAsync(nameof(ConfirmPrompt), promptOptions, cancellationToken);
        //    return await stepContext.ReplaceDialogAsync("helpWaterfall");
        //}

        //private async Task<DialogTurnResult> UnfollowSiteStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        //{
        //    // Retrieve userinput from the previous step in this WaterfallDialog and put the value in a global object.
        //    var groupNameInput = (ManageGroupClass)stepContext.Values[manageGroupInputStr];
        //    groupNameInput.groupId = (string)stepContext.Result;

        //    manageGroupClass.groupId = groupNameInput.groupId;

        //    // Use your authenticationtoken and groupId to retrieve the data from the existing group
        //    // with the SimpleGraphClient
        //    var tokenResponse = logintoken;
        //    var client = new SimpleGraphClient(tokenResponse.Token);
        //    var me = await client.GetMeAsync();
        //    var currentPermissions = await client.SharePointUnfollowSiteAsync(groupNameInput.groupId, me);

        //    // Define the messageprompt for this WaterfallStep.
        //    var promptOptions = new PromptOptions
        //    {
        //        Prompt = MessageFactory.Text("You unfollowed:" + groupNameInput.groupId),
        //    };

        //    // Redefine object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
        //    stepContext.Values[manageGroupInputStr] = new ManageGroupClass();
        //    await stepContext.PromptAsync(nameof(ConfirmPrompt), promptOptions, cancellationToken);
        //    return await stepContext.ReplaceDialogAsync("helpWaterfall");
        //}

        private async Task<DialogTurnResult> GetMessageForQueueStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Define the messageprompt for this WaterfallStep.
            var promptOptions = new PromptOptions
            {
                Prompt = ActivityFactory.FromObject(_templates.Evaluate("WebjobStartedMessage")),
            };

            var client = new SimpleStorageAccountClient();
            client.InsertMessage("queue", webjobQueueMessage);

            // Define object of ManageGroupClass to retrieve userinput in the next WaterfallStep.
            stepContext.Values[manageGroupInputStr] = new ManageGroupClass();

            await stepContext.PromptAsync(nameof(TextPrompt), promptOptions, cancellationToken);
            return await stepContext.ReplaceDialogAsync("helpWaterfall");
        }        
    }
}
