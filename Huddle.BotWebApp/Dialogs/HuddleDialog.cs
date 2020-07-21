/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System.Threading;
using System.Threading.Tasks;
using Huddle.BotWebApp.Bots;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;

namespace Huddle.BotWebApp.Dialogs
{

    public class HuddleDialog : ComponentDialog
    {
        private const string HelpMsgText = "Show help here";
        private const string CancelMsgText = "Cancelling...";

        protected readonly string ConnectionName;
        protected readonly UserState UserState;
        protected readonly IStatePropertyAccessor<UserProfile> UserProfileAccessor;

        public HuddleDialog(string id, IConfiguration configuration, UserState userState)
            : base(id)
        {
            ConnectionName = configuration["ConnectionName"];
            UserState = userState;
            UserProfileAccessor = userState.CreateProperty<UserProfile>("UserProfile");

            var oauthPromptSettings = new OAuthPromptSettings
            {
                ConnectionName = ConnectionName,
                Text = "Please Sign In",
                Title = "Sign In",
                Timeout = 300000, // User has 5 minutes to login (1000 * 60 * 5)
            };
            AddDialog(new OAuthPrompt(nameof(OAuthPrompt), oauthPromptSettings));
        }

        protected override async Task<DialogTurnResult> OnContinueDialogAsync(DialogContext innerDc, CancellationToken cancellationToken = default)
        {
            var result = await InterruptAsync(innerDc, cancellationToken);
            if (result != null)
            {
                return result;
            }

            return await base.OnContinueDialogAsync(innerDc, cancellationToken);
        }

        private async Task<DialogTurnResult> InterruptAsync(DialogContext innerDc, CancellationToken cancellationToken)
        {
            if (innerDc.Context.Activity.Type == ActivityTypes.Message)
            {
                var text = innerDc.Context.Activity.Text?.ToLowerInvariant();

                switch (text)
                {
                    case "help":
                    case "?":
                        var helpMessage = MessageFactory.Text(HelpMsgText, HelpMsgText, InputHints.ExpectingInput);
                        await innerDc.Context.SendActivityAsync(helpMessage, cancellationToken);
                        return new DialogTurnResult(DialogTurnStatus.Waiting);

                    case "cancel":
                    case "quit":
                        var cancelMessage = MessageFactory.Text(CancelMsgText, CancelMsgText, InputHints.IgnoringInput);
                        await innerDc.Context.SendActivityAsync(cancelMessage, cancellationToken);
                        return await innerDc.CancelAllDialogsAsync(cancellationToken);
                }

                // Allow logout anywhere in the command
                if (text?.IndexOf("logout") >= 0)
                {
                    // The bot adapter encapsulates the authentication processes.
                    var botAdapter = (BotFrameworkAdapter)innerDc.Context.Adapter;
                    await botAdapter.SignOutUserAsync(innerDc.Context, ConnectionName, null, cancellationToken);
                    await innerDc.Context.SendActivityAsync(MessageFactory.Text("You have been signed out."), cancellationToken);
                    return await innerDc.CancelAllDialogsAsync(cancellationToken);
                }
            }

            return null;
        }
    }
}
