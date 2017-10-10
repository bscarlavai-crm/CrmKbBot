using CrmKbBot.Service;
using HtmlToMarkdown.Net;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace CrmKbBot.Dialog
{
    [Serializable]
    public class CaseDialog : IDialog<object>
    {
        private bool firstPrompt = true;
        private string lastSearch;

        public Task StartAsync(IDialogContext context)
        {
            context.Wait(this.StartDialog);
            return Task.CompletedTask;
        }
        public Task StartDialog(IDialogContext context, IAwaitable<IMessageActivity> input)
        {
            return InitialPrompt(context);
        }

        protected virtual Task InitialPrompt(IDialogContext context)
        {
            string prompt = "What can I help you with today?";

            if (!this.firstPrompt)
            {
                prompt = "What else can I help you with?";
            }

            this.firstPrompt = false;

            PromptDialog.Text(context, this.Search, prompt);
            return Task.CompletedTask;
        }

        public async Task Search(IDialogContext context, IAwaitable<string> input)
        {
            string text = input != null ? await input : null;

            var client = new CrmSearchClient();
            var result = await client.SearchAsync(text);

            var cards = result.EntityCollection.Entities.Select(e => ToThumbNailCard(e));

            var replyToConversation = context.MakeMessage();
            replyToConversation.Text = "Here are some potential solutions...";
            replyToConversation.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            replyToConversation.Attachments = new List<Attachment>();

            foreach (var card in cards)
            {
                replyToConversation.Attachments.Add(card.ToAttachment());
            }

            await context.PostAsync(replyToConversation);

            PromptDialog.Confirm(context, this.SolveProblem, "Did any of the articles solve your issue?");
        }

        private async Task SolveProblem(IDialogContext context, IAwaitable<bool> input)
        {
            try
            {
                bool solveProblem = await input;
                if (solveProblem)
                {
                    await context.PostAsync("Thank you for contacting support. Let us know if there is anything else I can help with.");
                    context.Done(true);
                }
                else
                {
                    PromptDialog.Confirm(context, this.ShouldCreateCase, "Sorry we couldn't help.  Would you like to create a support case for this issue?");
                }
            }
            catch (TooManyAttemptsException)
            {
                context.Done(true);
            }
        }

        private async Task ShouldCreateCase(IDialogContext context, IAwaitable<bool> input)
        {
            try
            {
                bool createCase = await input;
                if (createCase)
                {
                    var client = new CrmSearchClient();
                    var result = await client.CreateCase(this.lastSearch);
                    await context.PostAsync($"A case has been created. Your case number is {result}");
                    context.Done(true);
                }
                else
                {
                    await context.PostAsync("Thank you for contacting support. Let me know if there is anything else I can help with.");
                    context.Done(true);
                }
            }
            catch (TooManyAttemptsException)
            {
                context.Done(true);
            }
        }

        private ThumbnailCard ToThumbNailCard(Microsoft.Xrm.Sdk.Entity kbArticle)
        {
            var publicNumber = kbArticle.GetAttributeValue<string>("articlepublicnumber");

            HtmlToMarkdownConverter converter = new HtmlToMarkdownConverter();
            var content = converter.Convert(kbArticle.GetAttributeValue<string>("content"));

            return new ThumbnailCard()
            {
                Title = kbArticle.GetAttributeValue<string>("title"),
                Buttons = new[] { new CardAction(ActionTypes.OpenUrl, "View Article", value: $"https://portalurl.microsoftcrmportals.com/knowledgebase/article/{publicNumber}/en-us") },
                Text = content
            };
        }
    }
}