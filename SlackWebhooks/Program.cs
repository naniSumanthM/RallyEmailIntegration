using System;
using System.Collections.Generic;
using Slack.Webhooks;

namespace SlackWebhooks
{
    class Program
    {
        static void Main(string[] args)
        {
            SlackClient client = new SlackClient(SlackConstant.webhookURL);
            SlackMessage message = new SlackMessage
            {
                Channel = "#random",
                Text = SlackConstant.title,
                IconEmoji = Emoji.SmallRedTriangle,
                Username = SlackConstant.username
            };

            SlackAttachment attachment = new SlackAttachment
            {
                Fallback = "New open task [Urgent]: <http://url_to_task|Test out Slack message attachments>",
                Text = "Userstory US667: <https://rally1.rallydev.com/#/36903994832ud/detail/userstory/96328719420 | User Story Title >",
                Color = "#4ef442",
                Fields = new List<SlackField>
                        {
                            new SlackField
                                {
                                    Value = "User Story Description"
                                }
                        }
            };

            message.Attachments = new List<SlackAttachment> { attachment };
            client.Post(message);

            Console.ReadLine();
        }
    }
}
