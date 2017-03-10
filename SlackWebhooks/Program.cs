using System;
using System.Collections.Generic;
using Slack.Webhooks;

namespace SlackWebhooks
{
    class Program
    {
        static void Main(string[] args)
        {
            //Slackclient object
            SlackClient client = new SlackClient(SlackConstant.webhookURL,100);

            //Message object
            SlackMessage message = new SlackMessage
            {
                Channel = "#general",
                Text = "*Rally Notification*",
                IconEmoji = Emoji.SmallRedTriangle,
                Username = "sumanth"
            };

            //attachment - https://api.slack.com/docs/message-attachments
            //lets us create links, and add attachments
            var slackAttachment = new SlackAttachment
            {
                Fallback = "New open task [Urgent]: <http://url_to_task|Test out Slack message attachments>",
                //_refUUID + _ref + _refObjectName
                Text = "Userstory US667: <https://rally1.rallydev.com/#/36903994832ud/detail/userstory/96328719420 | User Story Title >",
                Color = "#4ef442",
                Fields =
                new List<SlackField>
                {
                    new SlackField
                        {
                            Value = "User Story Description"
                        }
                }
            };

            //add attachmentList to message
            message.Attachments = new List<SlackAttachment> { slackAttachment };
            //post to slack server
            client.Post(message);


            Console.ReadLine();
        }
    }
}
