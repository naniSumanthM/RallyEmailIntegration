using System;

namespace Slack
{
    class Test
    {
        static void Main(string[] args)
        {
            SlackClient client = new SlackClient(SlackConstant.webhookURL);

            for (int i = 0; i < 5; i++)
            {
                client.PostMessage
                   (
                       username: "sumanth",
                       text: "Notification Service"+i,
                       channel: "#general"
                   );
            }

            Console.ReadLine();
        }
    }
}
