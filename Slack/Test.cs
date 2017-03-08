using System;

namespace Slack
{
    class Test
    {
        static void Main(string[] args)
        {
            SlackClient client = new SlackClient(SlackConstant.webhookURL);

                client.PostMessage
                   (
                       channel: "#general",
                       username: "sumanth",
                       text: "Need to add some front end"
                   );

            Console.ReadLine();
        }
    }
}
