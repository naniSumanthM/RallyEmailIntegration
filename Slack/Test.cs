using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Slack
{
    class Test
    {
        static void Main(string[] args)
        {
            string urlWithAccessToken = "https://hooks.slack.com/services/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";

            SlackClient client = new SlackClient(urlWithAccessToken);

            client.PostMessage(username: "sumanth", text: "SlackBot", channel: "#general");

            Console.ReadLine();
        }
    }
}
