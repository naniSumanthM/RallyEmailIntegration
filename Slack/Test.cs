using System;

namespace Slack
{
    class Test
    {
        static void Main(string[] args)
        {
            SlackClient client = new SlackClient(SlackConstant.webhookURL);
            client.PostMessage(username: "sumanth", text: "SlackBot", channel: "#general");



            Console.ReadLine();
        }
    }
}


//can iterate throgh the list of messages, create the user story, and create a slack notification with the title of the user story
//THis might become spammy becuase users cannot see the story from slack.

//Purpose: 