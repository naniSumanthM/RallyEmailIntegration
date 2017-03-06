using Newtonsoft.Json;
using Slack;

public class Payload
{
    //This class serializes into the Json payload required by Slack Incoming WebHooks

    [JsonProperty(SlackConstant.channel)]
    public string Channel { get; set; }

    [JsonProperty(SlackConstant.username)]
    public string Username { get; set; }

    [JsonProperty(SlackConstant.text)]
    public string Text { get; set; }
}