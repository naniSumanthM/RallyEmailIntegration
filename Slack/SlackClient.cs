using Newtonsoft.Json;
using System;
using System.Collections.Specialized;
using System.Net;
using System.Text;
using Slack;

//A simple C# class to post messages to a Slack channel  
//Note: This class uses the Newtonsoft Json.NET serializer available via NuGet  
public class SlackClient
{
    private readonly Uri _uri;
    private readonly Encoding _encoding = new UTF8Encoding();

    public SlackClient(string urlWithAccessToken)
    {
        _uri = new Uri(urlWithAccessToken);
    }

    //Post a message using simple strings  
    public void PostMessage(string channel, string username, string text)
    {
        Payload payload = new Payload()
        {
            Channel = channel,
            Username = username,
            Text = text
        };

        PostMessage(payload);
    }

    //Post a message using a Payload object  
    public void PostMessage(Payload payload)
    {
        string payloadJson = JsonConvert.SerializeObject(payload);

        using (WebClient client = new WebClient())
        {
            NameValueCollection data = new NameValueCollection();
            data[SlackConstant.payload] = payloadJson;
            var response = client.UploadValues(_uri, SlackConstant.httpPost, data);

            //The response text is usually "ok"  
            string responseText = _encoding.GetString(response);
        }
    }

}