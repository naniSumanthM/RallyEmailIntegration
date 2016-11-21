using System;
using System.Collections.Generic;
using Rally.RestApi;
using System.Net;
using ActiveUp.Net.Mail;
using Rally.RestApi.Json;
using Rally.RestApi.Response;

namespace Rally_Email_Integration
{
    class Program
    {
        //Objects
        RallyRestApi _api;
        Imap4Client _imap;

        //Constants
        public const string RallyServerName = RallyField.serverID;
        public const string OutlookServerName = RallyField.outlookImapHost;
        public const int OutlookServerPort = RallyField.outlookImapPort;

        //properties
        public string RallyUserName { get; set; }
        public string RallyPassword { get; set; }

        //Default Constructor
        public Program(string rallyUserName, string rallyPassword)
        {
            _imap = new Imap4Client();
            _api = new RallyRestApi();

            this.RallyUserName = rallyUserName;
            this.RallyPassword = rallyPassword;
            this.EnsureRallyIsAuthenticated();
        }

        private void EnsureRallyIsAuthenticated()
        {
            if (this._api.AuthenticationState != RallyRestApi.AuthenticationResult.Authenticated)
            {
                _api.Authenticate(this.RallyUserName, this.RallyPassword, RallyServerName, null, RallyField.allowSSO);
            }
        }


        public void syncUserStories()
        {
            //Authenticate Outlook
            this.EnsureRallyIsAuthenticated();

            List<Message> unreadList = new List<Message>();


            try
            {
                //Authenticate
                _imap = new Imap4Client();
                _imap.ConnectSsl(OutlookServerName, OutlookServerPort);
                _imap.Login("sumanthmaddirala@outlook.com","iYmcmb24");

                //setup Enviornment
                Mailbox inbox = _imap.SelectMailbox(Credential.inboxFolder);
                int[] unread = inbox.Search("UNSEEN");
                Console.WriteLine("Unread Messages: " + unread.Length);

                //Crawl through the inbox and parse unread subject lines, then move those email objects to a folder
                if (unread.Length > 0)
                {
                    //Add the unread emails to a collection
                    for (int i = 0; i < unread.Length; i++)
                    {
                        Message msg = inbox.Fetch.MessageObject(unread[i]);
                        unreadList.Add(msg);
                    }

                    //print out the unread subejct line
                    foreach (var item in unreadList)
                    {
                        Console.WriteLine(item.Subject);
                        //Console.WriteLine(item.BodyText.Text);
                    }

                    //for (int i = 0; i < unreadList.Count; i++)
                    //{
                    //    Console.WriteLine("Creating");
                    //    //Problem may occur when converting the Message DATATYPE TO a string
                    //    toCreate[RallyField.nameForWSorUSorTA] = unreadList[i]; //i=Message
                    //    CreateResult cr = _api.Create(RallyField.hierarchicalRequirement, toCreate);
                    //}

                    //Test
                    toCreate[RallyField.nameForWSorUSorTA] = "EmailIntegrationTest"; //i=Message
                    CreateResult cr = _api.Create(RallyField.hierarchicalRequirement, toCreate);


                    //Move messages to the processed folder
                    foreach (var item in unread)
                    {
                        inbox.MoveMessage(item, Credential.processedFolder);
                    }
                    //line could cause an error
                    //Mailbox movedFrom = imap.SelectMailbox(Credential.inboxFolder);
                }
                else
                {
                    Console.WriteLine("No Unread Email");
                }

            }
            catch (Imap4Exception)
            {
                throw new Imap4Exception();
            }
            catch (WebException)
            {
                Console.WriteLine("WebException");
            }
            catch (Exception)
            {
                throw new Exception();
            }
            finally
            {
                _imap.Disconnect();
                unreadList.Clear();
            }
        }

    }
}

