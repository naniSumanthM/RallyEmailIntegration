using System;

namespace Rally
{
    /// <summary>
    /// Getters and Setters for the values that can be set
    /// </summary>

    internal class UserCredential
    {
        ///<summary>
        ///Program will help create a userStory by identifying the parameters passed as the workspace, and proeject
        /// </summary>

        private string username;
        public string Username
        {
            get
            {
                return this.username;
            }

            set
            {
                if (value.Length > 0)
                {
                    this.username = value;
                }
            }
        }

        private string password;
        public string Password
        {
            get
            {
                return this.password;
            }
            set
            {
                if (value.Length > 0)
                {
                    this.password = value;
                }
            }
        }

        private string serverID;
        public string ServerID
        {
            get
            {
                return this.serverID;
            }
            set
            {
                if (value.Length > 0)
                {
                    this.serverID = value;
                }
            }
        }

    }

}



