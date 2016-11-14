using System;

namespace Rally
{
    /// <summary>
    /// Getters and Setters for the Rally Enviornment
    /// </summary>

    internal class RallyEnviornment
    {
        private string workspaceReference;
        public string WorkspaceReference
        {
            get
            {
                return this.workspaceReference;
            }
            set
            {
                if (value.Length > 0)
                {
                    this.workspaceReference = value;
                }
            }
        }

        private string projectReference;
        public string ProjectReference
        {
            get
            {
                return this.projectReference;
            }
            set
            {
                if (value.Length > 0)
                {
                    this.projectReference = value;
                }
            }
        }

        private bool scopingProjectUp;
        public bool ScopingProjectUp
        {
            get
            {
                return this.scopingProjectUp;
            }
            set
            {
                this.scopingProjectUp = value;
            }
        }

        private bool scopingProjectDown;
        public bool ScopingProjectDown
        {
            get
            {
                return this.scopingProjectDown;
            }
            set
            {
                this.scopingProjectDown = value;
            }
        }

    }
}
