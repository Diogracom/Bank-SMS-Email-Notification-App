namespace eNotieService
{
    partial class Service1
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.callAPI = new System.ComponentModel.BackgroundWorker();
            // 
            // callAPI
            // 
            this.callAPI.DoWork += new System.ComponentModel.DoWorkEventHandler(this.callAPI_DoWork);
            // 
            // Service1
            // 
            this.ServiceName = "eNotifications";

        }

        #endregion

        private System.ComponentModel.BackgroundWorker callAPI;
    }
}
