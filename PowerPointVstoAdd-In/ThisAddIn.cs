using System;
using System.Diagnostics;

namespace PowerPointVstoAdd_In
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(
            object sender,
            EventArgs e)
        {
            Debug.WriteLine(
                "[INFO] PowerPoint Add-In Started");
        }

        private void ThisAddIn_Shutdown(
            object sender,
            EventArgs e)
        {
            Debug.WriteLine(
                "[INFO] PowerPoint Add-In Shutdown");
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup +=
                new EventHandler(
                    ThisAddIn_Startup);

            this.Shutdown +=
                new EventHandler(
                    ThisAddIn_Shutdown);
        }

        #endregion
    }
}