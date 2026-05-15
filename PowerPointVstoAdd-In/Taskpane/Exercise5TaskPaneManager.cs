using System;
using Microsoft.Office.Tools;
using PowerPointVstoAdd_In.TaskPane;

namespace PowerPointVstoAdd_In
{
    public class Exercise5TaskPaneManager
    {
        private CustomTaskPane taskPane;

        /// <summary>
        /// Shows task pane.
        /// </summary>
        public void ShowTaskPane()
        {
            try
            {
                if (taskPane != null)
                {
                    taskPane.Visible = true;

                    return;
                }

                BasicTaskPaneControl control =
                    new BasicTaskPaneControl(
                        Globals.ThisAddIn.Application);

                control.CloseRequested +=
                    Control_CloseRequested;

                taskPane =
                    Globals.ThisAddIn.CustomTaskPanes.Add(
                        control,
                        "Basic Task Pane");

                taskPane.Width = 350;

                taskPane.Visible = true;
            }
            catch (Exception exception)
            {
                System.Windows.Forms.MessageBox.Show(
                    exception.Message);
            }
        }

        /// <summary>
        /// Handles close requests.
        /// </summary>
        private void Control_CloseRequested(
            object sender,
            EventArgs e)
        {
            if (taskPane != null)
            {
                taskPane.Visible = false;
            }
        }
    }
}