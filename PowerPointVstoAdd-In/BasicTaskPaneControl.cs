using System;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointVstoAdd_In.TaskPane
{
    public partial class BasicTaskPaneControl : UserControl
    {
        private readonly PowerPoint.Application application;

        public event EventHandler CloseRequested;

        public BasicTaskPaneControl(
            PowerPoint.Application powerpointApplication)
        {
            InitializeComponent();

            application =
                powerpointApplication;
        }

        private void BasicTaskPaneControl_Load(
            object sender,
            EventArgs e)
        {
            RefreshSlideInformation();
        }

        private void btnRefresh_Click(
            object sender,
            EventArgs e)
        {
            RefreshSlideInformation();
        }

        private void btnClose_Click(
            object sender,
            EventArgs e)
        {
            CloseRequested?.Invoke(
                this,
                EventArgs.Empty);
        }

        /// <summary>
        /// Refreshes current slide information.
        /// </summary>
        private void RefreshSlideInformation()
        {
            try
            {
                lblError.Text = "";

                if (application == null)
                {
                    ShowError(
                        "Application unavailable.");

                    return;
                }

                if (application.Presentations.Count == 0)
                {
                    ShowError(
                        "No active presentation.");

                    return;
                }

                if (application.ActiveWindow == null ||
                    application.ActiveWindow.View == null ||
                    application.ActiveWindow.View.Slide == null)
                {
                    ShowError(
                        "No active slide.");

                    return;
                }

                PowerPoint.Slide slide =
                    application.ActiveWindow
                               .View
                               .Slide;

                lblSlideInfo.Text =
                    "Current Slide Number : " +
                    slide.SlideIndex +
                    Environment.NewLine +
                    Environment.NewLine +
                    "Slide Name : " +
                    slide.Name;
            }
            catch (Exception exception)
            {
                Debug.WriteLine(
                    exception.Message);

                ShowError(
                    exception.Message);
            }
        }

        /// <summary>
        /// Displays errors safely.
        /// </summary>
        private void ShowError(
            string message)
        {
            lblError.Text =
                "ERROR : " + message;
        }

        private void lblSlideInfo_Click(object sender, EventArgs e)
        {

        }
    }
}