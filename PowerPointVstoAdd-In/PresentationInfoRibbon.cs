using System;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointVstoAdd_In
{
    public partial class PresentationInfoRibbon
    {
        private void PresentationInfoRibbon_Load(
            object sender,
            RibbonUIEventArgs e)
        {
            LogInfo("Ribbon loaded.");
        }

        private void btnReadInfo_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            LogInfo(
                "START : Reading presentation information.");

            try
            {
                PowerPoint.Application application =
                    Globals.ThisAddIn.Application;

                if (!ValidatePresentation(application))
                {
                    return;
                }

                PowerPoint.Presentation presentation =
                    application.ActivePresentation;

                StringBuilder builder =
                    new StringBuilder();

                AppendPresentationInformation(
                    presentation,
                    builder);

                AppendActiveSlideInformation(
                    application,
                    builder);

                AppendSelectionInformation(
                    application,
                    builder);

                ShowInformation(builder.ToString());

                LogInfo(
                    "END : Reading presentation information.");
            }
            catch (Exception exception)
            {
                LogError(
                    "Unhandled exception occurred.",
                    exception);

                ShowError(exception.Message);
            }
        }

        private void btnAddRectangle_Click(
            object sender
            ,RibbonControlEventArgs e)
        {
            Exercise2RectangleAdder rectangleAdder =
                new Exercise2RectangleAdder();

            rectangleAdder.Execute(
                Globals.ThisAddIn.Application);
        }

        #region Presentation Information


        private void AppendPresentationInformation(
            PowerPoint.Presentation presentation,
            StringBuilder builder)
        {
            LogInfo(
                "Reading presentation information.");

            builder.AppendLine(
                "Presentation Name:");

            builder.AppendLine(
                presentation.Name);

            builder.AppendLine();

            builder.AppendLine(
                "Total Slide Count:");

            builder.AppendLine(
                presentation.Slides.Count.ToString());

            builder.AppendLine();
        }

        #endregion

        #region Active Slide Information

        private void AppendActiveSlideInformation(
            PowerPoint.Application application,
            StringBuilder builder)
        {
            LogInfo(
                "Reading active slide information.");

            try
            {
                if (!HasActiveSlide(application))
                {
                    builder.AppendLine(
                        "Active Slide Number:");

                    builder.AppendLine(
                        "No active slide.");

                    builder.AppendLine();

                    return;
                }

                int slideNumber =
                    application.ActiveWindow
                               .View
                               .Slide
                               .SlideIndex;

                builder.AppendLine(
                    "Active Slide Number:");

                builder.AppendLine(
                    slideNumber.ToString());

                builder.AppendLine();
            }
            catch (Exception exception)
            {
                LogError(
                    "Failed to read active slide.",
                    exception);

                builder.AppendLine(
                    "Failed to read active slide.");

                builder.AppendLine();
            }
        }

        #endregion

        #region Selection Information

        private void AppendSelectionInformation(
            PowerPoint.Application application,
            StringBuilder builder)
        {
            LogInfo(
                "Reading selection information.");

            try
            {
                if (application.ActiveWindow == null)
                {
                    builder.AppendLine(
                        "No active window.");

                    return;
                }

                PowerPoint.Selection selection =
                    application.ActiveWindow.Selection;

                if (selection == null)
                {
                    builder.AppendLine(
                        "Selection unavailable.");

                    return;
                }

                if (selection.Type !=
                    PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    builder.AppendLine(
                        "Selected Shape Count:");

                    builder.AppendLine("0");

                    builder.AppendLine();

                    builder.AppendLine(
                        "No shapes selected.");

                    return;
                }

                PowerPoint.ShapeRange shapeRange =
                    selection.ShapeRange;

                builder.AppendLine(
                    "Selected Shape Count:");

                builder.AppendLine(
                    shapeRange.Count.ToString());

                builder.AppendLine();

                builder.AppendLine(
                    "Selected Shape Details:");

                builder.AppendLine();

                AppendShapeInformation(
                    shapeRange,
                    builder);
            }
            catch (Exception exception)
            {
                LogError(
                    "Failed to read selection information.",
                    exception);

                builder.AppendLine(
                    "Selection read failed.");
            }
        }

        private void AppendShapeInformation(
            PowerPoint.ShapeRange shapeRange,
            StringBuilder builder)
        {
            for (int index = 1;
                 index <= shapeRange.Count;
                 index++)
            {
                try
                {
                    PowerPoint.Shape shape =
                        shapeRange[index];

                    if (shape == null)
                    {
                        continue;
                    }

                    builder.AppendLine(
                        "Shape Name : " +
                        shape.Name);

                    builder.AppendLine(
                        "Shape Type : " +
                        shape.Type);

                    builder.AppendLine(
                        "Is Group Shape : " +
                        (shape.Type ==
                         Office.MsoShapeType.msoGroup));

                    builder.AppendLine(
                        "Has Text Frame : " +
                        (shape.HasTextFrame ==
                         Office.MsoTriState.msoTrue));

                    builder.AppendLine(
                        "--------------------------------");
                }
                catch (Exception exception)
                {
                    LogError(
                        "Failed to read shape.",
                        exception);

                    builder.AppendLine(
                        "Failed to read one shape.");

                    builder.AppendLine(
                        "--------------------------------");
                }
            }
        }

        #endregion

        #region Validation

        private bool ValidatePresentation(
            PowerPoint.Application application)
        {
            try
            {
                if (application == null)
                {
                    ShowWarning(
                        "PowerPoint application unavailable.");

                    return false;
                }

                if (application.Presentations.Count == 0)
                {
                    ShowWarning(
                        "No active presentation found.");

                    return false;
                }

                if (application.ActivePresentation == null)
                {
                    ShowWarning(
                        "Active presentation unavailable.");

                    return false;
                }

                return true;
            }
            catch (Exception exception)
            {
                LogError(
                    "Presentation validation failed.",
                    exception);

                ShowError(
                    "Presentation validation failed.");

                return false;
            }
        }

        private bool HasActiveSlide(
            PowerPoint.Application application)
        {
            try
            {
                return application.ActiveWindow != null &&
                       application.ActiveWindow.View != null &&
                       application.ActiveWindow.View.Slide != null;
            }
            catch (Exception exception)
            {
                LogError(
                    "Failed to validate active slide.",
                    exception);

                return false;
            }
        }

        #endregion

        #region Logging

        private void LogInfo(string message)
        {
            Debug.WriteLine(
                "[INFO] " + message);
        }

        private void LogError(
            string message,
            Exception exception)
        {
            Debug.WriteLine(
                "[ERROR] " + message);

            Debug.WriteLine(
                exception.Message);

            Debug.WriteLine(
                exception.StackTrace);
        }

        #endregion

        #region Message Helpers

        private void ShowInformation(
            string message)
        {
            MessageBox.Show(
                message,
                "Presentation Information",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void ShowWarning(
            string message)
        {
            MessageBox.Show(
                message,
                "Validation",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private void ShowError(
            string message)
        {
            MessageBox.Show(
                message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        #endregion

        //private void btnAddRectangle_Click_1(
        // object sender, 
        // RibbonControlEventArgs e)
        //{
        //}

        private void btnEnumerateShapes_Click(
       object sender,
       RibbonControlEventArgs e)
        {
            Exercise3ShapeEnumerator shapeEnumerator =
                new Exercise3ShapeEnumerator();

            shapeEnumerator.Execute(
                Globals.ThisAddIn.Application);
        }

        private void btnGroupTraversal_Click(
    object sender,
    RibbonControlEventArgs e)
        {
            Exercise4GroupShapeTraversal traversal =
                new Exercise4GroupShapeTraversal();

            traversal.Execute(
                Globals.ThisAddIn.Application);
        }
    }
}