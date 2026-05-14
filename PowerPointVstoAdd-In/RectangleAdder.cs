using System;
using System.Diagnostics;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointVstoAdd_In
{
    public class Exercise2RectangleAdder
    {
        /// <summary>
        /// Executes Exercise 2 rectangle creation workflow.
        /// </summary>
        public void Execute(
            PowerPoint.Application application)
        {
            LogInfo(
                "START : Exercise 2");

            try
            {
                if (!ValidateApplication(application))
                {
                    return;
                }

                if (!HasActiveSlide(application))
                {
                    ShowWarning(
                        "No active slide available.");

                    return;
                }

                PowerPoint.Slide activeSlide =
                    application.ActiveWindow
                               .View
                               .Slide;

                AddRectangle(activeSlide);

                ShowInformation(
                    "Rectangle added successfully.");

                LogInfo(
                    "END : Exercise 2");
            }
            catch (Exception exception)
            {
                LogError(
                    "Exercise 2 failed.",
                    exception);

                ShowError(
                    exception.Message);
            }
        }

        #region Rectangle Logic

        /// <summary>
        /// Adds a rectangle to the active slide.
        /// </summary>
        private void AddRectangle(
            PowerPoint.Slide slide)
        {
            LogInfo(
                "Adding rectangle.");

            PowerPoint.Shape rectangle =
                slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    100,
                    100,
                    300,
                    120);

            ApplyRectangleStyle(rectangle);

            ApplyRectangleText(rectangle);
        }

        /// <summary>
        /// Applies rectangle formatting.
        /// </summary>
        private void ApplyRectangleStyle(
            PowerPoint.Shape rectangle)
        {
            LogInfo(
                "Applying rectangle style.");

            rectangle.Fill.ForeColor.RGB =
                System.Drawing.Color.LightBlue.ToArgb();

            rectangle.Line.ForeColor.RGB =
                System.Drawing.Color.DarkBlue.ToArgb();

            rectangle.Line.Weight = 2;
        }

        /// <summary>
        /// Applies rectangle text.
        /// </summary>
        private void ApplyRectangleText(
            PowerPoint.Shape rectangle)
        {
            LogInfo(
                "Applying rectangle text.");

            try
            {
                if (rectangle.HasTextFrame ==
                    Office.MsoTriState.msoTrue)
                {
                    rectangle.TextFrame.TextRange.Text =
                        "Exercise 2 Rectangle";

                    rectangle.TextFrame.TextRange.Font.Size =
                        20;

                    rectangle.TextFrame.TextRange.Font.Bold =
                        Office.MsoTriState.msoTrue;

                    rectangle.TextFrame.TextRange.ParagraphFormat.Alignment =
                        PowerPoint.PpParagraphAlignment.ppAlignCenter;
                }
            }
            catch (Exception exception)
            {
                LogError(
                    "Failed to apply rectangle text.",
                    exception);

                throw;
            }
        }

        #endregion

        #region Validation

        /// <summary>
        /// Validates PowerPoint application state.
        /// </summary>
        private bool ValidateApplication(
            PowerPoint.Application application)
        {
            try
            {
                if (application == null)
                {
                    ShowWarning(
                        "Application unavailable.");

                    return false;
                }

                if (application.Presentations.Count == 0)
                {
                    ShowWarning(
                        "No active presentation.");

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
                    "Validation failed.",
                    exception);

                return false;
            }
        }

        /// <summary>
        /// Validates active slide availability.
        /// </summary>
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
                    "Active slide validation failed.",
                    exception);

                return false;
            }
        }

        #endregion

        #region Logging

        /// <summary>
        /// Writes informational logs.
        /// </summary>
        private void LogInfo(
            string message)
        {
            Debug.WriteLine(
                "[INFO] " + message);
        }

        /// <summary>
        /// Writes error logs.
        /// </summary>
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

        /// <summary>
        /// Shows information message.
        /// </summary>
        private void ShowInformation(
            string message)
        {
            MessageBox.Show(
                message,
                "Information",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// Shows warning message.
        /// </summary>
        private void ShowWarning(
            string message)
        {
            MessageBox.Show(
                message,
                "Validation",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Shows error message.
        /// </summary>
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
    }
}