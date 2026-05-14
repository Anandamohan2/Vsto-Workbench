using System;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointVstoAdd_In
{
   
    public class Exercise3ShapeEnumerator
    {
        /// <summary>
        /// Executes Exercise 3 shape enumeration workflow.
        /// </summary>
        public void Execute(
            PowerPoint.Application application)
        {
            LogInfo(
                "START : Exercise 3");

            try
            {
                if (!ValidateApplication(application))
                {
                    return;
                }

                PowerPoint.Presentation presentation =
                    application.ActivePresentation;

                StringBuilder builder =
                    new StringBuilder();

                EnumerateSlides(
                    presentation,
                    builder);

                ShowInformation(
                    builder.ToString());

                LogInfo(
                    "END : Exercise 3");
            }
            catch (Exception exception)
            {
                LogError(
                    "Exercise 3 failed.",
                    exception);

                ShowError(
                    exception.Message);
            }
        }

        #region Enumeration

        /// <summary>
        /// Enumerates all slides.
        /// </summary>
        private void EnumerateSlides(
            PowerPoint.Presentation presentation,
            StringBuilder builder)
        {
            foreach (PowerPoint.Slide slide
                in presentation.Slides)
            {
                EnumerateShapes(
                    slide,
                    builder);
            }
        }

        /// <summary>
        /// Enumerates all shapes in a slide.
        /// </summary>
        private void EnumerateShapes(
            PowerPoint.Slide slide,
            StringBuilder builder)
        {
            LogInfo(
                "Enumerating slide : " +
                slide.SlideIndex);

            foreach (PowerPoint.Shape shape
                in slide.Shapes)
            {
                AppendShapeInformation(
                    slide,
                    shape,
                    builder);
            }
        }

        /// <summary>
        /// Appends shape diagnostic information.
        /// </summary>
        private void AppendShapeInformation(
            PowerPoint.Slide slide,
            PowerPoint.Shape shape,
            StringBuilder builder)
        {
            try
            {
                builder.AppendLine(
                    "Slide Number : " +
                    slide.SlideIndex);

                builder.AppendLine(
                    "Shape Name : " +
                    shape.Name);

                builder.AppendLine(
                    "Shape Type : " +
                    shape.Type);

                builder.AppendLine(
                    "Left : " +
                    shape.Left);

                builder.AppendLine(
                    "Top : " +
                    shape.Top);

                builder.AppendLine(
                    "Width : " +
                    shape.Width);

                builder.AppendLine(
                    "Height : " +
                    shape.Height);

                builder.AppendLine(
                    "Has Text Frame : " +
                    (shape.HasTextFrame ==
                     Office.MsoTriState.msoTrue));

                builder.AppendLine(
                    "Is Group Shape : " +
                    (shape.Type ==
                     Office.MsoShapeType.msoGroup));

                AppendPlaceholderInformation(
                    shape,
                    builder);

                builder.AppendLine(
                    "--------------------------------");
            }
            catch (Exception exception)
            {
                LogError(
                    "Failed to read shape information.",
                    exception);

                builder.AppendLine(
                    "Shape read failed.");

                builder.AppendLine(
                    "--------------------------------");
            }
        }

        /// <summary>
        /// Reads placeholder information safely.
        /// </summary>
        private void AppendPlaceholderInformation(
            PowerPoint.Shape shape,
            StringBuilder builder)
        {
            try
            {
                if (shape.Type ==
                    Office.MsoShapeType.msoPlaceholder)
                {
                    builder.AppendLine(
                        "Placeholder Type : " +
                        shape.PlaceholderFormat.Type);
                }
                else
                {
                    builder.AppendLine(
                        "Placeholder Type : None");
                }
            }
            catch (Exception exception)
            {
                LogError(
                    "Placeholder read failed.",
                    exception);

                builder.AppendLine(
                    "Placeholder Type : Unknown");
            }
        }

        #endregion

        #region Validation

        /// <summary>
        /// Validates PowerPoint state.
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
                "Shape Diagnostics",
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