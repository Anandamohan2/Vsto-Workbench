using System;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointVstoAdd_In
{
    public class Exercise4GroupShapeTraversal
    {
        /// <summary>
        /// Executes recursive group traversal workflow.
        /// </summary>
        public void Execute(
            PowerPoint.Application application)
        {
            LogInfo(
                "START : Exercise 4");

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

                TraversePresentation(
                    presentation,
                    builder);

                ShowInformation(
                    builder.ToString());

                LogInfo(
                    "END : Exercise 4");
            }
            catch (Exception exception)
            {
                LogError(
                    "Exercise 4 failed.",
                    exception);

                ShowError(
                    exception.Message);
            }
        }

        #region Traversal Logic

        /// <summary>
        /// Traverses all slides.
        /// </summary>
        private void TraversePresentation(
            PowerPoint.Presentation presentation,
            StringBuilder builder)
        {
            foreach (PowerPoint.Slide slide
                in presentation.Slides)
            {
                builder.AppendLine(
                    "================================");

                builder.AppendLine(
                    "Slide Number : " +
                    slide.SlideIndex);

                builder.AppendLine(
                    "================================");

                TraverseShapes(
                    slide.Shapes,
                    builder,
                    0);
            }
        }

        /// <summary>
        /// Recursively traverses shapes.
        /// </summary>
        private void TraverseShapes(
            PowerPoint.Shapes shapes,
            StringBuilder builder,
            int level)
        {
            foreach (PowerPoint.Shape shape
                in shapes)
            {
                AppendShapeInformation(
                    shape,
                    builder,
                    level);

                if (shape.Type ==
                    Office.MsoShapeType.msoGroup)
                {
                    LogInfo(
                        "Group shape found : " +
                        shape.Name);

                    TraverseGroupItems(
                        shape,
                        builder,
                        level + 1);
                }
            }
        }

        /// <summary>
        /// Traverses grouped items recursively.
        /// </summary>
        private void TraverseGroupItems(
            PowerPoint.Shape groupShape,
            StringBuilder builder,
            int level)
        {
            try
            {
                PowerPoint.GroupShapes groupItems =
                    groupShape.GroupItems;

                for (int index = 1;
                     index <= groupItems.Count;
                     index++)
                {
                    PowerPoint.Shape childShape =
                        groupItems[index];

                    AppendShapeInformation(
                        childShape,
                        builder,
                        level);

                    if (childShape.Type ==
                        Office.MsoShapeType.msoGroup)
                    {
                        TraverseGroupItems(
                            childShape,
                            builder,
                            level + 1);
                    }
                }
            }
            catch (Exception exception)
            {
                LogError(
                    "Failed to traverse group.",
                    exception);
            }
        }

        /// <summary>
        /// Appends shape information.
        /// </summary>
        private void AppendShapeInformation(
            PowerPoint.Shape shape,
            StringBuilder builder,
            int level)
        {
            string indent =
                new string(' ', level * 4);

            builder.AppendLine(
                indent +
                "Shape Name : " +
                shape.Name);

            builder.AppendLine(
                indent +
                "Shape Type : " +
                shape.Type);

            builder.AppendLine(
                indent +
                "Is Group : " +
                (shape.Type ==
                 Office.MsoShapeType.msoGroup));

            builder.AppendLine(
                indent +
                "Left : " +
                shape.Left);

            builder.AppendLine(
                indent +
                "Top : " +
                shape.Top);

            builder.AppendLine(
                indent +
                "Width : " +
                shape.Width);

            builder.AppendLine(
                indent +
                "Height : " +
                shape.Height);

            builder.AppendLine(
                indent +
                "--------------------------------");
        }

        #endregion

        #region Validation

        /// <summary>
        /// Validates PowerPoint application.
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

        private void LogInfo(
            string message)
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
                "Exercise 4 Report",
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
    }
}