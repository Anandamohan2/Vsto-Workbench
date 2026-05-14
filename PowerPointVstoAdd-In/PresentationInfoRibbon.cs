using System;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointVstoAdd_In
{
    public partial class PresentationInfoRibbon
    {
        private void PresentationInfoRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnReadInfo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Get PowerPoint Application
                PowerPoint.Application app = Globals.ThisAddIn.Application;

                // Validate active presentation
                if (app.Presentations.Count == 0)
                {
                    MessageBox.Show(
                        "No active presentation found.",
                        "Validation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // Get active presentation
                PowerPoint.Presentation presentation = app.ActivePresentation;

                StringBuilder builder = new StringBuilder();

                // Presentation Name
                builder.AppendLine("Presentation Name:");
                builder.AppendLine(presentation.Name);
                builder.AppendLine();

                // Slide Count
                builder.AppendLine("Total Slide Count:");
                builder.AppendLine(presentation.Slides.Count.ToString());
                builder.AppendLine();

                // Active Slide Number
                if (app.ActiveWindow != null &&
                    app.ActiveWindow.View != null &&
                    app.ActiveWindow.View.Slide != null)
                {
                    int slideNumber =
                        app.ActiveWindow.View.Slide.SlideIndex;

                    builder.AppendLine("Active Slide Number:");
                    builder.AppendLine(slideNumber.ToString());
                }
                else
                {
                    builder.AppendLine("Active Slide Number:");
                    builder.AppendLine("No active slide.");
                }

                builder.AppendLine();

                // Selection Information
                PowerPoint.Selection selection =
                    app.ActiveWindow.Selection;

                if (selection != null &&
                    selection.Type ==
                    PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange =
                        selection.ShapeRange;

                    builder.AppendLine("Selected Shape Count:");
                    builder.AppendLine(shapeRange.Count.ToString());

                    builder.AppendLine();

                    builder.AppendLine("Selected Shape Details:");
                    builder.AppendLine();

                    for (int i = 1; i <= shapeRange.Count; i++)
                    {
                        PowerPoint.Shape shape = shapeRange[i];

                        builder.AppendLine(
                            "Shape Name : " + shape.Name);

                        builder.AppendLine(
                            "Shape Type : " + shape.Type);

                        builder.AppendLine(
                            "--------------------------------");
                    }
                }
                else
                {
                    builder.AppendLine("Selected Shape Count:");
                    builder.AppendLine("0");

                    builder.AppendLine();

                    builder.AppendLine("No shapes selected.");
                }

                // Show Result
                MessageBox.Show(
                    builder.ToString(),
                    "Presentation Information",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}