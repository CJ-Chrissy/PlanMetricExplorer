using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

[assembly: AssemblyVersion("1.0.0.1")]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context, System.Windows.Window window, ScriptEnvironment environment)
        {
            // TODO : Add here the code that is called when the script is launched from Eclipse.
            PlanSetup plan = context.PlanSetup;
            if(plan == null)
            {
                MessageBox.Show("No valid plan selected");
                return;
            }
            if (!plan.IsDoseValid)
            {
                MessageBox.Show("The plan selected has no valid dose.");
                return;
            }
            // LINQ
            Structure target = plan.StructureSet.Structures.FirstOrDefault(st=>st.Id.Equals(plan.TargetVolumeID));
            if(target == null)
            {
                MessageBox.Show("Plan contains no target volume");
                return;
            }

            //set up plan metrics
            Dictionary<string, string> planMetrics = new Dictionary<string, string>();
            //add metrics to report.
            planMetrics.Add("Target", target.Id);
            planMetrics.Add("Target Volume", target.Volume.ToString("F1")+"cc");
            planMetrics.Add("Dose per Fraction", plan.DosePerFraction.ToString());
            plan.DoseValuePresentation = DoseValuePresentation.Absolute;
            planMetrics.Add("Max Dose", plan.Dose.DoseMax3D.ToString());

            // MU Ratio = Sum(MU)/DosePerFraction

            double mu = 0;
            foreach (var beam in plan.Beams.Where(b=>!Double.IsNaN(b.Meterset.Value)))
            {
                mu += beam.Meterset.Value;
            }
            double dosePerFraction = mu / plan.DosePerFraction.Dose;
            planMetrics.Add("MU Ratio", dosePerFraction.ToString("F2"));

            var mu2 = plan.Beams.Where(b => !Double.IsNaN(b.Meterset.Value)).Sum(b => b.Meterset.Value);
            planMetrics.Add("MU Ratio 2", (mu2/plan.DosePerFraction.Dose).ToString("F2"));

            // ConfOrmity Index V100% body / target volume.
            Structure body = context.PlanSetup.StructureSet.Structures.FirstOrDefault(st => st.DicomType == "EXTERNAL");
            double v100 = plan.GetVolumeAtDose(body, plan.TotalDose, VolumePresentation.AbsoluteCm3);
            planMetrics.Add("Conformity Index", (v100 / target.Volume).ToString("F2"));
            

            // Gradient Index V50% (body)/V100% (body)
            double v50 = plan.GetVolumeAtDose(body, 0.5*plan.TotalDose, VolumePresentation.AbsoluteCm3);
            planMetrics.Add("Gradient Index", (v50 / v100).ToString("F2"));

            // Homogeneity Index HI = (D5% - D95%) / Prescript
            double D5 = plan.GetDoseAtVolume(target, 5, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
            double D95 = plan.GetDoseAtVolume(target, 95, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
            planMetrics.Add("Homogenety Index", ((D5-D95)/plan.TotalDose.Dose).ToString("F2"));

            window.Width = 450;
            window.Height = 600;
            window.Content = AddMetricsToView(planMetrics);
            
        }
        private FlowDocumentScrollViewer AddMetricsToView(Dictionary<string,string> parameters)
        {
            FlowDocumentScrollViewer flowScroller = new FlowDocumentScrollViewer();
            FlowDocument flowDocument = new FlowDocument() { FontFamily=new FontFamily("Arial")};
            flowDocument.Blocks.Add(new Paragraph(new Run("Dosimetric Plan Metrics")) { TextAlignment = TextAlignment.Center });
            //add values to table.
            Table table = new Table();
            table.RowGroups.Add(new TableRowGroup());
            table.RowGroups.First().Rows.Add(new TableRow());
            table.RowGroups.Last().Rows.Last().Cells.Add(new TableCell(new Paragraph(new Run("Property") { FontWeight = FontWeights.Bold })));
            table.RowGroups.Last().Rows.Last().Cells.Add(new TableCell(new Paragraph(new Run("Value") { FontWeight = FontWeights.Bold })));
            foreach(var metric in parameters)
            {
                table.RowGroups.First().Rows.Add(new TableRow());
                table.RowGroups.Last().Rows.Last().Cells.Add(new TableCell(new Paragraph(new Run(metric.Key))) { BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E,0x52,0x88)) });
                table.RowGroups.Last().Rows.Last().Cells.Add(new TableCell(new Paragraph(new Run(metric.Value))) { BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E,0x52,0x88)) });
            }
            flowDocument.Blocks.Add(table);
            flowScroller.Document = flowDocument;
            return flowScroller;
        }
    }
}
