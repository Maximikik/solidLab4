using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private ISldWorks app = null;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                app = new SldWorks();
                app.FrameState = (int)swWindowState_e.swWindowMaximized;
                app.Visible = true;
            }
            catch
            {
                try
                {
                   app = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch
                {
                    MessageBox.Show("Не удалось открыть SolidWorks либо найти открытое приложение.");
                    return;
                }
            }
            
            ModelDoc2 doc;
            
            if (app.ActiveDoc == null)
            {
                doc = (ModelDoc2)app.INewDrawing((int)swDwgTemplates_e.swDwgTemplateCustom);
                
                doc.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
            }

            doc = (ModelDoc2)app.ActiveDoc;
            SketchManager sm = doc.SketchManager;
            var tangent = "sgTANGENT";
            int pref_toggle = (int)swUserPreferenceToggle_e.swInputDimValOnCreate;
            app.SetUserPreferenceToggle(pref_toggle, false);

            double dx = 0.1;
            double dy = 0.1;

            //0.01 =10
            //1
            sm.CreateLine(0.04 + dx, 0.04 + dy, 0, 0.04 + dx, 0.02 + dy, 0);
            doc.ClearSelection();
            sm.CreateLine(0.04 + dx, 0.02 + dy, 0, 0.02 + dx, 0 + dy, 0);
            doc.ClearSelection();
            //2
            sm.CreateLine(0.02 + dx, 0 + dy, 0, 0 + dx, 0.02 + dy, 0);
            doc.ClearSelection();
            //3
            sm.CreateLine(0 + dx, 0.02 + dy, 0, 0 + dx, 0.04 + dy, 0);
            doc.ClearSelection();
            
            sm.CreateArc(0.02 + dx, 0.04 + dy, 0, 0 + dx, 0.04 + dy, 0,  0.04 + dx, 0.04 + dy, 0, 2);
            doc.IAddRadialDimension2(dx - 0.055, dy + 0.01, 0); 
            doc.ClearSelection();
            //4
            sm.CreateLine(0 + dx, 0.04 + dy, 0, -0.02 + dx, 0.04 + dy, 0);
            doc.ClearSelection();
            //5
            sm.CreateLine(-0.02 + dx, 0.04 + dy, 0, -0.02 + dx, 0.08 + dy, 0);
            doc.ClearSelection();
            //6
            sm.CreateLine(0.04 + dx, 0.04 + dy, 0, 0.06 + dx, 0.04 + dy, 0);
            doc.ClearSelection();
            //7
            sm.CreateLine(0.06 + dx, 0.04 + dy, 0, 0.06 + dx, 0.08 + dy, 0);
            doc.ClearSelection();
            //8
            sm.CreateLine(-0.02 + dx, 0.08 + dy, 0, 0 + dx, 0.06 + dy, 0);
            doc.ClearSelection();
            //9
            sm.CreateLine(0.06 + dx, 0.08 + dy, 0, 0.04 + dx, 0.06 + dy, 0);
            doc.ClearSelection();
            //10
            sm.CreateLine(0 + dx, 0.06 + dy, 0, 0.01 + dx, 0.05 + dy, 0);
            doc.ClearSelection();
            //11
            sm.CreateLine(0.04 + dx, 0.06 + dy, 0, 0.03 + dx, 0.05 + dy, 0);
            doc.ClearSelection();


            sm.CreateArc(0.02 + dx, 0.04 + dy, 0, 0.01 + dx, 0.05 + dy, 0, 0.03 + dx, 0.05 + dy, 0, 1);
            doc.IAddRadialDimension2(dx - 0.055, dy + 0.045, 0);
            doc.ClearSelection();
            
            sm.CreateCircleByRadius(0.02 + dx, 0.04 + dy, 0, 0.01);
            doc.IAddDiameterDimension2(dx + 0, dy + 0.075, 0);
            doc.ClearSelection();

            var firstPoint = sm.CreatePoint(-0.02 + dx, 0.04 + dy, 0);
            var secondPoint = sm.CreatePoint(0.06 + dx, 0.04 + dy, 0);
            firstPoint.Select(true);
            secondPoint.Select(true);
            doc.IAddHorizontalDimension2(dx + 0.02, dy - 0.02, 0);
            doc.ClearSelection();

            var firstPoint1 = sm.CreatePoint(0.06 + dx, 0.04 + dy, 0);
            var secondPoint1 = sm.CreatePoint(0.06 + dx, 0.08 + dy, 0);
            firstPoint1.Select(true);
            secondPoint1.Select(true);
            doc.IAddVerticalDimension2(dx + 0.08, dy + 0.06, 0);
            doc.ClearSelection();


            var firstPoint2 = sm.CreatePoint(0.04 + dx, 0.04 + dy, 0);
            var secondPoint2 = sm.CreatePoint(0.06 + dx, 0.04 + dy, 0);
            firstPoint2.Select(true);
            secondPoint2.Select(true);
            doc.IAddHorizontalDimension2(dx + 0.05, dy , 0);
            doc.ClearSelection();

            var firstPoint3 = sm.CreatePoint(0.02 + dx, 0 + dy, 0);
            var secondPoint3 = sm.CreatePoint(dx, 0.02 + dy, 0);
            firstPoint3.Select(true);
            secondPoint3.Select(true);
            doc.IAddVerticalDimension2(dx - 0.01, dy + 0.01, 0);
            doc.ClearSelection();
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            app.CloseAllDocuments(true);
            app.ExitApp();
            app = null;

            try
            {
                Process[] processes = Process.GetProcessesByName("swvbaserver");
                foreach (var process in processes)
                    process.Kill();
            }
            catch { }
        }
    }
}
