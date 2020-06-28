using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Controls;
using Utilites;

namespace PowerPointGenerator
{
    public class PowerPointReports
    {
        public string TemplateFile { get; set; }
        public string NameListFile { get; set; }
        public string BaseSettingsFolder { get; set; }  // folder where above two files are stored

        public string SourceFolder { get; set; }
        public String DestinationFolder { get; set; }

        public TextBox ProcessTextBox { get; set; }
        public TextBox ProcessedCount { get; set; }

        private int SlideWidthCentre;
        private int SlideHeightCentre;

        public PowerPointReports()
        {
            
        }

        public List<string> GetNamesList()
        {
            List<string> names = new List<string>();

            string fileName = Path.Combine(BaseSettingsFolder, NameListFile);
            StreamReader reader = new StreamReader(fileName);

            names = reader.ReadToEnd()
                                    .Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries)                                    
                                    .ToList();
            return names;
        }

        public void GetBaseCounts()
        {

        }

        public int CreatePresentations()
        {
            int namesProcessed = 0;
            ProcessTextBox.Text = string.Empty;


            //  this is also the trust ID                
            List<string> names = GetNamesList();

            ProcessTextBox.AppendText("Processing: ");


            if (BuildPowerPointFile(names))
            {
                ProcessTextBox.AppendText(" - done");
                namesProcessed++;
            }

            ProcessTextBox.AppendText(Environment.NewLine);
            ProcessTextBox.ScrollToEnd();
            //ProcessedCount.Text = namesProcessed.ToString();

            ScreenEvents.DoEvents();


            return namesProcessed;
        }

        private bool BuildPowerPointFile(List<string> names)
        {
            //  open the base presentation
            // create the animations
            //  save as ... to the destination folder
            
            PowerPoint.Application ppApplication = null;
            PowerPoint.Presentations ppPresentations = null;
            PowerPoint.Presentation ppPresentation = null;

            try
            {
                ppApplication = new PowerPoint.Application();
                ppPresentations = ppApplication.Presentations;

                //  to create a new presentation
                ppPresentation = ppPresentations.Add(MsoTriState.msoTrue);
                ppApplication.Activate();

                SlideWidthCentre = (int)ppPresentation.PageSetup.SlideWidth / 2;
                SlideHeightCentre = (int)ppPresentation.PageSetup.SlideHeight / 2;

                ppPresentation.ApplyTemplate(Path.Combine(BaseSettingsFolder, TemplateFile));

                AddTitleSlide(ppPresentation, "My trust", "2015");

                AddAnimationNames(ppPresentation, names);

                //CentrePictures(ppPresentation);

                ppPresentation.SaveAs(Path.Combine(DestinationFolder, "AnimatedNames"),
                                                    PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

            }
            catch (Exception e)
            {
                ProcessTextBox.AppendText(Environment.NewLine);
                ProcessTextBox.AppendText("Error: " + e.Message);

                return false;
            }
            finally
            {
                try
                {
                    if (ppPresentation != null)
                    {
                        ppPresentation.Close();
                        ppApplication.Quit();

                        ppApplication = null;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                }
                catch (Exception e)
                {

                }
            }

            return true;
        }

        internal void ClearDestinationFolder()
        {
            Array.ForEach(Directory.GetFiles(DestinationFolder), File.Delete);
        }

        private void CentrePictures(PowerPoint.Presentation ppPresentation)
        {
            foreach (PowerPoint.Slide slide in ppPresentation.Slides)
            {
                // centre picture
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoPicture)
                    {
                        shape.Left = SlideWidthCentre - (shape.Width / 2);
                        shape.Top = SlideHeightCentre - (shape.Height / 2);
                    }
                }
            }
        }

        private void AddNoteSlide(PowerPoint.Presentation ppPresentation)
        {
            PowerPoint.Slide ppSlide = AddASlide(ppPresentation, "Mbrrace_note_slide");

        }


        private void AddAnimationNames(PowerPoint.Presentation ppPresentation, List<string> names)
        {
            PowerPoint.Slide ppSlide = AddASlide(ppPresentation, "Mbrrace picture slide");

            int idx = 1;
            foreach (string name in names)
            {
                AddAName(ppSlide, name, idx, names.Count());
                idx++;
            }

            idx = 1;

            foreach (PowerPoint.Shape shape in ppSlide.Shapes)
            {
                shape.AnimationSettings.AnimationOrder = idx++;
            }

        }

        private void AddAName(PowerPoint.Slide ppSlide, string name, int idx, int nameCount)
        {

            float left = 0;
            float top = 0;
            float radius = SlideHeightCentre;
            float offset = SlideHeightCentre / SlideWidthCentre;
            double t = 2 * Math.PI * idx / nameCount;
            left = (float)(SlideWidthCentre + radius * Math.Cos(t) * 1.5);
            top = (float)(SlideHeightCentre + radius * Math.Sin(t) * 0.7);

            PowerPoint.Shape tempShape = ppSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, 120, 60);
            tempShape.TextFrame.TextRange.Text = name;

            tempShape.AnimationSettings.Animate = MsoTriState.msoTrue;
            if (idx==1)
            {
                tempShape.AnimationSettings.AdvanceMode = PowerPoint.PpAdvanceMode.ppAdvanceOnClick;
            }
            else
            {
                tempShape.AnimationSettings.AdvanceMode = PowerPoint.PpAdvanceMode.ppAdvanceOnTime;
            }
            tempShape.AnimationSettings.AfterEffect = PowerPoint.PpAfterEffect.ppAfterEffectHide;
        }

        private void AddPictureSlide(PowerPoint.Presentation ppPresentation, string trustCode, string file)
        {
            PowerPoint.Slide ppSlide = AddASlide(ppPresentation, "Mbrrace picture slide");

            ppSlide.Shapes.AddPicture(Path.Combine(SourceFolder, trustCode, file), MsoTriState.msoFalse, MsoTriState.msoTrue, 1, 1);


        }

        private void AddTitleSlide(PowerPoint.Presentation ppPresentation, string trustName, string year)
        {
            PowerPoint.Slide ppSlide = null;
            PowerPoint.Shape ppShape = null;

            ppSlide = AddASlide(ppPresentation, "Mbrrace title slide");
            ppShape = GetAShape(ppSlide, "Title 2");
            ppShape.TextFrame.TextRange.Text = trustName;

            ppShape = GetAShape(ppSlide, "Subtitle 1");
            ppShape.TextFrame.TextRange.Text = "Perinatal mortality report: " + year + " births";

        }


        #region "Helper functions"

        private PowerPoint.Slide AddASlide(PowerPoint.Presentation ppPresentation, string slideLayout)
        {
            return AddASlide(ppPresentation, ppPresentation.Slides.Count + 1, slideLayout);
        }

        private PowerPoint.Slide AddASlide(PowerPoint.Presentation ppPresentation, int slidePosition, string slideLayout)
        {
            PowerPoint.Slide ppSlide = null;

            ppSlide = ppPresentation.Slides.AddSlide(slidePosition, GetCustomLayout(ppPresentation, slideLayout));

            return ppSlide;
        }

        private PowerPoint.CustomLayout GetCustomLayout(PowerPoint.Presentation ppPresentation, string slideLayout)
        {
            PowerPoint.CustomLayout ppCustomLayout = null;

            foreach (PowerPoint.CustomLayout customLayout in ppPresentation.SlideMaster.CustomLayouts)
            {
                if (customLayout.Name == slideLayout)
                {
                    return customLayout;
                }
            }

            return ppCustomLayout;
        }

        private PowerPoint.Shape GetAShape(PowerPoint.Slide ppSlide, string shapeName)
        {
            PowerPoint.Shape pShape = null;
            foreach (PowerPoint.Shape shape in ppSlide.Shapes)
            {
                if (shape.Name == shapeName)
                {
                    return shape;
                }
            }

            return pShape;
        }

        #endregion
    }
}
