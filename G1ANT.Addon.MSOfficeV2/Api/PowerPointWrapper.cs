/**
* Copyright(C) G1ANT Ltd, All rights reserved
* Solution G1ANT.Addon.MSOfficeV2, Project G1ANT.Addon.MSOfficeV2
* www.g1ant.com
*
* Licensed under the G1ANT license.
* See License.txt file in the project root for full license information.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using G1ANT.Language;
using System.Runtime.InteropServices;

namespace G1ANT.Addon.MSOfficeV2
{
    public class PowerPointWrapper
    {
        private readonly int id;

        private string path;
        private PowerPoint.Application application = null;
        private PowerPoint.Presentations presentationList = null;
        private PowerPoint.Presentation pres = null;
        PowerPoint.Slides oSlides;
        PowerPoint.Slide oSlide;
        PowerPoint.Shapes oShapes;

        internal PowerPointWrapper()
        {
            id = PowerPointManager.GetNextId();
        }

        public int Id { get { return id; } }

        public void Open(string path)
        {
            application = new PowerPoint.Application();
            application.WindowState = PowerPoint.PpWindowState.ppWindowNormal;
            application.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
            application.Visible = MsoTriState.msoTrue;
            if (string.IsNullOrEmpty(path))
            {
                presentationList = application.Presentations;
                pres = presentationList.Add(MsoTriState.msoTrue);
                oSlides = pres.Slides;
                pres.Application.Activate();
            }
            else
            {
                if (string.IsNullOrEmpty(Path.GetDirectoryName(path)))
                    path = application.ActivePresentation.Path + "\\" + path;
                pres = application.Presentations.Open(path);
                pres.Application.Activate();
            }
            this.path = path;
        }

        public void Show()
        {
            pres.Application.Activate();
            Language.RobotWin32.BringWindowToFront((IntPtr)pres.Application.ActiveWindow.HWND);
        }

        public void InsertSlide(int slidePosition,string titleText,string slideText,string layout )
        {
            PowerPoint.PpSlideLayout slideLayout= PowerPoint.PpSlideLayout.ppLayoutTitle;
            switch (layout)
            {
                case "ContentWithCaption":
                    slideLayout = PowerPoint.PpSlideLayout.ppLayoutTitle;
                    break;
                case "TitleOnly":
                    slideLayout = PowerPoint.PpSlideLayout.ppLayoutTitleOnly;
                    break;
                case "Title":
                    slideLayout = PowerPoint.PpSlideLayout.ppLayoutTitle;
                    break;
                case "ObjectAndText":
                    slideLayout = PowerPoint.PpSlideLayout.ppLayoutObjectAndText;
                    break;
                case "Object":
                    slideLayout = PowerPoint.PpSlideLayout.ppLayoutObject;
                    break;
                default:
                    slideLayout = PowerPoint.PpSlideLayout.ppLayoutObject;
                    break;
            }
            if (slidePosition <= oSlides.Count+1)
                oSlide = oSlides.Add(slidePosition, slideLayout);
            else
            {
                RobotMessageBox.Show("Your slide insert position is incorrect. Please check your number.");
                return;
            }
            oShapes = oSlide.Shapes;
            if (titleText != "")
            {
                PowerPoint.Shape oShape = oShapes[1];
                PowerPoint.TextFrame oTxtFrame = oShape.TextFrame;
                PowerPoint.TextRange oTxtRange = oTxtFrame.TextRange;
                oTxtRange.Text = titleText;
            }
            if (slideText != "")
            {
                if (oShapes[2] != null)
                {
                    PowerPoint.Shape oShape = oShapes[2];
                    PowerPoint.TextFrame oTxtFrame = oShape.TextFrame;
                    PowerPoint.TextRange oTxtRange = oTxtFrame.TextRange;
                    oTxtRange.Text = slideText;
                }
                else RobotMessageBox.Show("There is no element to add your text to. Have you selected appropriate slide layout?");
            }

        }

        public void InsertText(string textToBeInserted, int textPos)
        { 
            if (textToBeInserted != "")
            {
                if (oShapes[textPos] != null)
                {
                    PowerPoint.Shape oShape = oShapes[textPos];
                    PowerPoint.TextFrame oTxtFrame = oShape.TextFrame;
                    PowerPoint.TextRange oTxtRange = oTxtFrame.TextRange;
                    oTxtRange.Text = textToBeInserted;
                }
                else RobotMessageBox.Show("There is no element to add your text to. Have you selected appropriate slide layout?");
            }

        }

        public void InsertPhoto(string photoStr, int left, int top, int width, int height)
        {
            oShapes.AddPicture(photoStr, MsoTriState.msoTrue, MsoTriState.msoTrue, left, top, width, height);
        }

        public void Save(string presName, string path)
          {
              if (string.IsNullOrEmpty(path))
              {
                  pres.SaveAs(presName);
              }
              else
              {
                  if (string.IsNullOrEmpty(Path.GetDirectoryName(path)))
                    path = application.ActivePresentation.Path + "\\" + path;
                else
                  this.path = path;
                  pres.SaveAs(path+presName);
              }
          }
          
          
        private void Application_WindowDeactivate(PowerPoint.Presentation Doc, PowerPoint.DocumentWindow Wn)
        {
            Close();
        }

        public void Close()
          {
            try
            {
                if (pres != null)
                {
                    pres.Close();
                    Marshal.ReleaseComObject(pres);
                }

                application.WindowDeactivate -= Application_WindowDeactivate;
                PowerPointManager.Remove(this);

                bool allowQuit = true;
                if (presentationList != null)
                {
                    allowQuit = presentationList.Count == 0;
                    Marshal.ReleaseComObject(pres);
                }

                if (allowQuit)
                {
                    application.Quit();
                }

                Process[] pros = Process.GetProcesses();
                for (int i = 0; i < pros.Count(); i++)
                {
                    if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
                    {
                        pros[i].Kill();
                    }
                }

                Marshal.ReleaseComObject(application);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception)
            {

                throw;
            }
          }
    }
}
