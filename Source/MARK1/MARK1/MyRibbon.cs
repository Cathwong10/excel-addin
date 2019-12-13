using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace MARK1
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public MyRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MARK1.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        public void OnVisitorsRoomButton(Office.IRibbonControl control)
        {
            Excel.Window window = control.Context;
            Excel.Application myApplication = window.Application;

            //get current user name 
            string userName = Environment.UserName;

            MessageBox.Show("Generate Visitors Room Seating Plan.");
            var app = new PowerPoint.Application();
            var pres = app.Presentations;
            var file = pres.Open(@"C:\\Users\\"+userName+"\\Desktop\\Addin Development\\Templates\\VisitorsRoomSP.pptx", MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
            
            //get first slide
            PowerPoint.Slide slide = file.Slides[1];

            /* one way to add a text box
            PowerPoint.Shape shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            shape.TextFrame.TextRange.InsertAfter("This text was added by using code.");
            */
                        
            LinkedList<String> guestList = new LinkedList<string>();
            LinkedList<String> hostList = new LinkedList<string>();
            initialiseGuestList(guestList, myApplication);
            initialiseHostList(hostList, myApplication);

            //identify group shapes
            PowerPoint.Shape guestSideGroup = null;
            PowerPoint.Shape hostSideGroup = null;

            //replace title
            foreach (var s in slide.Shapes)
            {
                var shape = (PowerPoint.Shape) s;
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    var textRange = shape.TextFrame.TextRange;
                    var text = textRange.Text;
                    if (text == "<<VISIT_NAME>>")
                    {
                        shape.TextFrame.DeleteText();
                        shape.TextFrame.TextRange.InsertAfter(getVisitName(myApplication).ToString().ToUpper());
                        shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                        shape.TextFrame.TextRange.Font.Underline = MsoTriState.msoTrue;

                    }
                    if (text == "<<VISIT_DATES>>")
                    {
                        shape.TextFrame.DeleteText();
                        shape.TextFrame.TextRange.InsertAfter(getVisitDates(myApplication).ToString().ToUpper());
                        shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                        shape.TextFrame.TextRange.Font.Underline = MsoTriState.msoTrue;
                    }
                }                

                //identify group shapes for host and guest groups
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach(var item in shape.GroupItems)
                    {
                        var subShape = (PowerPoint.Shape)item;
                        if (subShape.HasTextFrame == MsoTriState.msoTrue)
                        {                            
                            var subTextRange = subShape.TextFrame.TextRange;
                            var subText = subTextRange.Text;

                            if (subText == "RIGHT_ALIGNED_NAME")
                            {
                                guestSideGroup = shape;

                            } else if (subText == "LEFT_ALIGNED_NAME")
                            {
                                hostSideGroup = shape;

                            } else if (subText == "GUEST_NOTETAKER")
                            {
                                //populate notetakers
                                insertNameIntoTextBox(shape, getNoteTaker(guestList));

                            } else if (subText == "HOST_NOTETAKER")
                            {
                                //populate notetakers
                                insertNameIntoTextBox(shape, getNoteTaker(hostList));

                            }
                        }
                    }
                    
                }                
            }
            //done iterating all shapes.
                        
            arrangeGuestSP(guestSideGroup,slide,guestList);
            arrangeHostSP(hostSideGroup, slide, hostList);
            
            //save file
            file.SaveCopyAs(@"C:\\Users\\"+userName+"\\Desktop\\Addin Development\\MyVisitorsRoomSP", PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            file.Close();
            app.Quit();
        }

        public void OnTanglinRoomButton(Office.IRibbonControl control)
        {
            MessageBox.Show("[To be implemented] Generate Room Seating Plan.");
        }

        public void OnNOCButton(Office.IRibbonControl control)
        {
            Excel.Window window = control.Context;
            Excel.Application myApplication = window.Application;

            //get current user name 
            string userName = Environment.UserName;

            MessageBox.Show("Generate NOC.");
            Word.Application app = new Word.Application();
            Document document = app.Documents.Open(@"C:\\Users\\"+userName+"\\Desktop\\Addin Development\\Templates\\NOC.docx");

            //should do this again in case the names changed when preparing NOCs
            //getting delegation list
            LinkedList<String> guestList = new LinkedList<string>();
            LinkedList<String> hostList = new LinkedList<string>();
            initialiseGuestList(guestList, myApplication);
            initialiseHostList(hostList, myApplication);

            //update involvement list at Annex
            updateNocInvolvement(hostList,guestList, document);
            replacePlaceholder(app,"<<HOST_NAME>>",getInfoFromExcel(myApplication,"hostHOD").ToString().ToUpper());
            replacePlaceholder(app, "<<GUEST_NAME>>", getInfoFromExcel(myApplication,"guestHOD").ToString().ToUpper());
            replacePlaceholder(app, "<<COUNTRY_NAME>>", getInfoFromExcel(myApplication, "country"));
     
            document.SaveAs2(@"C:\\Users\\"+userName+"\\Desktop\\Addin Development\\NOC.docx");
            document.Close();
            app.Quit();
        }

        private object getInfoFromExcel(Excel.Application myApplication, string v)
        {
            Excel.Worksheet visitDetailsWS = (Excel.Worksheet)myApplication.Worksheets["VisitDetails"];
            if (v == "hostHOD")
            {
                return visitDetailsWS.get_Range("I11").Text + " " + visitDetailsWS.get_Range("H11").Text;
            } else if (v == "guestHOD")
            {
                return visitDetailsWS.get_Range("D11").Text + " " + visitDetailsWS.get_Range("C11").Text;
            } else if (v == "country")
            {
                return visitDetailsWS.get_Range("F2").Text;
            } else
            {
                return "";
            }
        }


        //substitute placeholder in a document
        private void replacePlaceholder(Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        //update NOC Annex
        private void updateNocInvolvement(LinkedList<string> hostList, LinkedList<string> guestList, Document document)
        {
            Word.Table annexTable = document.Tables[1];
            //find max rows to add
            int i = guestList.Count;
            if (hostList.Count > guestList.Count)
            {
                i = hostList.Count;
            }

            int j = 0;
            while (j < i-1)
            {
                annexTable.Rows.Add();
                j++;
            }

            i = 2;//i is row
            j = 1; 
            Word.Cell cell;
            foreach (string official in hostList)
            {
                cell = annexTable.Cell(i, j);
                cell.Range.Text = official+"\n";
                i++;
            }

            i = 2;//i is row
            j = 2;
            foreach (string official in guestList)
            {
                cell = annexTable.Cell(i, j);
                cell.Range.Text = official + "\n";
                i++;
            }

        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers


        private object getVisitDates(Excel.Application myApplication)
        {
            Excel.Worksheet visitDetailsWS = (Excel.Worksheet)myApplication.Worksheets["VisitDetails"];
            Excel.Range firstNameRange = visitDetailsWS.get_Range("C4");
            Excel.Range secondNameRange = visitDetailsWS.get_Range("C6");
            return firstNameRange.Text + " to " + secondNameRange.Text;
        }

        private object getVisitName(Excel.Application myApplication)
        {
            Excel.Worksheet visitDetailsWS = (Excel.Worksheet)myApplication.Worksheets["VisitDetails"];
            Excel.Range visitName = visitDetailsWS.get_Range("C2");
            return visitName.Text;
        }

        //search a list to find the notetaker
        private string getNoteTaker(LinkedList<string> delegationList)
        {
            foreach (string s in delegationList)
            {
                if (s.ToLower().Contains("note-taker"))
                {

                    return s;
                }
            }
            return "";
        }

        //read from Excel and populate Host List
        private void initialiseHostList(LinkedList<string> hostList, Excel.Application myApplication)
        {
            Excel.Worksheet visitDetailsWS = (Excel.Worksheet)myApplication.Worksheets["VisitDetails"];
            int i = 0;
            string salutation = "";
            string name = "";
            string designation = "";
            string org = "";
            string official;

            //get the first host name
            Excel.Range firstNameRange = visitDetailsWS.get_Range("H11");
            name = firstNameRange.Text;
            while (name != "")
            {
                salutation = visitDetailsWS.get_Range("G" + (11 + i)).Text;
                designation = visitDetailsWS.get_Range("I" + (11 + i)).Text;
                org = visitDetailsWS.get_Range("J" + (11 + i)).Text;
                official = salutation + " " + name + "\n" + designation + "\n" + org;
                hostList.AddLast(official);

                i++;
                name = visitDetailsWS.get_Range("H" + (11 + i)).Text;
            }
        }

        //read from Excel and populate Guest List
        private void initialiseGuestList(LinkedList<string> guestList, Excel.Application myApplication)
        {
            Excel.Worksheet visitDetailsWS = (Excel.Worksheet)myApplication.Worksheets["VisitDetails"];
            int i = 0;
            string salutation = "";
            string name = "";
            string designation = "";
            string org = "";
            string official;

            //get the first host name
            Excel.Range firstNameRange = visitDetailsWS.get_Range("C11");
            name = firstNameRange.Text;
            while (name != "")
            {
                salutation = visitDetailsWS.get_Range("B" + (11 + i)).Text;
                designation = visitDetailsWS.get_Range("D" + (11 + i)).Text;
                org = visitDetailsWS.get_Range("E" + (11 + i)).Text;
                official = salutation + " " + name + "\n" + designation + "\n" + org;
                guestList.AddLast(official);

                i++;
                name = visitDetailsWS.get_Range("C" + (11 + i)).Text;
            }
        }

        //arrange seating arrangement at the host side
        private void arrangeHostSP(PowerPoint.Shape hostSideGroup, PowerPoint.Slide slide, LinkedList<string> hostList)
        {
            //copy first group
            hostSideGroup.Copy();

            //get coordinate of first group
            float refGroupTop = hostSideGroup.Top;
            float refGroupLeft = hostSideGroup.Left;

            //populate the first group
            insertNameIntoTextBox(hostSideGroup, hostList.First.Value);

            //continue with the rest of the delegation
            PowerPoint.Shape s;
            int counter = 0;
            foreach (string official in hostList)
            {
                //note-takers are out of scope for SP as they have special seats (TODO: do the same for interpreters)
                if (official.ToLower().Contains("note-taker"))
                {
                    continue;
                }

                counter++;
                if (counter == 1)
                {
                    //the first man is done
                    continue;
                }
                else
                {
                    s = slide.Shapes.Paste()[1];
                    s.Top = refGroupTop + 60 * (counter - 1);
                    s.Left = refGroupLeft;
                    insertNameIntoTextBox(s, official);
                }

            }

        }

        //arrange guest seating plan
        private void arrangeGuestSP(PowerPoint.Shape guestSideGroup, PowerPoint.Slide slide, LinkedList<string> guestList)
        {
            //copy first group
            guestSideGroup.Copy();

            //get coordinate of first group
            float refGroupTop = guestSideGroup.Top;
            float refGroupLeft = guestSideGroup.Left;

            //populate the first group
            insertNameIntoTextBox(guestSideGroup, guestList.First.Value);

            //continue with the rest of the delegation
            PowerPoint.Shape s;
            int counter = 0;
            foreach (string official in guestList)
            {
                //note-takers are out of scope for SP as they have special seats (TODO: do the same for interpreters)
                if (official.ToLower().Contains("note-taker"))
                {
                    continue;
                }

                counter++;
                if (counter == 1)
                {
                    //the first man is done
                    continue;
                }
                else
                {
                    s = slide.Shapes.Paste()[1];
                    s.Top = refGroupTop + 60 * (counter - 1);
                    s.Left = refGroupLeft;
                    insertNameIntoTextBox(s, official);
                }

            }
        }

        //get a group of textbox objects and populte the name into the textbox in the group
        private void insertNameIntoTextBox(PowerPoint.Shape shape, string name)
        {
            //identify group shapes for host and guest groups
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                //search for left or right aligned named key words text boxes
                foreach (var item in shape.GroupItems)
                {
                    var subShape = (PowerPoint.Shape)item;
                    if (subShape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        var subTextRange = subShape.TextFrame.TextRange;
                        var subText = subTextRange.Text;

                        if (subText == "RIGHT_ALIGNED_NAME" || subText == "LEFT_ALIGNED_NAME" || subText == "HOST_NOTETAKER" || subText == "GUEST_NOTETAKER")
                        {
                            subShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                            subShape.TextFrame.TextRange.Text = name;
                            subShape.TextFrame.TextRange.Paragraphs(1).Lines(1, 1).Font.Bold = MsoTriState.msoTrue;
                        }

                    }
                }

            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
