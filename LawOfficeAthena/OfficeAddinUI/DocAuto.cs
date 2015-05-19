using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace OfficeAddinUI
{
    public partial class DocAuto
    {

        private Dictionary<string, OfficeAddinCustomTaskPane> OfficeAddinCustomTaskPanes;
        private Dictionary<string, DocData> DocDatas;

        public OfficeAddinCustomTaskPane GetOfficeAddinCustomTaskPane()
        {
            string test = Application.ActiveDocument.Name;

            if (OfficeAddinCustomTaskPanes.ContainsKey(test))
            {

                return OfficeAddinCustomTaskPanes[test];
            }
            var pane = new OfficeAddinCustomTaskPane();

            OfficeAddinCustomTaskPanes.Add(test, pane);
            return pane;
        }

        public DocData GetDocData()
        {
            string test = Application.ActiveDocument.Name;

            return DocDatas[test];
        }


    
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            OfficeAddinCustomTaskPanes = new Dictionary<string, OfficeAddinCustomTaskPane>();
            DocDatas = new Dictionary<string, DocData>();
         
            Application.DocumentOpen += HookUpEvents;
            ((Word.ApplicationEvents4_Event)Application).NewDocument += HookUpEvents;
            Application.DocumentBeforeClose += ReleaseEvents;
        }

        private void ReleaseEvents(Word.Document doc, ref bool cancel)
        {
            ReleasePane();
            ReleaseDocData();

            string test = Application.ActiveDocument.Name;
            var pane = GetOfficeAddinCustomTaskPane();

            if (CustomTaskPanes.Contains(pane.CustomPane))
            {
                CustomTaskPanes.Remove(pane.CustomPane);
            }
        }

        private void ReleaseDocData()
        {
            string test = Application.ActiveDocument.Name;
            if (DocDatas.ContainsKey(test))
            {
                DocDatas.Remove(test);
            }
            
        }

        private void ReleasePane()
        {
            string test = Application.ActiveDocument.Name;

            if (OfficeAddinCustomTaskPanes.ContainsKey(test))
            {

                OfficeAddinCustomTaskPanes.Remove(test);
            }
        }


        private void HookUpEvents(Word.Document doc)
        {

            string test = Application.ActiveDocument.Name;

            var pane = GetOfficeAddinCustomTaskPane();

            var customPane = CustomTaskPanes.Add(pane, test);
            customPane.Visible = true;
            pane.CustomPane = customPane;
            
            Application.DocumentChange += DocumentSectionChange;
            pane.RefreshEvent += DocumentSectionChange;
            pane.SectionChangeEvent += SectionItemIndexChange;
            pane.RemoveSectionsEvent += RemoveSections;
            pane.SectionGroupingChangeEvent += GroupSectionsChange;
            pane.FindReplaceChangeEvent += FindReplaceChange;
            pane.ReplaceEvent += ReplaceText;

            pane.ClearHighlighting += ClearHighlighting;

        }

        private void ClearHighlighting(object sender, EventArgs e)
        {
            ClearFormatting();
        }

        public void ShowPane()

         {
            string test = Application.ActiveDocument.Name;
            var pane = GetOfficeAddinCustomTaskPane();

            if (CustomTaskPanes.Contains(pane.CustomPane))
            {
                pane.CustomPane.Visible = true;
            }
            else
            {
                HookUpEvents(null);
            }
        }

        private void ReplaceText(object sender, ReplaceEventArgs e)
        {
            var selectionRange = Application.ActiveDocument.Range();
            var findLocal = selectionRange.Find;

            var sectedItem = (FindReplaceSection) e.SelectedItem;


            if (sectedItem == null)
            {
                return;
            }

            findLocal.ClearFormatting();
            findLocal.Wrap = Word.WdFindWrap.wdFindContinue;
            findLocal.Text = sectedItem.Key;

            while (findLocal.Execute())  //If Found...
            {
                //change font and format of matched words
                selectionRange.Text  = e.ReplaceText;
            }

            GetOfficeAddinCustomTaskPane().FindReplaceList.Items.Remove(sectedItem);

        }

        private void FindReplaceChange(object sender, FindReplaceEventArgs e)
        {
            FindReplaceChange((FindReplaceSection)e.SelectedItem);
        }

        private void FindReplaceChange(FindReplaceSection selectedItem)
        {
            var selectionRangeReset = Application.ActiveDocument.Range();
            var findLocalReset = selectionRangeReset.Find;


            findLocalReset.ClearFormatting();
            findLocalReset.Format = true;
            findLocalReset.Wrap = Word.WdFindWrap.wdFindContinue;
            findLocalReset.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorOrange;

            while (findLocalReset.Execute()) 
            {
                selectionRangeReset.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorLavender;
            }

            var selectionRange = Application.ActiveDocument.Range();
            var  findLocal = selectionRange.Find;

            findLocal.ClearFormatting();
            findLocal.Text = selectedItem.Key;

            while (findLocal.Execute())  //If Found...
            {
                //change font and format of matched words
                selectionRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorOrange;
            }


        }

        private void GroupSectionsChange(object sender, SectionGroupingEventArgs sectionGroupingEventArgs)
        {
            this.GroupUpdate = true;
            DocumentSectionChange();
            this.GroupUpdate = false;
        }

        private bool GroupUpdate { get; set; }

        private void DocumentSectionChange(object sender, EventArgs e)
        {
            DocumentSectionChange();
        }

        private void RemoveSections(object sender, EventArgs e)
        {

            var selectionRange = Application.ActiveDocument.Range();
            var findLocal = selectionRange.Find;

            findLocal.ClearFormatting();
            findLocal.Format = true;
            findLocal.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorRed;
            findLocal.Execute(Replace: Word.WdReplace.wdReplaceAll);
            
            DocumentSectionChange();
        }

        private void SectionItemIndexChange(object sender, SectionChangeEventArgs sectionChangeEventArgs)
        {
             SectionItemIndexChange(sectionChangeEventArgs.SectionName, sectionChangeEventArgs.SectionSelected);
        }

        private void SectionItemIndexChange(string sectionName, bool sectionSelected)
        {
            var bookmarks = GetDocData().GetSections(sectionName);


            foreach (var bookmark in bookmarks)
            {
                try
                {
                     if (bookmark.Range != null)
                {
                    var range = bookmark.Range;

                    range.Font.Shading.BackgroundPatternColor = sectionSelected == false
                        ? Word.WdColor.wdColorRed
                        : Word.WdColor.wdColorGray25;
                }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    
                }
               
            }
        }

        private void DocumentSectionChange()
        {
            if (Application.Documents.Count >= 1)
            {
                //Application.ActiveDocument.TrackRevisions = true;
                //Application.ActiveWindow.View.ShowRevisionsAndComments = true;

                var pane = GetOfficeAddinCustomTaskPane();

                pane.ClearBookmarks();
                pane.ClearSearchReplace();
                ClearFormatting();

                var markers = CreateFindAndReplaceListItems();


                DocData docdata;
                if(!DocDatas.ContainsKey(Application.ActiveDocument.Name))
                {
                    docdata = new DocData(pane.GroupSections, Application.ActiveDocument.Bookmarks, markers);
                    DocDatas.Add(Application.ActiveDocument.Name, docdata);
                }
                else
                {
                    docdata = DocDatas[Application.ActiveDocument.Name];
                    docdata.Markers = markers;

                    docdata.SetBookmarks(pane.GroupSections,Application.ActiveDocument.Bookmarks);
                }

         

                docdata.UpdateSections_CheckedListBox(pane.SelectionsCheckList);
                docdata.UpdateFindAndReplace_ListBox(pane.FindReplaceList);

                pane.BookmarkCount = docdata.DocSectionsList.Count;
            }
        }

        private void ClearFormatting()
        {
            var selectionRangeReset = Application.ActiveDocument.Range();
            var findLocalReset = selectionRangeReset.Find;
            Word.Range rng = this.Application.ActiveDocument.Range(0, 0);
            rng.Select();

            findLocalReset.ClearFormatting();
            findLocalReset.Format = true;
            findLocalReset.Wrap = Word.WdFindWrap.wdFindContinue;


            var wdColors = new List<Word.WdColor>
            {
                Word.WdColor.wdColorLavender,
                Word.WdColor.wdColorRed,
                Word.WdColor.wdColorOrange,
                Word.WdColor.wdColorGray25,

            };

            foreach (var color in wdColors)
            {
                findLocalReset.Font.Shading.BackgroundPatternColor = color;

                while (findLocalReset.Execute())  //If Found...
                {
                    //change font and format of matched words
                    selectionRangeReset.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                }
            }

        }

        private List<Word.Range> CreateFindAndReplaceListItems()
        {
            Application.Selection.Find.ClearFormatting();
            Application.Selection.Find.MatchWildcards = true;
            Application.Selection.Find.Wrap = Word.WdFindWrap.wdFindContinue;


            var ranges = new List<Word.Range>();

            object findStr = @"\[*\]";

            while (Application.Selection.Find.Execute(ref findStr)) //If Found...
            {
                Application.Selection.Range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorLavender;
                ranges.Add(Application.Selection.Range);
            }

            return ranges;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
        
        
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
         
        }
       

        #endregion

       
    }
}
