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
        private OfficeAddinCustomTaskPane _officeAddinCustomTaskPane;
        public Microsoft.Office.Tools.CustomTaskPane _myCustomTaskPane;

        public DocData DocData { get; set; }
    
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _officeAddinCustomTaskPane = new OfficeAddinCustomTaskPane();

            Application.DocumentOpen += HookUpEvents;
            ((Word.ApplicationEvents4_Event)Application).NewDocument += HookUpEvents;
        }


        private void HookUpEvents(Word.Document doc)
        {
            _myCustomTaskPane = CustomTaskPanes.Add(_officeAddinCustomTaskPane, "Draft Assist");
            _myCustomTaskPane.Visible = true;

            Application.DocumentChange += DocumentSectionChange;
            _officeAddinCustomTaskPane.RefreshEvent += DocumentSectionChange;
            _officeAddinCustomTaskPane.SectionChangeEvent += SectionItemIndexChange;
            _officeAddinCustomTaskPane.RemoveSectionsEvent += RemoveSections;
            _officeAddinCustomTaskPane.SectionGroupingChangeEvent += GroupSectionsChange;
            _officeAddinCustomTaskPane.FindReplaceChangeEvent += FindReplaceChange;
            _officeAddinCustomTaskPane.ReplaceEvent += ReplaceText;
        }

        public void ShowPane()
        {
            _myCustomTaskPane = CustomTaskPanes.Add(_officeAddinCustomTaskPane, "Draft Assist");
            _myCustomTaskPane.Visible = true; 
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

            _officeAddinCustomTaskPane.FindReplaceList.Items.Remove(sectedItem);

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
            DocumentSectionChange();
        }

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
            var bookmarks = DocData.GetSections(sectionName);


            foreach (var bookmark in bookmarks)
            {
                var range = bookmark.Range;

                range.Font.Shading.BackgroundPatternColor = sectionSelected == false
                    ? Word.WdColor.wdColorRed
                    : Word.WdColor.wdColorGray25;
            }
        }

        private void DocumentSectionChange()
        {
            if (Application.Documents.Count >= 1)
            {
                //Application.ActiveDocument.TrackRevisions = true;
                //Application.ActiveWindow.View.ShowRevisionsAndComments = true;

                _officeAddinCustomTaskPane.ClearBookmarks();
                _officeAddinCustomTaskPane.ClearSearchReplace();
                ClearFormatting();

                var markers = CreateFindAndReplaceListItems();

               DocData = new DocData(_officeAddinCustomTaskPane.GroupSections, Application.ActiveDocument.Bookmarks, markers);

                DocData.UpdateSections_CheckedListBox(_officeAddinCustomTaskPane.SelectionsCheckList);
                DocData.UpdateFindAndReplace_ListBox(_officeAddinCustomTaskPane.FindReplaceList);
          
                _officeAddinCustomTaskPane.BookmarkCount = DocData.DocSectionsList.Count;
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
                Word.WdColor.wdColorOrange
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
            _myCustomTaskPane = null;
        }
       

        #endregion

       
    }
}
