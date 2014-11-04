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
        private Microsoft.Office.Tools.CustomTaskPane _myCustomTaskPane;

        public DocData DocData { get; set; }
    
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _officeAddinCustomTaskPane = new OfficeAddinCustomTaskPane();
            _myCustomTaskPane = this.CustomTaskPanes.Add(_officeAddinCustomTaskPane, "Draft Assist");
            _myCustomTaskPane.Visible = true;

            this.Application.DocumentChange += ThisAddIn_DocumentChange;
            _officeAddinCustomTaskPane.RefreshEvent += ThisAddIn_DocumentChange;
            _officeAddinCustomTaskPane.SectionChangeEvent += ThisAddIn_SectionChange;
            _officeAddinCustomTaskPane.RemoveSectionsEvent += ThisAddInRemoveSectionsEvent;
            _officeAddinCustomTaskPane.SectionGroupingChangeEvent += ThisAddIn_GroupSectionsSections;
        }

        private void ThisAddIn_GroupSectionsSections(object sender, OfficeAddinCustomTaskPane.SectionGroupingEventArgs sectionGroupingEventArgs)
        {
            ThisAddIn_DocumentChange();
        }

        private void ThisAddInRemoveSectionsEvent(object sender, EventArgs e)
        {

            var selectionRange = Application.ActiveDocument.Range();
            var findLocal = selectionRange.Find;

            findLocal.ClearFormatting();
            findLocal.Format = true;
            findLocal.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorRed;
            findLocal.Execute(Replace: Word.WdReplace.wdReplaceAll);
            
            ThisAddIn_DocumentChange();
        }

        private void ThisAddIn_SectionChange(object sender, OfficeAddinCustomTaskPane.SectionChangeEventArgs sectionChangeEventArgs)
        {
            var section = this.Application.ActiveDocument.Bookmarks[sectionChangeEventArgs.SectionName];

            var range = section.Range;

            range.Font.Shading.BackgroundPatternColor = sectionChangeEventArgs.SectionSelected == false 
                ? Word.WdColor.wdColorRed : Word.WdColor.wdColorGray25;

        }

        private void ThisAddIn_DocumentChange(object sender, EventArgs e)
        {
            ThisAddIn_DocumentChange();
        }

        private void ThisAddIn_DocumentChange()
        {

            if (this.Application.Documents.Count >= 1)
            {
                var bookmarkCount = this.Application.ActiveDocument.Bookmarks.Count;
                           
                DocData = new DocData { BookmarkCount = bookmarkCount };

                _officeAddinCustomTaskPane.BookmarkCount = bookmarkCount;
                _officeAddinCustomTaskPane.ClearBookmarks();

                foreach ( Word.Bookmark bookmark in Application.ActiveDocument.Bookmarks)
                {

                    if (_officeAddinCustomTaskPane.GroupSections)
                    {

                        if (bookmark.Name.LastIndexOf('_') >= 1)
                        {
                            var sectionName = bookmark.Name.Substring(0, bookmark.Name.LastIndexOf('_'));
                            _officeAddinCustomTaskPane.AddBookmark(sectionName);
                        }   else
                        {
                            _officeAddinCustomTaskPane.AddBookmark(bookmark.Name);
                        }

                   
                    }
                    else
                    {
                        _officeAddinCustomTaskPane.AddBookmark(bookmark.Name);
                    }

                }
            }
 
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
        
        
       

        #endregion
    }
}
