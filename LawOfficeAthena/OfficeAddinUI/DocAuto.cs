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
        }

        private void ThisAddIn_SectionChange(object sender, OfficeAddinCustomTaskPane.SectionChangeEventArgs sectionChangeEventArgs)
        {
            var section = this.Application.ActiveDocument.Bookmarks[sectionChangeEventArgs.SectionName];
            var range = section.Range;

            if (sectionChangeEventArgs.SectionSelected == false)
            {
                range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorRed;
            }
            else
            {
                range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
            }



            //ThisAddIn_DocumentChange();

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
                    _officeAddinCustomTaskPane.AddBookmark(bookmark.Name);
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
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        
       

        #endregion
    }
}
