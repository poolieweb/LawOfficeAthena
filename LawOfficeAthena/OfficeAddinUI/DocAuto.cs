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
            if (Application.Documents.Count >= 1)
            {
                _officeAddinCustomTaskPane.ClearBookmarks();

                DocData = new DocData(_officeAddinCustomTaskPane.GroupSections,Application.ActiveDocument.Bookmarks);

                DocData.UpdateSections_CheckedListBox(_officeAddinCustomTaskPane.SelectionsCheckList);
                _officeAddinCustomTaskPane.BookmarkCount = DocData.DocSectionsList.Count;
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
