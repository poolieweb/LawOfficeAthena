using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace OfficeAddinUI
{
    public class DocData
    {
        public List<Range> Markers { get; set; }

        public List<DocSection> DocSectionsList { get; set; }

        public DocData(bool groupSections, Bookmarks bookmarks, List<Range> markers)
        {
            Markers = markers;

            SetBookmarks(groupSections, bookmarks);
        }

        public void SetBookmarks(bool groupSections, Bookmarks bookmarks)
        {
            DocSectionsList = new List<DocSection>();

            if (groupSections)
            {
                foreach (Bookmark bookmark in bookmarks)
                {
                    var docSection = new DocSection(bookmark, true);

                    if (DocSectionsList.All(d => d.ToString() != docSection.ToString()))
                    {
                        DocSectionsList.Add(docSection);
                    }
                    else
                    {
                        docSection = DocSectionsList.FirstOrDefault(d => d.ToString() == docSection.ToString());
                        if (docSection != null) docSection.Bookmarks.Add(bookmark);
                    }
                }
            }
            else
            {
                foreach (Bookmark bookmark in bookmarks)
                {
                    DocSectionsList.Add(new DocSection(bookmark, false));
                }
            }
        }

        public void UpdateSections_CheckedListBox(CheckedListBox selectionsCheckList)
        {
            foreach (var docSection in DocSectionsList)
            {
                selectionsCheckList.Items.Add(docSection, true);
            }
        }

        public void UpdateFindAndReplace_ListBox(ListBox findReplaceList)
        {

            findReplaceList.Items.Clear();

            var results = from m in Markers
                where m.Text != null
                group m by m.Text
                into grp
                select new FindReplaceSection(grp.Key, grp.Count());


            foreach (var marker in results)
            {
                findReplaceList.Items.Add(marker);
            }
        }

        public List<Bookmark> GetSections(string sectionName)
        {
            var selctionList = new List<Bookmark>();

            foreach (var docSection in DocSectionsList)
            {
                if (docSection.ToString() == sectionName)
                {
                    selctionList.AddRange(docSection.Bookmarks);
                }
            }

            return selctionList;
        }
    }
}