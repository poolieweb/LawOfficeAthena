using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeAddinUI
{
    public class DocData
    {
        public List<Word.Range> Markers { get; set; }

        public List<DocSection> DocSectionsList { get; set; }


        public DocData(bool groupSections, Word.Bookmarks bookmarks, List<Word.Range> markers)
        {
            Markers = markers;


            DocSectionsList = new List<DocSection>();

           if (groupSections)
           {
                foreach (Word.Bookmark bookmark in bookmarks)
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
                foreach (Word.Bookmark bookmark in bookmarks)
                {
                    DocSectionsList.Add(new DocSection(bookmark, false));
                } 
            }



          

        }

        public void UpdateSections_CheckedListBox(CheckedListBox selectionsCheckList)
        {
            foreach (DocSection docSection in DocSectionsList)
            {
                selectionsCheckList.Items.Add(docSection,true);
            }
        }



        public class DocSection
        {
            public List<Word.Bookmark> Bookmarks { get; set; }
            private bool GroupSections { get; set; }

            public override string ToString()
            {
             
                var firstName = Bookmarks.FirstOrDefault().Name;

                if (GroupSections && firstName.LastIndexOf('_') != -1)
                {
                    return firstName.Substring(firstName.LastIndexOf('_') + 1, firstName.Length - 1 - firstName.LastIndexOf('_'));
                }

                if (!GroupSections && firstName.LastIndexOf('_') != -1)
                {
                    return firstName.Substring(0, firstName.LastIndexOf('_'));
                }

                return firstName;
            }


            public DocSection(Word.Bookmark bookmark, bool groupSections)
            {
                Bookmarks = new List<Word.Bookmark>();
                Bookmarks.Add(bookmark);
                GroupSections = groupSections;
            }

        }

        public List<Word.Bookmark> GetSections(string sectionName)
        {

            var selctionList = new List<Word.Bookmark>();

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