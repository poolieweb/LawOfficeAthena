using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeAddinUI
{
    public class DocData
    {


        public List<DocSection> DocSectionsList { get; set; }


        public DocData(bool groupSections, Word.Bookmarks bookmarks)
        {

           DocSectionsList = new List<DocSection>();
           foreach (Word.Bookmark bookmark in bookmarks)
           {
               DocSectionsList.Add(new DocSection(bookmark.Name, groupSections)); 
           }

        }

        public void UpdateSections_CheckedListBox(CheckedListBox selectionsCheckList)
        {
            foreach (DocSection docSection in DocSectionsList)
            {
                selectionsCheckList.Items.Add(docSection);
            }

        }



        public class DocSection
        {
            public bool GroupSections { get; set; }
            public string DetailName { get; set; }
            public string GroupName { get; set; }


            public override string ToString()
            {

                if (GroupSections && GroupName != null)
                {
                    return GroupName; 
                }

                if (DetailName.LastIndexOf('_') == -1)
                {
                    return DetailName;
                }

                return DetailName + "_" + GroupName;
            }

            public DocSection(string name, bool groupSections)
            {
                GroupSections = groupSections;


                if (name.LastIndexOf('_') == -1)
                {
                    DetailName = name;
                }
                else
                {
                    DetailName = name.Substring(0, name.LastIndexOf('_'));
                    GroupName = name.Substring(name.LastIndexOf('_')+1, name.Length -1 - name.LastIndexOf('_'));
                }
            }
        }

  
    }
}