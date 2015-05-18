using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace OfficeAddinUI
{
    public class DocSection
    {
        public List<Bookmark> Bookmarks { get; set; }
        private bool GroupSections { get; set; }

        public override string ToString()
        {

            string toString="";

            try
            {
                var firstOrDefault = Bookmarks.FirstOrDefault();
                if (firstOrDefault != null)
                {
                    var firstName = firstOrDefault.Name;

                    if (GroupSections && firstName.LastIndexOf('_') != -1)
                    {
                        return firstName.Substring(firstName.LastIndexOf('_') + 1,
                            firstName.Length - 1 - firstName.LastIndexOf('_'));
                    }

                    //if (!GroupSections && firstName.LastIndexOf('_') != -1)
                    //{
                    //    return firstName.Substring(0, firstName.LastIndexOf('_'));
                    //}

                    toString=  firstName;
                }
                else
                {
                       toString= "Unknown";
                }

               
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Debug.Print(ex.ToString());
                
            }
            return toString;
        }


        public DocSection(Bookmark bookmark, bool groupSections)
        {
            Bookmarks = new List<Bookmark>();
            Bookmarks.Add(bookmark);
            GroupSections = groupSections;
        }
    }
}