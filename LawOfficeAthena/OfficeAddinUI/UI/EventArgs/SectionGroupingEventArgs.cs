using System;

namespace OfficeAddinUI
{
    public class SectionGroupingEventArgs : EventArgs
    {
        public SectionGroupingEventArgs(bool groupSections)
        {
            GroupSections = groupSections;
        }


        public bool GroupSections { get; set; }
    }
}