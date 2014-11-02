using Microsoft.Office.Interop.Word;

namespace OfficeAddinUI
{
    public class DocData
    {
        public Range oldrange;
        public int BookmarkCount { get; set; }

        public void TrackChange(Bookmark section)
        {
             oldrange = section.Range.Duplicate;
        }
    }
}