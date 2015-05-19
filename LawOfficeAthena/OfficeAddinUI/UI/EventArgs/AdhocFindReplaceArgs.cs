using System;

namespace OfficeAddinUI
{
    public class AdhocFindReplaceArgs : EventArgs
    {
        public string TxtFind { get; set; }
        public string TxtReplace { get; set; }

        public AdhocFindReplaceArgs(string txtFind, string txtReplace)
        {
            TxtFind = txtFind;
            TxtReplace = txtReplace;
        }
    }
}