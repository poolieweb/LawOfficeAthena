using System;

namespace OfficeAddinUI
{
    public class ReplaceEventArgs : EventArgs
    {
        public object SelectedItem { get; set; }
        public string ReplaceText { get; set; }

        public ReplaceEventArgs(object selectedItem, string replaceText)
        {
            SelectedItem = selectedItem;
            ReplaceText = replaceText;
        }
    }
}