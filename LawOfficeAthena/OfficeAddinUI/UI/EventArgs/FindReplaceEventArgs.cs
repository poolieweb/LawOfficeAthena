using System;

namespace OfficeAddinUI
{
    public class FindReplaceEventArgs : EventArgs
    {
        public object SelectedItem { get; set; }

        public FindReplaceEventArgs(object selectedItem)
        {
            SelectedItem = selectedItem;
        }
    }
}