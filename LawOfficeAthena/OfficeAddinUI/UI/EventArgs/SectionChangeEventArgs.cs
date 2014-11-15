using System;
using System.Windows.Forms;

namespace OfficeAddinUI
{
    public class SectionChangeEventArgs : EventArgs
    {
        public SectionChangeEventArgs(string sectionName, CheckState currentValue)
        {
            SectionName = sectionName;
            _currentValue = currentValue;
        }

        private readonly CheckState _currentValue;

        public string SectionName { get; set; }


        public bool SectionSelected
        {
            get { return _currentValue == CheckState.Checked; }
        }
    }
}