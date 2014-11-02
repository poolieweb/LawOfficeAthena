﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeAddinUI
{
    public partial class OfficeAddinCustomTaskPane : UserControl
    {

        public event EventHandler RefreshEvent;
        public event EventHandler<SectionChangeEventArgs> SectionChangeEvent;


        // Define a class to hold custom event info 
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


        public OfficeAddinCustomTaskPane()
        {
            InitializeComponent();
        }

        public int BookmarkCount
        {
            set { label2.Text = value.ToString(CultureInfo.InvariantCulture); }
        }

        public void AddBookmark(string bookmark)
        {
            checkedListBox1.Items.Add(bookmark,true);
        }

        private void button2_Click(object sender, EventArgs e)
        {

            // Make a temporary copy of the event to avoid possibility of 
            // a race condition if the last subscriber unsubscribes 
            // immediately after the null check and before the event is raised.
            EventHandler handler = RefreshEvent;

            // Event will be null if there are no subscribers 
            if (handler != null)
            {
                // Use the () operator to raise the event.
                handler(this, e);
            }
        }


        public void ClearBookmarks()
        {
            checkedListBox1.Items.Clear();
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Make a temporary copy of the event to avoid possibility of 
            // a race condition if the last subscriber unsubscribes 
            // immediately after the null check and before the event is raised.
            EventHandler<SectionChangeEventArgs> handler = SectionChangeEvent;

            // Event will be null if there are no subscribers 
            if (handler != null)
            {

                var args = new SectionChangeEventArgs(checkedListBox1.Items[e.Index].ToString(), e.NewValue);

                // Use the () operator to raise the event.
                handler(this, args);
            }
        }
        
    }
}
