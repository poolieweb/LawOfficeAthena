﻿using System;
using System.Globalization;
using System.Windows.Forms;

namespace OfficeAddinUI
{
    public partial class OfficeAddinCustomTaskPane : UserControl
    {
        public event EventHandler RefreshEvent;
        public event EventHandler RemoveSectionsEvent;
        public event EventHandler<ReplaceEventArgs> ReplaceEvent;
        public event EventHandler<SectionChangeEventArgs> SectionChangeEvent;
        public event EventHandler<SectionGroupingEventArgs> SectionGroupingChangeEvent;
        public event EventHandler<FindReplaceEventArgs> FindReplaceChangeEvent;

        public OfficeAddinCustomTaskPane()
        {
            InitializeComponent();
        }

        public int BookmarkCount
        {
            set { label2.Text = value.ToString(CultureInfo.InvariantCulture); }
        }


        public bool GroupSections
        {
            get { return radioButton_groupSections.Checked; }
        }


        public CheckedListBox SelectionsCheckList { get; private set; }

        public ListBox FindReplaceList { get; private set; }


        private void button2_Click(object sender, EventArgs e)
        {
            // Make a temporary copy of the event to avoid possibility of 
            // a race condition if the last subscriber unsubscribes 
            // immediately after the null check and before the event is raised.
            var handler = RefreshEvent;

            // Event will be null if there are no subscribers 
            if (handler != null)
            {
                // Use the () operator to raise the event.
                handler(this, e);
            }
        }


        public void ClearBookmarks()
        {
            SelectionsCheckList.Items.Clear();
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Make a temporary copy of the event to avoid possibility of 
            // a race condition if the last subscriber unsubscribes 
            // immediately after the null check and before the event is raised.
            var handler = SectionChangeEvent;

            // Event will be null if there are no subscribers 
            if (handler != null)
            {
                var args = new SectionChangeEventArgs(SelectionsCheckList.Items[e.Index].ToString(), e.NewValue);

                // Use the () operator to raise the event.
                handler(this, args);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Make a temporary copy of the event to avoid possibility of 
            // a race condition if the last subscriber unsubscribes 
            // immediately after the null check and before the event is raised.
            var handler = RemoveSectionsEvent;

            // Event will be null if there are no subscribers 
            if (handler != null)
            {
                // Use the () operator to raise the event.
                handler(this, e);
            }
        }

        private void sectionGroup_CheckedChanged(object sender, EventArgs e)
        {
            // Make a temporary copy of the event to avoid possibility of 
            // a race condition if the last subscriber unsubscribes 
            // immediately after the null check and before the event is raised.
            var handler = SectionGroupingChangeEvent;

            // Event will be null if there are no subscribers 
            if (handler != null)
            {
                var args = new SectionGroupingEventArgs(radioButton_groupSections.Checked);

                // Use the () operator to raise the event.
                handler(this, args);
            }
        }


        // Define a class to hold custom event info 

        public void ClearSearchReplace()
        {
            FindReplaceList.Items.Clear();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Make a temporary copy of the event to avoid possibility of 
            // a race condition if the last subscriber unsubscribes 
            // immediately after the null check and before the event is raised.
            var handler = FindReplaceChangeEvent;

            var lbi = FindReplaceList.SelectedItem;

            // Event will be null if there are no subscribers 
            if (handler != null)
            {
                var args = new FindReplaceEventArgs(lbi);

                // Use the () operator to raise the event.
                handler(this, args);
            }
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            // Make a temporary copy of the event to avoid possibility of 
            // a race condition if the last subscriber unsubscribes 
            // immediately after the null check and before the event is raised.
            var handler = ReplaceEvent;
            var lbi = FindReplaceList.SelectedItem;
            var txt = textBox1.Text;

            // Event will be null if there are no subscribers 
            if (handler != null)
            {

                var args = new ReplaceEventArgs(lbi,txt);
                // Use the () operator to raise the event.
                handler(this, args);
            }
        }
    }
}