namespace OfficeAddinUI
{
    public class FindReplaceSection

    {
        public FindReplaceSection(string key, int count)
        {
            Key = key;
            Count = count;
        }

        public string Key { get; set; }
        public int Count { get; set; }

        public override string ToString()
        {
            return "(" + Count + ") : " + Key;
        }
    }
}