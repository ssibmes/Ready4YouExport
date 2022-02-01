using System.Globalization;

namespace SolEolImportExport.Domain
{
    public class Company
    {
        public long Code { get; private set; }

        public long Hid { get; private set; }

        public string Description { get; private set; }

        public string LongDescription
        {
            get { return string.Format("{0} - {1}", Hid.ToString(CultureInfo.InvariantCulture).PadLeft(6, '0'), Description); }
        }

        public bool Current { get; set; }

        public Company(long code, long hid, string description, bool current)
        {
            Code = code;
            Hid = hid;
            Description = description;
            Current = current;
        }

        public override string ToString()
        {
            return string.Format("{0} - {1} - {2}", Code, Hid, Description);
        }
    }
}
