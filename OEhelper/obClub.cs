using System.Xml;

namespace OEHelper
{
    class obClub
    {
        public string Id { get; }
        public string ShortName { get; }
        public string ClubShort { get; }
        public string Name;
        public string Type { get; }
        public string OfficialName;
        public string OfficialStreet;
        public string OfficialZIP;
        public string OfficialCity;
        public string OfficialRegNum;
        public string ContactFirstName;
        public string ContactLastName;
        public string ContactStreet;
        public string ContactCity;
        public string ContactZIP;

        public obClub(string clubShort, string name)
        {
            ClubShort = clubShort;
            Name = name;
            OfficialName = string.Empty;
            OfficialStreet = string.Empty;
            OfficialZIP = string.Empty;
            OfficialCity = string.Empty;
            OfficialRegNum = string.Empty;
            ContactFirstName = string.Empty;
            ContactLastName = string.Empty;
            ContactStreet = string.Empty;
            ContactCity = string.Empty;
            ContactZIP = string.Empty;
        }

        public obClub(XmlNode xmlClub, XmlNamespaceManager nsmgr)
        {
            Type = xmlClub.Attributes["type"].Value;
            Id = InnerText(xmlClub, nsmgr, "c:Id", string.Empty);
            Name = InnerText(xmlClub, nsmgr, "c:Name", string.Empty);
            ClubShort = InnerText(xmlClub, nsmgr, "c:Extensions/ex:Abbreviation", string.Empty);
        }

        private string InnerText(XmlNode xml, XmlNamespaceManager nsmgr, string xpath, string defaultString)
        {
            string ret = defaultString;
            var n = xml.SelectSingleNode(xpath, nsmgr);
            if (n != null)
            {
                ret = n.InnerText;
            }
            return ret;
        }
    }
}
