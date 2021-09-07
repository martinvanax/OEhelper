using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OEHelper
{
    class EnvelopeLabel
    {
        public bool Select { get; set; }
        public string ClubId { get; set; }
        public string Club { get; set; }
        public string ClubAbbr { get; set; }
        public string Entry { get; set; }
        public string SI { get; set; }
        public string Service { get; set; }
        public string Total { get; set; }
        public string Paid { get; set; }
        public string ToBePaid { get; set; }
       
        public EnvelopeLabel(string clubId, string clubName, string clubAbbr, string entry, string SI, string service, string total, string paid, string toBePaid)
        {
            this.Select = false;
            this.ClubId = clubId;
            this.Club = clubName;
            this.ClubAbbr = clubAbbr;
            this.Entry = entry;
            this.SI = SI;
            this.Service = service;
            this.Total = total;
            this.Paid = paid;
            this.ToBePaid = toBePaid;
        }

        
    }

    class CeremonyLabel
    {
        public bool Select { get; set; }
        public string ClassName { get; set; }
        public string ClassCEC { get; set; }
        public string Place { get; set; }
        
        public CeremonyLabel(string className, string classCEC, string place)
        {
            this.Select = false;
            this.ClassName = className;
            this.ClassCEC = classCEC;
            this.Place = place;
        }
    }
}
