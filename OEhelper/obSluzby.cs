using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OEHelper
{
    class obSluzby
    {
        private string club;
        private string name;
        private int number;
        private string fee;
        private int totalFee;
        private string note;
        private string orderedBy;
        private string time;
        // #ID; Klub;  Typ; Nazev                                  ;Pocet;JednotkovaCena; CelkovaCena; Poznamka;ProKoho ;DatumACas        ;Termin#
        public obSluzby(string clubShort, string name, int number, string fee, int totalFee, string note, string orderedBy, string time)
        {
            this.club = clubShort;
            this.name = name;
            this.number = number;
            this.fee = fee;
            this.totalFee = totalFee;
            this.note = note;
            this.orderedBy = orderedBy;
            this.time = time;
        }

        public string Club
        {
            get { return this.club; }
        }

        public string Name
        {
            get { return this.name; }
        }

        public int Number
        {
            set { this.number = value; }
            get { return this.number; }
        }

        public string Fee
        {
            get { return this.fee; }
        }

        public int TotalFee
        {
            get { return this.totalFee; }
        }

        public string Note
        {
            get { return this.note; }
        }

        public string OrderedBy
        {
            get { return this.orderedBy; }
        }

        public string Time
        {
            get { return this.time; }
        }
    }
}
