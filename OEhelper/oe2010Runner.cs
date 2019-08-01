using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OEHelper
{
    class oe2010Runner
    {
        public string ClubShort { get; }
        public string Club { get; }
        public string Line { get; }
        public int Fee { get; }
        public int SIRent { get; }
        public string Category { get; }
        public string Name { get; }
        public string RegCode { get; }

        public oe2010Runner(string line, string oeType, bool fromOE)
        {
            /* TODO tady to chce dodelat aby volba sloubcu byla nastavitelna, pro jednodenni zavod to asi bude jinak */
            /*                                                                                                       10                                                               20                                     30                                                                                         40                                                                                          50                                                                             60                                                                            70                                                                                             80                                                                            */
            string[] parts = line.Split(';');
            Line = line;
            if ("OE0001" == oeType)
            {
                /* 0 OE0001;StČí;Xčís.;Číslo čipu;Id databáze;Příjmení;Jméno (křest.);RN;S;Blok;
                 * 10 ms;Start;Cíl;Čas;Klasifikace;Kredit -;Penalizace +;Komentář;Č. oddílu;Název oddílu;
                 * 20 Město;Stát;Místo;Zkratka;Kat. č.;Krátký;Dlouhý;Kat. č.;Kat. (krátká);Kat. (dlouhá);
                 * 30 Ranking;Rankingové body;Num1;Num2;Num3;Text1;Text2;Text3;Příjmení;Jméno;
                 * 40 Ulice;2.řádek;PSČ;Město;Tel;Mobil;Fax;E-mail;Půjčeno;Vklady;
                 * 50 Placeno;Družstvo;Číslo tratě;Trať;km;m;Kontroly tratě;
                 0- ;599;;9852;"TYN7200";"Adamovský";"Michal";1972;M;;
                 1- 0;31:00;;;0;;;"";182;"";
                 2- "OOB SK Týniště nad Orlicí";"CZE";"";"TYN";101594;"H45";"H45";101594;"H45";"H45";
                 3- ;;;;;"";"";"";;;
                 4- ;;;;;;;;0;"120,00";
                 5- 0;;9015;H45;"3,0";190;17; 

                 0- ; ; ; 8546; LLI5550; Korejčíková; Dáďa; 1955; F; ;
                 1- ; ; ; ; ; ; ; ; 89; LLI;
                 2- LOK - o Liberec; ; ; ; 107055; D35; D35; ; ; ;
                 3- ; ; ; 0; ; LLI5550; C; ; ; ; ; ; ; ; ; ; ; ; ; 120; ; ; ; ; ; ;
                */
                string tempName = parts[5].Replace("\"", string.Empty);
                if (("Vakant" == tempName) || ("Vacant" == tempName))
                {
                    Name = "Vakant";
                }
                else
                {
                    /*přihlášky z ORISU */
                    if (!fromOE)
                    {
                        ClubShort = parts[19].Replace("\"", string.Empty);
                        Club = parts[20].Replace("\"", string.Empty);
                        Fee = int.Parse(parts[49].Replace("\"", string.Empty), System.Globalization.NumberStyles.Any);
                        if (parts[48] == "X")
                        {
                            SIRent = Convert.ToInt32(Properties.Settings.Default.rentSIFee);
                        }
                        else
                        {
                            if (parts[33] != string.Empty)
                            {
                                if (parts[33] != "0")
                                {
                                    SIRent = Convert.ToInt32(Properties.Settings.Default.rentSIFee);
                                }
                                else
                                {
                                    SIRent = 0;
                                }
                            }
                            else
                            {
                                SIRent = 0;
                            }
                        }

                        Category = parts[25].Replace("\"", string.Empty);
                        Name = parts[5].Replace("\"", string.Empty) + " " + parts[6].Replace("\"", string.Empty);
                        RegCode = parts[4].Replace("\"", string.Empty);
                    }
                    else
                    {    /* přihlášky generované z OE */
                        ClubShort = parts[23].Replace("\"", string.Empty);
                        Club = parts[20].Replace("\"", string.Empty);
                        Fee = int.Parse(parts[49].Replace("\"", string.Empty), System.Globalization.NumberStyles.Any);
                        SIRent = (parts[48] == "X") ? Convert.ToInt32(Properties.Settings.Default.rentSIFee) : 0;
                        Category = parts[25].Replace("\"", string.Empty);
                        Name = parts[5].Replace("\"", string.Empty) + " " + parts[6].Replace("\"", string.Empty);
                        RegCode = parts[4].Replace("\"", string.Empty);
                    }
                }
            }
            else if ("OE0002" == oeType)
            {
                /* 0 OE0002;StČí;Xčís.;Číslo čipu1;Číslo čipu2;Číslo čipu3;Číslo čipu4;Číslo čipu5;Číslo čipu6;Id databáze;
                 * 10 Příjmení;Jméno (křest.);RN;S;Blok1;Blok2;Blok3;Blok4;Blok5;Blok6;
                 * 20 P1;P2;P3;P4;P5;P6;ms1;Start1;Cíl1;Čas1;
                 * 30 Klasifikace1;Kredit -1;Penalizace +1;Komentář1;ms2;Start2;Cíl2;Čas2;Klasifikace2;Kredit -2;
                 * 40 Penalizace +2;Komentář2;ms3;Start3;Cíl3;Čas3;Klasifikace3;Kredit -3;Penalizace +3;Komentář3;
                 * 50 ms4;Start4;Cíl4;Čas4;Klasifikace4;Kredit -4;Penalizace +4;Komentář4;ms5;Start5;
                 * 60 Cíl5;Čas5;Klasifikace5;Kredit -5;Penalizace +5;Komentář5;ms6;Start6;Cíl6;Čas6;
                 * 70 Klasifikace6;Kredit -6;Penalizace +6;Komentář6;Č. oddílu;Název oddílu;Město;Stát;Místo;Zkratka;
                 * 80 Kat. č.;Krátký;Dlouhý;Kat. č.;Kat. (krátká);Kat. (dlouhá);Ranking;Rankingové body;Num1;Num2;
                 * 90 Num3;Text1;Text2;Text3;Příjmení;Jméno;Ulice;2.řádek;PSČ;Město;
                 * 100 Tel;Mobil;Fax;E-mail;Půjčeno;Vklady;Placeno;Družstvo;
                ;732;;9620;9620;9620;9620;9620;9620;"SNA5052";"Rufferová";"Iva";1950;F;;;;;;;X;X;X;X;X;X;0;;;;0;;;"";0;;;;0;;;"";0;110;;;0;;;"";0;;;;0;;;"";0;;;;0;;;"";0;;;;0;;;"";156;"";"TJ START Náchod";"CZE";"";"SNA";85311;"D65";"D65";85310;"D60";"D60";;;;;;"";"";"";;;;;;;;;;;0;"450,00";0;;*/
                string tempName = parts[10].Replace("\"", string.Empty);
                if (("Vakant" == tempName) || ("Vacant" == tempName))
                {
                    Name = "Vakant";
                }
                else
                {
                    ClubShort = parts[79].Replace("\"", string.Empty);
                    Club = parts[76].Replace("\"", string.Empty);
                    Fee = int.Parse(parts[105].Replace("\"", string.Empty), System.Globalization.NumberStyles.Any);
                    SIRent = (parts[104] == "X") ? Convert.ToInt32(Properties.Settings.Default.rentSIFee) : 0;
                    Category = parts[81].Replace("\"", string.Empty); /* kat kratky */
                    Name = parts[10].Replace("\"", string.Empty) + " " + parts[11].Replace("\"", string.Empty);
                    RegCode = parts[9].Replace("\"", string.Empty);
                }
            }
        }
        /* TODO import IOF3.0 */
    }
}
