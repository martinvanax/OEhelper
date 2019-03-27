using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OEHelper
{
    class CastkaSlovne
    {
        public static string CisloSlovne(int cislo, bool capital = false)
        {
            string[] aJednotky = {"", "jedna", "dva", "tři", "čtyři", "pět", "šest", "sedm", 
                        "osm", "devět", "deset", "jedenáct", "dvanáct", "třináct", "čtrnáct",
                        "patnáct", "šestnáct", "sedmnáct", "osmnáct", "devatenáct"};
            string[] aDesitky = {"", "deset", "dvacet", "třicet", "čtyřicet", "padesát",
                        "šedesát", "sedmdesát", "osmdesát", "devadesát"};
            string[] aStovky = {"", "sto", "dvěstě", "třista", "čtyřista", "pětset",
                        "šestset", "sedmset", "osmset", "devětset"};
            string[] aRady = { "", "tisíc", "milionů", "miliard" };
            string[] aRady1 = { "", "tisíc", "milion", "miliarda" };
            string[] aRady234 = { "", "tisíce", "miliony", "miliardy" };

            int iStovky;
            int iDesitkyJednotky;

            string strRady;
            string strStovky;
            string strDesitkyJednotky;
            string strCisloText = string.Empty;

            //'skutečná délka čísla
            int iDelka = cislo.ToString().Length;
            //'počet trojic
            int iPocet3 = Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(iDelka) / 3));
            //'délka čísla po zaokrouhlení na trojice nahoru
            int iDelka3 = iPocet3 * 3;
            //'číslo formátované do trojic
            string strCislo3 = cislo.ToString().PadLeft(iDelka3, '0');



            //'pro všechny trojice
            for (int i = 1; i <= iPocet3; i++)
            {
                //'reset proměnných
                strStovky = string.Empty;
                strDesitkyJednotky = string.Empty;
                strRady = string.Empty;

                //'počet stovek
                //iStovky = (cislo % 1000) / 100; 
                iStovky = Convert.ToInt32(strCislo3.Substring(3 * (i - 1), 1));
                
                //'počet desítek a jednotek
                //iDesitkyJednotky = (cislo % 100); 
                iDesitkyJednotky = Convert.ToInt32(strCislo3.Substring((3 * (i - 1)) + 1, 2));

                /*'a) bez ošetření "jednosto"
                'strStovky = aStovky(iStovky + 1)

                'b) s ošetřením "jednosto"
                'If iStovky = 1 And i = 1 Then*/
                if (iStovky == 1)
                {
                    strStovky = "jedno" + aStovky[iStovky];
                }
                else
                {
                    strStovky = aStovky[iStovky];
                }

                //'rozlišení desítek a jednotek
                switch (iDesitkyJednotky)
                {
                    case 0:
                        if (iStovky == 0)
                        {
                            if (iPocet3 == 1)
                            {
                                strDesitkyJednotky = "nula";
                            }
                        }
                        else
                        {
                            //'text tisíců, milionů, ...
                            strRady = aRady[iPocet3 - i];
                        }
                        break;
                    case 1:
                        //'výjimka, "jeden" namísto "jedna" z pole
                        //'pro "jedentisíc", "jedenmilion", ...
                        if ((iStovky == 0) && (iPocet3 > 1) && (i != iPocet3))
                        {
                            //'text desítek a jednotek
                            strDesitkyJednotky = "jeden";
                        }
                        else
                        {
                            //'text desítek a jednotek
                            strDesitkyJednotky = aJednotky[iDesitkyJednotky];
                        }
                        //'text tisíců, milionů, ...
                        strRady = aRady1[iPocet3 - i];
                        break;

                    case 2:
                        //'výjimka, "dvě" namísto "dva" z pole
                        //'pro "dvě" (koruny, miliardy)
                        if (((iStovky == 0) && (iPocet3 == 1)) || ((iStovky == 0) && (iPocet3 == 4)))
                        {
                            //'text desítek a jednotek
                            strDesitkyJednotky = "dvě";
                        }
                        else
                        {
                            //'text desítek a jednotek
                            strDesitkyJednotky = aJednotky[iDesitkyJednotky];
                        }
                        //'text tisíců, milionů, ...
                        strRady = aRady234[iPocet3 - i];
                        break;
                    case 3:
                    case 4:
                        strDesitkyJednotky = aJednotky[iDesitkyJednotky];
                        //'text tisíců, milionů, ...
                        strRady = aRady234[iPocet3 - i];
                        break;
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                    case 9:
                    case 10:
                    case 11:
                    case 12:
                    case 13:
                    case 14:
                    case 15:
                    case 16:
                    case 17:
                    case 18:
                    case 19:
                        //'text desítek a jednotek
                        strDesitkyJednotky = aJednotky[iDesitkyJednotky];
                        //'text tisíců, milionů, ...
                        strRady = aRady[iPocet3 - i];
                        break;
                    default:
                        //'text desítek a jednotek
                        strDesitkyJednotky = aDesitky[(iDesitkyJednotky / 10)] + aJednotky[(iDesitkyJednotky % 10)];
                        //'text tisíců, milionů, ...
                        strRady = aRady[iPocet3 - i];
                        break;
                }

                strCisloText = strCisloText + strStovky + strDesitkyJednotky + strRady;

            }
            if(capital)
            {
                strCisloText = strCisloText.ToUpper();
            }
            return strCisloText;
        }
    }
}
