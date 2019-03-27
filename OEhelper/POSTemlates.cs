using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace OEHelper
{
    class POSTemlates
    {
        const int TOP_MARGIN_800 = 800;
        int TOP_MARGIN;
        int top_margin;
        int bot_margin = 25;
        int left_margin = 0;
        Document document;
        PdfWriter writer;
        PdfContentByte cb;

        public POSTemlates(Document document, PdfWriter writer, int top_margin_from_bot = TOP_MARGIN_800)
        {
            this.document = document;
            this.writer = writer;
            this.TOP_MARGIN = top_margin_from_bot;
            this.top_margin = this.TOP_MARGIN;
        }

        public void printBillOverview(obClub club, List<oe2010Runner> clubRunners, List<obSluzby> clubSluzby, List<obPlatby> clubPlatby, out int feeEntry, out int feeSI, out int feeService, out int paid)
        {
            feeEntry = 0;
            feeSI = 0;
            feeService = 0;
            paid = 0;
            int runnersCount = 0;
            int serviceCount = 0;

            try
            {
                left_margin = 40;
                
                BaseFont f_cb = BaseFont.CreateFont("c:\\windows\\fonts\\calibrib.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                document.NewPage();

                cb = writer.DirectContent;
                // First we must activate writing
                cb.BeginText();
                // First we write out the header information
                // Start with the invoice type header
                writeText(cb, "Vyúčtování vkladů", left_margin, 800, f_cb, 14);
                top_margin = 785;
                writeText(cb, "Závod:", left_margin, top_margin, f_cb, 10);
                writeText(cb, Properties.Settings.Default.eventName + " - " + Properties.Settings.Default.organizerIdentificationAbbr, left_margin + 50, top_margin, f_cn, 10);
                writeText(cb, "Pořadatel:", left_margin, top_margin - 24, f_cb, 10);
                writeText(cb, Properties.Settings.Default.organizerIdentificationL1, left_margin, top_margin - 36, f_cn, 10);
                writeText(cb, Properties.Settings.Default.organizerIdentificationL2, left_margin, top_margin - 48, f_cn, 10);
                writeText(cb, Properties.Settings.Default.organizerIdentificationL3, left_margin, top_margin - 60, f_cn, 10);
                    
                writeText(cb, "Plátce:", left_margin + 270 , top_margin - 24, f_cb, 10);
                if (club.OfficialName != string.Empty)
                {
                    writeText(cb, club.OfficialName, left_margin + 270, top_margin - 36, f_cn, 10);
                    writeText(cb, club.OfficialStreet + ", " + club.OfficialCity + ", " + club.OfficialZIP, left_margin + 270, top_margin - 48, f_cn, 10);
                    writeText(cb, "IČ: " + club.OfficialRegNum, left_margin + 270, top_margin - 60, f_cn, 10);
                }
                else
                {
                    writeText(cb, club.Name, left_margin + 270, top_margin - 36, f_cn, 10);
                    writeText(cb, club.ContactFirstName + " " + club.ContactLastName, left_margin + 270, top_margin - 48, f_cn, 10);
                    writeText(cb, club.ContactStreet + " " + club.ContactCity + " " + club.ContactZIP, left_margin + 270, top_margin - 60, f_cn, 10);
                }
               
                writeText(cb, club.ClubShort, left_margin + 460, top_margin, f_cb, 32);

                moveTop(84); 
                writeText(cb, "Přihlášky:", left_margin, top_margin, f_cb, 10);
                // Line headers
                moveTop(12);
                writeText(cb, "KATEGORIE", left_margin, top_margin, f_cb, 10);
                writeText(cb, "REG.ČÍSLO", left_margin + 70, top_margin, f_cb, 10);
                writeText(cb, "JMÉNO", left_margin + 140, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, "VKLAD", left_margin + 410, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, "PŮJČENÍ SI", left_margin + 480, top_margin, f_cb, 10);

                foreach (oe2010Runner runner in clubRunners)
                {
                    moveTop(12);
                    writeText(cb, runner.Category, left_margin, top_margin, f_cn, 10);
                    writeText(cb, runner.RegCode, left_margin + 70, top_margin, f_cn, 10);
                    writeText(cb, runner.Name, left_margin + 140, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, runner.Fee + " Kč", left_margin + 410, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, runner.SIRent + " Kč", left_margin + 480, top_margin, f_cn, 10);
                    feeEntry += runner.Fee;
                    feeSI += runner.SIRent;
                    runnersCount++;
                }
                
                //tab total
                moveTop(12);
                writeText(cb, "celkem", left_margin + 70, top_margin, f_cb, 10);
                writeText(cb, runnersCount + " ks", left_margin + 140, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, feeEntry + " Kč", left_margin + 410, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, feeSI + " Kč", left_margin + 480, top_margin, f_cb, 10);
                
                moveTop(12);
                moveTop(12);
                writeText(cb, "Doplňkové služby:", left_margin, top_margin, f_cb, 10);
                // Line headers
                moveTop(12);
                writeText(cb, "POPIS", left_margin, top_margin, f_cb, 10);
                writeText(cb, "CENA", left_margin + 180, top_margin, f_cb, 10);
                writeText(cb, "MNOŽSTVÍ", left_margin + 220, top_margin, f_cb, 10);
                writeText(cb, "CELKOVÁ CENA", left_margin + 290, top_margin, f_cb, 10);
                writeText(cb, "VYTVOŘIL", left_margin + 365, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, "VYTVOŘENO", left_margin + 480, top_margin, f_cb, 10);

                foreach (obSluzby service in clubSluzby)
                {
                    moveTop(12);
                    writeText(cb, service.Name, left_margin, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, service.Fee + " Kč", left_margin + 205, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, service.Number.ToString(), left_margin + 262, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, service.TotalFee + " Kč", left_margin + 350, top_margin, f_cn, 10);
                    writeText(cb, service.OrderedBy, left_margin + 365, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, service.Time, left_margin + 480, top_margin, f_cn, 10);
                    feeService += service.TotalFee;
                    serviceCount++;
                }

                //tab total
                moveTop(12);
                writeText(cb, "celkem", left_margin, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, serviceCount.ToString(), left_margin + 262, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, feeService + " Kč", left_margin + 350, top_margin, f_cb, 10);
                
                moveTop(12); 
                moveTop(12); 
                writeText(cb, "Přehled plateb:", left_margin, top_margin, f_cb, 10);
                // Line headers
                moveTop(12);
                writeText(cb, "DATUM PLATBY", left_margin, top_margin, f_cb, 10);
                writeText(cb, "ČÁSTKA", left_margin + 120, top_margin, f_cb, 10);
                writeText(cb, "REG.ČÍSLO", left_margin + 190, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, "POZNÁMKA", left_margin + 480, top_margin, f_cb, 10);

                foreach (obPlatby platba in clubPlatby)
                {
                    moveTop(12);
                    writeText(cb, platba.Time, left_margin, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, platba.TotalFee + " Kč", left_margin + 152, top_margin, f_cn, 10);
                    writeText(cb, string.Empty, left_margin + 190, top_margin, f_cn, 10);
                    writeText(cb, platba.Note, left_margin + 260, top_margin, f_cn, 10);
                    paid += platba.TotalFee;
                }

                //tab total
                moveTop(12);
                writeText(cb, "celkem", left_margin, top_margin, f_cb, 10);
                writeTextAlignedRight(cb, paid + " Kč", left_margin + 152, top_margin, f_cb, 10);

                int feeTotal = feeEntry + feeSI + feeService;
                moveTop(12);
                moveTop(12);
                writeText(cb, "Poplatky za přihlášku celkem", left_margin, top_margin, f_cb, 11);
                writeTextAlignedRight(cb, feeTotal + " Kč", left_margin + 260, top_margin, f_cb, 11);
                moveTop(13); 
                writeText(cb, "Uhrazeno předem převodem z účtu", left_margin, top_margin, f_cb, 11);
                writeTextAlignedRight(cb, paid + " Kč", left_margin + 260, top_margin, f_cb, 11);
                moveTop(13); 
                writeText(cb, "Rozdíl - přijato/vráceno v hotovosti", left_margin, top_margin, f_cb, 11);
                writeTextAlignedRight(cb, (feeTotal - paid) + " Kč", left_margin + 260, top_margin, f_cb, 11);

                moveTop(60);
                writeText(cb, "..........................................................................................", left_margin, top_margin, f_cn, 10);
                moveTop(13);
                writeText(cb, "Datum, podpis příjemce", left_margin, top_margin, f_cn, 10);

                cb.EndText();                                
            }
            catch (Exception rror)
            {
                System.Windows.Forms.MessageBox.Show(rror.Message);
            }
        }

        private void moveTop(int move)
        {
            if ((top_margin - move) < bot_margin)
            {
                cb.EndText();
                document.NewPage();
                cb.BeginText();
                top_margin = TOP_MARGIN;
            }
            else
            {
                top_margin -= move;
            }
        }

        // This is the method writing text to the content byte
        private void writeText(PdfContentByte cb, string Text, int X, int Y, BaseFont font, int Size)
        {
            cb.SetFontAndSize(font, Size);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Text, X, Y, 0);
        }

        private void writeTextAlignedRight(PdfContentByte cb, string Text, int X, int Y, BaseFont font, int Size)
        {
            cb.SetFontAndSize(font, Size);
            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, Text, X, Y, 0);
        }

        public void printCashBill(obClub club, int fee)
        {
            try
            {
                bool income = true;
                if (fee < 0)
                {
                    income = false;
                    fee = fee * (-1);
                }

                left_margin = 20;

                BaseFont f_cb = BaseFont.CreateFont("c:\\windows\\fonts\\calibrib.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                document.NewPage();

                cb = writer.DirectContent;
                // First we must activate writing
                cb.BeginText();
                // First we write out the header information
                // Start with the invoice type header

                if (income)
                {
                    writeText(cb, "Příjmový pokladní doklad", left_margin, 260, f_cb, 12);
                }
                else
                {
                    writeText(cb, "Výdajový pokladní doklad", left_margin, 260, f_cb, 12);
                }
                writeText(cb, "č.:........................... ze dne " + Properties.Settings.Default.eventDate, left_margin + 138, 260, f_cn, 10);
                top_margin = 240;
                writeText(cb, "Organizace:", left_margin, 240, f_cb, 10);
                writeText(cb, Properties.Settings.Default.organizerIdentificationL1, left_margin, top_margin - 12, f_cn, 10);
                writeText(cb, Properties.Settings.Default.organizerIdentificationL2, left_margin, top_margin - 24, f_cn, 10);
                writeText(cb, Properties.Settings.Default.organizerIdentificationL3, left_margin, top_margin - 36, f_cn, 10);

                if (income)
                {
                    writeText(cb, "Přijato od:", left_margin, 140, f_cb, 10);
                }
                else
                {
                    writeText(cb, "Vyplaceno komu:", left_margin, 140, f_cb, 10);
                }

                if (club.OfficialName != string.Empty)
                {
                    writeText(cb, club.OfficialName, left_margin + 75, 140, f_cn, 10);
                    writeText(cb, club.OfficialStreet + ", " + club.OfficialCity + ", " + club.OfficialZIP, left_margin + 75, 128, f_cn, 10);
                    writeText(cb, "IČ: " + club.OfficialRegNum, left_margin + 75, 116, f_cn, 10);
                }
                else
                {
                    writeText(cb, club.Name, left_margin + 75, 140, f_cn, 10);
                    writeText(cb, club.ContactFirstName + " " + club.ContactLastName, left_margin + 75, 128, f_cn, 10);
                    writeText(cb, club.ContactStreet + " " + club.ContactCity + " " + club.ContactZIP, left_margin + 75, 116, f_cn, 10);
                }

                writeText(cb, "Účel platby:", left_margin, 96, f_cb, 10);
                writeText(cb, "vklady " + Properties.Settings.Default.eventName, left_margin + 75, 96, f_cn, 10);

                writeText(cb, "Celkem: " + fee + " Kč", left_margin + 230, 180, f_cb, 14);
                writeText(cb, "Slovy:", left_margin, 160, f_cb, 10);
                writeText(cb, "---" + CastkaSlovne.CisloSlovne(fee) + "---", left_margin + 75, 160, f_cn, 10);

                writeText(cb, club.ClubShort, left_margin + 340, 260, f_cb, 16);
                if (income)
                {
                    writeText(cb, "Příjal:", left_margin + 180, 50, f_cb, 10);
                }
                else
                {
                    writeText(cb, "Vydal:", left_margin + 180, 50, f_cb, 10);
                }
                
                cb.EndText();
            }
            catch (Exception rror)
            {
                System.Windows.Forms.MessageBox.Show(rror.Message);
            }
        }

        public void printServiceOverviewEmpty()
        {
            try
            {
                BaseFont f_cb = BaseFont.CreateFont("c:\\windows\\fonts\\calibrib.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                left_margin = 25;
                cb = writer.DirectContent;
                // First we must activate writing
                cb.BeginText();
                // First we write out the header information
                // Start with the invoice type header
                writeText(cb, "Žádné služby nebyly objednány!", left_margin, top_margin, f_cb, 20);
                moveTop(60);
                cb.EndText();
            }

            catch (Exception rror)
            {
                System.Windows.Forms.MessageBox.Show(rror.Message);
            }
        }

        public void printServiceOverview(obClub club, List<obSluzby> clubSluzby)
        {
            try
            {
                if (clubSluzby.Count > 0)
                {
                    int left_margin = 40;

                    BaseFont f_cb = BaseFont.CreateFont("c:\\windows\\fonts\\calibrib.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                    BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);

                    cb = writer.DirectContent;
                    // First we must activate writing
                    cb.BeginText();
                    // First we write out the header information
                    // Start with the invoice type header
                    writeText(cb, club.ClubShort, left_margin, top_margin, f_cb, 14);
                    
                    clubSluzby.Sort((x, y) => x.Name.CompareTo(y.Name));
                    obSluzby prevServ = clubSluzby.First();
                    int count = -prevServ.Number; //v prvním kroku se vynuluje a při výstupu z cyklu přičte správná hodnota 
                    foreach (obSluzby service in clubSluzby)
                    {
                        if (service.Name.Equals(prevServ.Name))
                        {
                            count += prevServ.Number;
                        }
                        else
                        {
                            count += prevServ.Number;
                            writeText(cb, prevServ.Name, left_margin + 30, top_margin, f_cn, 10);
                            writeTextAlignedRight(cb, count.ToString(), left_margin + 200, top_margin, f_cn, 10);
                            writeTextAlignedRight(cb, prevServ.Fee + " Kč", left_margin + 275, top_margin, f_cn, 10);
                            writeTextAlignedRight(cb, (int.Parse(prevServ.Fee) * count) + " Kč", left_margin + 350, top_margin, f_cn, 10);
                            moveTop(12);
                            count = 0;
                        }
                        prevServ = service;
                    }
                    count += prevServ.Number;
                    writeText(cb, prevServ.Name, left_margin + 30, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, count.ToString(), left_margin + 200, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, prevServ.Fee + " Kč", left_margin + 275, top_margin, f_cn, 10);
                    writeTextAlignedRight(cb, (int.Parse(prevServ.Fee) * count) + " Kč", left_margin + 350, top_margin, f_cn, 10);
                    moveTop(12);
                    writeText(cb, "-------------------------------------------------------------------------------------------------------------------------------", left_margin, top_margin, f_cn, 10);
                    //tab total
                    moveTop(16);
                    cb.EndText();
                }
            }
            catch (Exception error)
            {
                System.Windows.Forms.MessageBox.Show(error.Message);
            }
        }


        public void printClubsOKHeader()
        {
            try
            {
                BaseFont f_cb = BaseFont.CreateFont("c:\\windows\\fonts\\calibrib.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                left_margin = 25;
                cb = writer.DirectContent;
                // First we must activate writing
                cb.BeginText();
                // First we write out the header information
                // Start with the invoice type header
                writeText(cb, "Kluby platby OK", left_margin, top_margin, f_cb, 50);
                moveTop(60);
                cb.EndText();
            }

            catch (Exception rror)
            {
                System.Windows.Forms.MessageBox.Show(rror.Message);
            }
        }
    

        public void printClubsOK(obClub club)
        {
            try
            {
                BaseFont f_cb = BaseFont.CreateFont("c:\\windows\\fonts\\calibrib.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1250, BaseFont.NOT_EMBEDDED);

                cb = writer.DirectContent;
                // First we must activate writing
                cb.BeginText();
                // First we write out the header information
                // Start with the invoice type header
                writeText(cb, club.ClubShort, left_margin, top_margin, f_cb, 50);
                left_margin += 135;
                if (left_margin > 500)
                {
                    left_margin = 25;
                    moveTop(60);
                }
                cb.EndText();
            }

            catch (Exception rror)
            {
                System.Windows.Forms.MessageBox.Show(rror.Message);
            }
        }
    }
}
