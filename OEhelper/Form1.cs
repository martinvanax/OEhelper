using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.IO.Ports;
using System.Timers;
using System.Threading;
using System.Net;
using System.Xml;
using bpac;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace OEHelper
{
    public partial class Form1 : Form
    {
        delegate void SetTextCallback(string text);
        
        Dictionary<int, string> category = new Dictionary<int, string>();
        //Dictionary<int, string> categoryRace1 = new Dictionary<int, string>();
        //Dictionary<int, string> categoryRace2 = new Dictionary<int, string>();
        Dictionary<int, string> clubs = new Dictionary<int, string>();
        //Dictionary<int, string> clubsRace1 = new Dictionary<int, string>();
        //Dictionary<int, string> clubsRace2 = new Dictionary<int, string>();
        /* 2017 for ORIS */
        List<EnvelopeLabel> envelopeLabels = new List<EnvelopeLabel>();
        BindingSource bsEnvelopeLabels = new BindingSource();
        List<CeremonyLabel> ceremonyLabels = new List<CeremonyLabel>();
        BindingSource bsCeremonyLabels = new BindingSource();
        WebClient webClient = new WebClient();
        XmlDocument xdoc;

        System.Timers.Timer timerSMSPause = new System.Timers.Timer(3000);

        public Form1()
        {
            InitializeComponent();

        }

        private void Log(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.textBoxLog.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(Log);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.textBoxLog.AppendText("\r\n" + text);
                this.textBoxLog.ScrollToCaret();
            }
        }

        private void buttonClearLog_Click(object sender, EventArgs e)
        {
            this.textBoxLog.Clear();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            nudOrisID.Value = Convert.ToDecimal(Properties.Settings.Default.ORISid);
            nudRentSIFee.Value = Properties.Settings.Default.rentSIFee;
            tbOrisSecKey.Text = Properties.Settings.Default.ORISSecKey;
            textBoxWorkingDir.Text = Properties.Settings.Default.labelWorkingDir;
            tabControl.SelectTab(Properties.Settings.Default.activeTab);
            tbOrganizerIdentificationLine1.Text = Properties.Settings.Default.organizerIdentificationL1;
            tbOrganizerIdentificationAbbr.Text = Properties.Settings.Default.organizerIdentificationAbbr;
            tbOrganizerIdentificationLine2.Text = Properties.Settings.Default.organizerIdentificationL2;
            tbOrganizerIdentificationLine3.Text = Properties.Settings.Default.organizerIdentificationL3;
            tbEventName.Text = Properties.Settings.Default.eventName;
            tbEventDate.Text = Properties.Settings.Default.eventDate;
            rbOE2010.Checked = Properties.Settings.Default.fromOE;
            rbORIS.Checked = !rbOE2010.Checked;
        }
        #region Vyuctovani poplatku
        List<string> obClubsShort = new List<string>();
        List<obClub> obClubs = new List<obClub>();
        List<oe2010Runner> oe2010Runners = new List<oe2010Runner>();
        List<obSluzby> obSluzbyFees = new List<obSluzby>();
        List<obPlatby> obPlatbyFees = new List<obPlatby>();

        private void buttonLoadClubsFee_Click(object sender, EventArgs e)
        {
            if (openFileDialogPoplatky.ShowDialog() == DialogResult.OK)
            {
                string clublistFile = openFileDialogPoplatky.FileName;
                loadClubs(clublistFile);
            }
        }

        private void loadClubs(string file)
        {
            try
            {
                StreamReader sr = new StreamReader(new FileStream(file, FileMode.Open, FileAccess.Read)/*, Encoding.GetEncoding(1250)*/);
                obClubs.Clear();
                /* IOF3.0 */
                xdoc = new XmlDocument();
                xdoc.Load(file);
                // Add the namespace.  
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xdoc.NameTable);
                nsmgr.AddNamespace("c", xdoc.DocumentElement.NamespaceURI);
                nsmgr.AddNamespace("ex", "ex");

                foreach (XmlNode xmlClub in xdoc.SelectNodes("//c:OrganisationList/c:Organisation", nsmgr))
                {
                    obClubs.Add(new obClub(xmlClub, nsmgr));
                }

                labelObClubsCount.Text = obClubs.Count.ToString();
            }
            catch (Exception e)
            {
                Log("LoadClubsFee: The file could not be read:");
                Log(e.Message);
            }
        }



        private void buttonLoadClubStartlist_Click(object sender, EventArgs e)
        {
            if (obClubs.Count == 0)
            {
                MessageBox.Show("Nemas nactene zadne kluby, stahni si aktualni z orisu, export IOFv3.0");
            }
            else
            {
                if (openFileDialogPoplatky.ShowDialog() == DialogResult.OK)
                {
                    string startlistFile = openFileDialogPoplatky.FileName;
                    loadOE2010Runners(startlistFile);
                }
            }
        }

        private void loadOE2010Runners(string file)
        {
            string line;
            string oeType;
            int lineCounter = 1;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                TextReader sr = new StreamReader(file, Encoding.GetEncoding(1250));
                oe2010Runners.Clear();
                oe2010Runner runner;
                if ((line = sr.ReadLine()) != null)
                {
                    oeType = line.Split(';').First();
                    if (!oeType.StartsWith("OE0001") && !oeType.StartsWith("OE0002"))
                    {
                        MessageBox.Show("File start with wrong tag import may be corrupted.\nBeggins with " + line.Substring(0, 50));
                    }
                    else
                    {
                        while ((line = sr.ReadLine()) != null)
                        {
                            lineCounter++;
                            int partsCount = line.Split(';').Length;
                            if ((oeType == "OE0001" && line.Split(';').Length >= 57)
                              ||(oeType == "OE0002" && line.Split(';').Length >= 108))
                            {
                                runner = new oe2010Runner(line, oeType, Properties.Settings.Default.fromOE);
                                if (runner.Name != "Vakant")
                                {
                                    oe2010Runners.Add(runner);
                                    if (null == obClubs.Find(c => (runner.ClubShort == c.ClubShort)))
                                    {
                                        obClubs.Add(new obClub(runner.ClubShort, runner.Club));
                                    }
                                }
                            }
                        }
                        labelRunnerFeeCount.Text = oe2010Runners.Count.ToString();
                        labelObClubsCount.Text = obClubs.Count.ToString();
                    }
                }
            }
            catch (Exception e)
            {
                Log("LoadRunners: The file could not be read: error on line: " + lineCounter);
                Log(e.Message);
                MessageBox.Show("LoadRunners Error check Log: " + e.Message);
            }
            Cursor.Current = Cursors.Default;
        }

        private void buttonSluzby_Click(object sender, EventArgs e)
        {
            if (obClubs.Count == 0)
            {
                MessageBox.Show("Nemas nactene zadne kluby, stahni si aktualni z orisu, export IOFv3.0");
            }
            else
            {
                if (openFileDialogPoplatky.ShowDialog() == DialogResult.OK)
                {
                    string sluzbyFeeFile = openFileDialogPoplatky.FileName;
                    loadSluzbyFees(sluzbyFeeFile);
                }
            }
        }

        private void loadSluzbyFees(string file)
        {
            try
            {
                StreamReader sr = new StreamReader(file);
                string line;
                int totalFee = 0;
                int number = 0;
                string type = string.Empty;
                obSluzbyFees.Clear();
                obPlatbyFees.Clear();
                while ((line = sr.ReadLine()) != null)
                {
                    /*
                    #ORIS#Sluzby#1#2013.08.20#21:25:39#2263#2013.08.24#TUR#
                    #ID; Klub;  Typ; Nazev                                  ;Pocet;JednotkovaCena; CelkovaCena; Poznamka;ProKoho ;DatumACas        ;Termin#
                    1179;KRL;   6;   Apartmán 4 postele + 2 přistýlky čt-pa; 3    ;1000          ; 3000       ;         ;KRL7151 ;20.06.2013 21:41 ;1
                    1180;KRL;6;Apartmán 4 postele + 2 přistýlky pa-ne;3;2400;7200;;KRL7151;20.06.2013 21:41;1

                    1485;FSP;7;Platba;1;360;360;FSP3400;;24.06.2013 00:00;0;
                    1515;ZBP;7;Platba;1;3450;3450;;;27.06.2013 00:00;0;

                     */
                    if (!line.StartsWith("#"))
                    {
                        string[] cols = line.Split(';');
                        if (int.TryParse(cols[6], out totalFee))
                        {
                            type = cols[2].Trim();
                            if (type == "6")
                            { /* sluzby */
                                if (int.TryParse(cols[4], out number))
                                {
                                    obSluzbyFees.Add(new obSluzby(cols[1], cols[3], number, cols[5], totalFee, cols[7], cols[8], cols[9]));
                                }
                                else
                                {
                                    MessageBox.Show("Chyba parsovani sluzby pocet: " + line);
                                }
                            }
                            else if (type == "7")
                            { /* platby */
                                obPlatbyFees.Add(new obPlatby(cols[1], cols[3], cols[4], cols[5], totalFee, cols[7], cols[8], cols[9]));
                            }
                            //Log(line);
                        }
                        else
                        {
                            MessageBox.Show("Chyba parsovani sluzby: " + line);
                        }
                    }
                }
                sr.Close();
                labelSluzbyFeeCount.Text = obSluzbyFees.Count.ToString();
                labelPlatbyFeeCount.Text = obPlatbyFees.Count.ToString();
            }
            catch (Exception e)
            {
                Log("LoadSluzbyFee: The file could not be read:");
                Log(e.Message);
            }
        }

        private void buttonComputeFees_Click(object sender, EventArgs e)
        {
            computeFees();
        }


        private void computeFees()
        {
            Cursor.Current = Cursors.WaitCursor;
            envelopeLabels.Clear();
            if (obClubs.Count == 0)
            {
                MessageBox.Show("Nemas nactene zadne kluby, stahni si aktualni z orisu, IOFv3.0");
                return;
            }
            try
            {
                int feeEntry;
                int feeSI;
                int feeService;
                int feeTotal;
                int paid = 0;
                int toBePaid = 0;

                int totalFeeEntry = 0;
                int totalFeeSI = 0;
                int totalFeeService = 0;
                int totalPaid = 0;
                int totalToBePaid = 0;

                XmlDocument xmlClub;

                FileStream fs = new FileStream(textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "vyuctovani_" + nudOrisID.Value + ".pdf", FileMode.Create);
                iTextSharp.text.Document document = new iTextSharp.text.Document(PageSize.A4, 25, 25, 30, 1);

                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                // Add meta information to the document
                document.AddAuthor("Martin Vana");
                document.AddCreator("OE2010 helper");
                //document.AddKeywords("PDF tutorial education");
                //document.AddSubject("Describing the steps creating a PDF document");
                document.AddTitle("Vyúčtování vkladů");
                // Open the document to enable you to write to the document
                document.Open();
                POSTemlates pt = new POSTemlates(document, writer);
                /* end prehled */
                /* pokladna */
                FileStream pfs = new FileStream(textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "pokladna_" + nudOrisID.Value + ".pdf", FileMode.Create);
                iTextSharp.text.Document pdocument = new iTextSharp.text.Document(PageSize.A6.Rotate(), 10, 10, 10, 10);

                PdfWriter pwriter = PdfWriter.GetInstance(pdocument, pfs);
                // Add meta information to the document
                pdocument.AddAuthor("Martin Vana");
                pdocument.AddCreator("OE2010 helper");
                //document.AddKeywords("PDF tutorial education");
                //document.AddSubject("Describing the steps creating a PDF document");
                pdocument.AddTitle("Pokladna");
                // Open the document to enable you to write to the document
                pdocument.Open();

                POSTemlates ppt = new POSTemlates(pdocument, pwriter);
                /* end pokladna */

                /* start souhrn sluzeb po klubech */
                FileStream sfs = new FileStream(textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "souhrn_sluzeb_" + nudOrisID.Value + ".pdf", FileMode.Create);
                iTextSharp.text.Document sdocument = new iTextSharp.text.Document(PageSize.A4, 25, 25, 30, 1);

                PdfWriter swriter = PdfWriter.GetInstance(sdocument, sfs);
                // Add meta information to the document
                sdocument.AddAuthor("Martin Vana");
                sdocument.AddCreator("OE2010 helper");
                //document.AddKeywords("PDF tutorial education");
                //document.AddSubject("Describing the steps creating a PDF document");
                sdocument.AddTitle("Souhrn služeb");
                // Open the document to enable you to write to the document
                sdocument.Open();

                POSTemlates spt = new POSTemlates(sdocument, swriter);

                /* end souhrn sluzeb po klubech */

                /* start kluby ok */
                FileStream kfs = new FileStream(textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "kluby_ok_" + nudOrisID.Value + ".pdf", FileMode.Create);
                iTextSharp.text.Document kdocument = new iTextSharp.text.Document(PageSize.A4, 25, 25, 30, 1);

                PdfWriter kwriter = PdfWriter.GetInstance(kdocument, kfs);
                // Add meta information to the document
                kdocument.AddAuthor("Martin Vana");
                kdocument.AddCreator("OE2010 helper");
                //document.AddKeywords("PDF tutorial education");
                //document.AddSubject("Describing the steps creating a PDF document");
                kdocument.AddTitle("Kluby OK");
                // Open the document to enable you to write to the document
                kdocument.Open();

                POSTemlates kpt = new POSTemlates(kdocument, kwriter, 760);
                kpt.printClubsOKHeader();
                /* end kluby ok */

                foreach (obClub club in obClubs)
                {
                    List<oe2010Runner> clubRunners = oe2010Runners.FindAll(r => (r.ClubShort == club.ClubShort));
                    if ((cb_onlyBillWithRunner.Checked == true) && (clubRunners.Count == 0))
                    {
                        continue;
                    }
                    List<obSluzby> clubSluzby = obSluzbyFees.FindAll(p => (p.Club == club.ClubShort));
                    List<obPlatby> clubPlatby = obPlatbyFees.FindAll(p => (p.Club == club.ClubShort));

                    if ((clubRunners.Count > 0) || (clubSluzby.Count > 0) || (clubPlatby.Count > 0))
                    { /* ma nejakou interakci se zavodem */
                        xmlClub = null;
                        if (cb_downloadClubInfoFromORIS.Checked)
                        {
                            xmlClub = downloadClubInfo(club.Id);
                        }
                        if (xmlClub != null)
                        {
                            club.OfficialName = xmlClub.SelectSingleNode("ORIS/Data/OfficialName").InnerText;
                            if (club.OfficialName != string.Empty)
                            {
                                club.OfficialStreet = xmlClub.SelectSingleNode("ORIS/Data/OfficialStreet").InnerText;
                                club.OfficialCity = xmlClub.SelectSingleNode("ORIS/Data/OfficialCity").InnerText;
                                club.OfficialZIP = xmlClub.SelectSingleNode("ORIS/Data/OfficialZIP").InnerText;
                                club.OfficialRegNum = xmlClub.SelectSingleNode("ORIS/Data/OfficialRegNum").InnerText;
                            }
                            else
                            {
                                club.Name = xmlClub.SelectSingleNode("ORIS/Data/Name").InnerText; /* vyresi kodovani */
                                club.ContactFirstName = xmlClub.SelectSingleNode("ORIS/Data/ContactFirstName").InnerText;
                                club.ContactLastName = xmlClub.SelectSingleNode("ORIS/Data/ContactLastName").InnerText;
                                club.ContactStreet = xmlClub.SelectSingleNode("ORIS/Data/ContactStreet").InnerText;
                                club.ContactCity = xmlClub.SelectSingleNode("ORIS/Data/ContactCity").InnerText;
                                club.ContactZIP = xmlClub.SelectSingleNode("ORIS/Data/ContactZIP").InnerText;
                            }
                        }
                        else
                        {
                            club.OfficialCity = "Debug";
                        }
                        pt.printBillOverview(club, clubRunners, clubSluzby, clubPlatby, out feeEntry, out feeSI, out feeService, out paid);
                        feeTotal = feeEntry + feeSI + feeService;
                        toBePaid = feeTotal - paid;
                        envelopeLabels.Add(new EnvelopeLabel(string.Empty, club.ClubShort, feeEntry.ToString(), feeSI.ToString(), feeService.ToString(), feeTotal.ToString(), paid.ToString(), toBePaid.ToString()));
                        if (toBePaid != 0)
                        {
                            ppt.printCashBill(club, toBePaid); //original
                            ppt.printCashBill(club, toBePaid); //kopie pro vydavatele
                        }
                        else
                        {
                            kpt.printClubsOK(club);
                        }

                        totalFeeEntry += feeEntry;
                        totalFeeSI += feeSI;
                        totalFeeService += feeService;
                        totalPaid += paid;

                        spt.printServiceOverview(club, clubSluzby);

                    }
                }
                // Close the document, the writer and the filestream! prehledy
                document.Close();
                writer.Close();
                fs.Close();
                // Close the document, the writer and the filestream! pokladna
                pdocument.Close();
                pwriter.Close();
                pfs.Close();
                // Close the document, the writer and the filestream! souhrn
                if (obSluzbyFees.Count == 0)
                {
                    spt.printServiceOverviewEmpty();
                }
                sdocument.Close();
                swriter.Close();
                sfs.Close();
                // Close the document, the writer and the filestream! kluby ok
                kdocument.Close();
                kwriter.Close();
                kfs.Close();

                //writer.WriteLine("\n\n\nHotovo hura rezat\nVklady = " + totalVklad + "\nPujcovne = " + totalPujcovne + "\nSluzby = " + totalSluzby + "\nPlatby = " + totalPlatbyPrijate + "\nCelkem vybrano = " +(totalVklad + totalSluzby + totalPujcovne));
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Bills created! Hotovo hura\nVklady = " + totalFeeEntry + "\nPujcovne = " + totalFeeSI + "\nSluzby = " + totalFeeService + "\nPlatby = " + totalPaid + "\nZbývá doplatit = " + (totalFeeEntry + totalFeeSI + totalFeeService - totalPaid) + "\nCelkem vybrano = " + (totalFeeEntry + totalFeeSI + totalFeeService));
                writer.Close();

                initDgvEnvelopeLabel();
                bsEnvelopeLabels.DataSource = envelopeLabels;
                dgvLabels.DataSource = bsEnvelopeLabels;
                bsEnvelopeLabels.ResetBindings(false);

            }
            catch (Exception e)
            {
                Log("ComputeFees: exception");
                Log(e.Message);
                MessageBox.Show("ComputeFees: exception" + e.Message);
            }
            Cursor.Current = Cursors.Default;

        }
        #endregion

        private void buttonWorkingDir_Click(object sender, EventArgs e)
        {
            folderBrowserDialogWD.SelectedPath = Properties.Settings.Default.labelWorkingDir;
            if (folderBrowserDialogWD.ShowDialog() == DialogResult.OK)
            {
                textBoxWorkingDir.Text = folderBrowserDialogWD.SelectedPath;
                Properties.Settings.Default.labelWorkingDir = textBoxWorkingDir.Text;
            }
        }

        private void nudOrisID_ValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ORISid = Convert.ToInt32(nudOrisID.Value);
        }

        private void tbOrisSecKey_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.ORISSecKey = tbOrisSecKey.Text;
        }

        private void buttonDownloadClubs_Click(object sender, EventArgs e)
        {
            try
            {
                string feesURL = "https://oris.orientacnisporty.cz/API/?format=xml&method=getEventBalance&eventid=" + nudOrisID.Value.ToString();
                string result = webClient.DownloadString(feesURL);
                if (result != null)
                {
                    File.WriteAllText(textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "eventBalance_" + nudOrisID.ToString() + ".xml", result);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void buttonDownloadFees_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                string feesURL = "https://oris.orientacnisporty.cz/API/?format=xml&method=getEventBalance&eventid=" + nudOrisID.Value.ToString();
                string result = webClient.DownloadString(feesURL);
                if (result != null)
                {
                    string file = textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "eventBalance_" + nudOrisID.Value.ToString() + ".xml";
                    File.WriteAllText(file, result);
                    loadFees(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            Cursor.Current = Cursors.Default;
        }

        private void buttonLoadFees_Click(object sender, EventArgs e)
        {
            openFileDialogXML.InitialDirectory = textBoxWorkingDir.Text;
            if (openFileDialogXML.ShowDialog() == DialogResult.OK)
            {
                loadFees(openFileDialogXML.FileName);
            }
        }

        private void loadFees(string filename)
        {
            string err = string.Empty;
            try
            {
                envelopeLabels.Clear();
                XmlDocument xdoc = new XmlDocument();  //https://msdn.microsoft.com/en-us/library/d271ytdx(v=vs.110).aspx
                xdoc.Load(filename);
                XmlNode root = xdoc.DocumentElement;

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xdoc.NameTable);
                nsmgr.AddNamespace("o", "oris");

                XmlNode nodeClubs = root.SelectSingleNode("Data/Clubs", nsmgr);
                foreach (XmlNode club in nodeClubs)
                {
                    if (club.Name.StartsWith("Club"))
                    {
                        string clubId = club.SelectSingleNode("ClubID", nsmgr).InnerText;
                        string clubAbbr = club.SelectSingleNode("ClubAbbr", nsmgr).InnerText;
                        string entry = club.SelectSingleNode("FeeEntry", nsmgr).InnerText;
                        string SI = club.SelectSingleNode("FeeSI", nsmgr).InnerText;
                        string service = club.SelectSingleNode("FeeService", nsmgr).InnerText;
                        string total = club.SelectSingleNode("FeeTotal", nsmgr).InnerText;
                        string paid = club.SelectSingleNode("Paid", nsmgr).InnerText;
                        string toBePaid = club.SelectSingleNode("ToBePaid", nsmgr).InnerText;
                        envelopeLabels.Add(new EnvelopeLabel(clubId, clubAbbr, entry, SI, service, total, paid, toBePaid));
                    }
                    else
                    {
                        err += "Unknown tag " + club.Name.ToString() + "\n";
                    }
                }

                initDgvEnvelopeLabel();
                bsEnvelopeLabels.DataSource = envelopeLabels;
                dgvLabels.DataSource = bsEnvelopeLabels;
                bsEnvelopeLabels.ResetBindings(false);

            }
            catch (Exception ex)
            {
                Log("Load Fees error\n" + ex.ToString());
            }
        }

        private void initDgvEnvelopeLabel()
        {
            // Initialize the DataGridView.
            dgvLabels.AutoGenerateColumns = false;
            dgvLabels.AutoSize = true;

            DataGridViewCheckBoxColumn selCol = new DataGridViewCheckBoxColumn();
            selCol.Name = "Select";
            selCol.DataPropertyName = "Select";
            selCol.ReadOnly = false;
            selCol.SortMode = DataGridViewColumnSortMode.Automatic;
            selCol.Width = 50;

            DataGridViewTextBoxColumn caCol = new DataGridViewTextBoxColumn();
            caCol.Name = "Abbr";
            caCol.DataPropertyName = "ClubAbbr";
            caCol.ReadOnly = true;

            DataGridViewTextBoxColumn idCol = new DataGridViewTextBoxColumn();
            idCol.Name = "ClubID";
            idCol.DataPropertyName = "ClubID";
            idCol.ReadOnly = true;

            DataGridViewTextBoxColumn clCol = new DataGridViewTextBoxColumn();
            clCol.Name = "Club";
            clCol.DataPropertyName = "Club";
            clCol.ReadOnly = true;

            DataGridViewTextBoxColumn entryCol = new DataGridViewTextBoxColumn();
            entryCol.Name = "Entry";
            entryCol.DataPropertyName = "Entry";
            entryCol.ReadOnly = true;
            entryCol.Width = 80;

            DataGridViewTextBoxColumn siCol = new DataGridViewTextBoxColumn();
            siCol.Name = "SI";
            siCol.DataPropertyName = "SI";
            siCol.ReadOnly = true;
            siCol.Width = 80;
            DataGridViewTextBoxColumn serviceCol = new DataGridViewTextBoxColumn();
            serviceCol.Name = "Service";
            serviceCol.DataPropertyName = "Service";
            serviceCol.ReadOnly = true;
            DataGridViewTextBoxColumn totalCol = new DataGridViewTextBoxColumn();
            totalCol.Name = "Total";
            totalCol.DataPropertyName = "Total";
            totalCol.ReadOnly = true;
            DataGridViewTextBoxColumn paidCol = new DataGridViewTextBoxColumn();
            paidCol.Name = "Paid";
            paidCol.DataPropertyName = "Paid";
            paidCol.ReadOnly = true; /* true; */
            DataGridViewTextBoxColumn tpaidCol = new DataGridViewTextBoxColumn();
            tpaidCol.Name = "ToBePaid";
            tpaidCol.DataPropertyName = "ToBePaid";
            tpaidCol.ReadOnly = true; /* true; */

            dgvLabels.Columns.Clear();
            dgvLabels.Columns.Add(selCol);
            dgvLabels.Columns.Add(idCol);
            dgvLabels.Columns.Add(caCol);
            dgvLabels.Columns.Add(clCol);
            dgvLabels.Columns.Add(entryCol);
            dgvLabels.Columns.Add(siCol);
            dgvLabels.Columns.Add(serviceCol);
            dgvLabels.Columns.Add(totalCol);
            dgvLabels.Columns.Add(paidCol);
            dgvLabels.Columns.Add(tpaidCol);

            // Add a CellClick handler to handle clicks in the button column.
            //dgvCANsim.CellClick += new DataGridViewCellEventHandler(dgvCANsim_CellClick);

            dgvLabels.Columns["ToBePaid"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private void initDgvCeremonyLabel()
        {
            // Initialize the DataGridView.
            dgvLabels.AutoGenerateColumns = false;
            dgvLabels.AutoSize = true;

            DataGridViewCheckBoxColumn selCol = new DataGridViewCheckBoxColumn();
            selCol.Name = "Select";
            selCol.DataPropertyName = "Select";
            selCol.ReadOnly = false;
            selCol.SortMode = DataGridViewColumnSortMode.Automatic;
            selCol.Width = 50;

            DataGridViewTextBoxColumn clCol = new DataGridViewTextBoxColumn();
            clCol.Name = "Class";
            clCol.DataPropertyName = "ClassName";
            clCol.ReadOnly = true;

            DataGridViewTextBoxColumn cecCol = new DataGridViewTextBoxColumn();
            cecCol.Name = "Current Entries";
            cecCol.DataPropertyName = "ClassCEC";
            cecCol.ReadOnly = true;

            DataGridViewTextBoxColumn placeCol = new DataGridViewTextBoxColumn();
            placeCol.Name = "Place";
            placeCol.DataPropertyName = "Place";
            placeCol.ReadOnly = true;

            dgvLabels.Columns.Clear();
            dgvLabels.Columns.Add(selCol);
            dgvLabels.Columns.Add(clCol);
            dgvLabels.Columns.Add(cecCol);
            dgvLabels.Columns.Add(placeCol);

            dgvLabels.Columns["Place"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private void buttonDownloadClubNames_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                foreach (EnvelopeLabel club in envelopeLabels)
                {
                    string clubURL;
                    if (club.ClubId == string.Empty)
                    {
                        clubURL = "https://oris.orientacnisporty.cz/API/?format=xml&method=getClub&id=" + club.ClubAbbr + "&eventkey=" + tbOrisSecKey.Text;
                        Log("Klub:{0} nemusí mít dobře údaje z ORISu, neexistuje u něj ID a udaje se zkusily získat podle Abbr=nespolehlivé!!!");
                    }
                    else
                    {
                        clubURL = "https://oris.orientacnisporty.cz/API/?format=xml&method=getClub&id=" + club.ClubId + "&eventkey=" + tbOrisSecKey.Text;
                    }
                    string result = webClient.DownloadString(clubURL);
                    if (result != null)
                    {
                        //string file = textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "club_" + club.ClubId + ".xml";
                        //File.WriteAllText(file, result);
                        XmlDocument xdoc = new XmlDocument();  //https://msdn.microsoft.com/en-us/library/d271ytdx(v=vs.110).aspx
                        xdoc.LoadXml(result);
                        club.Club = xdoc.SelectSingleNode("ORIS/Data/Name").InnerText;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            Cursor.Current = Cursors.Default;
            bsEnvelopeLabels.ResetBindings(false);
        }

        private XmlDocument downloadClubInfo(string id)
        {
            string clubURL = "https://oris.orientacnisporty.cz/API/?format=xml&method=getClub&id=" + id + "&eventkey=" + tbOrisSecKey.Text;
            string result = webClient.DownloadString(clubURL);
            XmlDocument xdoc = null;
            if (result != null)
            {
                //string file = textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "club_" + club.ClubId + ".xml";
                //File.WriteAllText(file, result);
                xdoc = new XmlDocument();  //https://msdn.microsoft.com/en-us/library/d271ytdx(v=vs.110).aspx
                xdoc.LoadXml(result);
                if ("OK" != xdoc.SelectSingleNode("ORIS/Status").InnerText)
                {
                    xdoc = null;
                }
            }
            return xdoc;
        }

        private void buttonDownloadEventInfo_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                string feesURL = "https://oris.orientacnisporty.cz/API/?format=xml&method=getEvent&id=" + nudOrisID.Value.ToString();
                string result = webClient.DownloadString(feesURL);
                if (result != null)
                {
                    string file = textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "eventInfo_" + nudOrisID.Value.ToString() + ".xml";
                    File.WriteAllText(file, result);
                    loadCategoriesCeremony(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            Cursor.Current = Cursors.Default;
        }

        private void buttonLoadCategories_Click(object sender, EventArgs e)
        {
            openFileDialogXML.InitialDirectory = textBoxWorkingDir.Text;
            if (openFileDialogXML.ShowDialog() == DialogResult.OK)
            {
                loadCategoriesCeremony(openFileDialogXML.FileName);
            }
        }

        /* <Classes><Class_85292><ID>85292</ID><Name>D10L</Name><Distance>0.00</Distance><Climbing>0</Climbing><Controls>0</Controls><PreFormattedHeader></PreFormattedHeader><Splits>0</Splits><ClassDefinition><ID>43</ID><AgeFrom>0</AgeFrom><AgeTo>10</AgeTo><Gender>F</Gender><Name>D10</Name></ClassDefinition><Fee>300</Fee><NoExtraFee>1</NoExtraFee><ManualFee>0</ManualFee><ManualFeeEntryDate2>0</ManualFeeEntryDate2><ManualFeeEntryDate3>0</ManualFeeEntryDate3><Ranking>0</Ranking><RankingKoef>0</RankingKoef><RankingKS>0</RankingKS><CurrentEntriesCount>35</CurrentEntriesCount></Class_85292>*/
        private void loadCategoriesCeremony(string filename)
        {
            string err = string.Empty;
            try
            {
                ceremonyLabels.Clear();
                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(filename);
                XmlNode root = xdoc.DocumentElement;

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xdoc.NameTable);
                nsmgr.AddNamespace("o", "oris");

                loadEventInfo(root, nsmgr);

                XmlNode nodeClasses = root.SelectSingleNode("Data/Classes", nsmgr);
                foreach (XmlNode nodeClass in nodeClasses)
                {
                    if (nodeClass.Name.StartsWith("Class"))
                    {
                        string className = nodeClass.SelectSingleNode("Name", nsmgr).InnerText;
                        string classCEC = nodeClass.SelectSingleNode("CurrentEntriesCount", nsmgr).InnerText;
                        for (int place = 1; place <= nudCeremonyPlaces.Value; place++)
                        {
                            ceremonyLabels.Add(new CeremonyLabel(className, classCEC, place + "."));
                        }
                    }
                    else
                    {
                        err += "Unknown tag " + nodeClass.Name.ToString() + "\n";
                    }
                }

                initDgvCeremonyLabel();
                bsCeremonyLabels.DataSource = ceremonyLabels;
                dgvLabels.DataSource = bsCeremonyLabels;
                bsCeremonyLabels.ResetBindings(false);

            }
            catch (Exception ex)
            {
                Log("Load Event Info categories for ceremony error\n" + ex.ToString());
            }
        }

        private void loadEventInfo(XmlNode xdoc, XmlNamespaceManager nsmgr)
        {
            try
            {
                XmlNode nodeRentSI = xdoc.SelectSingleNode("Data/EntryRentSIFee", nsmgr);
                if (nodeRentSI != null)
                {
                    nudRentSIFee.Value = Convert.ToDecimal(nodeRentSI.InnerText);
                    Log("Rent SI Fee loaded from ORIS: " + nudRentSIFee.Value);
                }
                XmlNode nodeEventName = xdoc.SelectSingleNode("Data/Name", nsmgr);
                XmlNode nodeEventDate = xdoc.SelectSingleNode("Data/Date", nsmgr);

                if ((nodeEventName != null) && (nodeEventDate != null))
                {
                    string date = DateTime.ParseExact(nodeEventDate.InnerText, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture).ToString("d.M.yyyy");
                    tbEventName.Text = nodeEventName.InnerText + " " + date;
                    tbEventDate.Text = date;
                    Log("Event name and date loaded from ORIS: " + tbEventName.Text);
                }
            }
            catch (InvalidCastException)
            {
                Log("Error to parse Event Info, Name, Date, Rent SI Fee");
            }

        }

        private void buttonLabelsSelectAll_Click(object sender, EventArgs e)
        {
            if (dgvLabels.DataSource == bsCeremonyLabels)
            {
                foreach (CeremonyLabel cl in ceremonyLabels)
                {
                    cl.Select = true;
                }
                bsCeremonyLabels.ResetBindings(false);
            }
            else if (dgvLabels.DataSource == bsEnvelopeLabels)
            {
                foreach (EnvelopeLabel el in envelopeLabels)
                {
                    el.Select = true;
                }
                bsEnvelopeLabels.ResetBindings(false);
            }
        }

        private void buttonLabelsSelectNone_Click(object sender, EventArgs e)
        {
            if (dgvLabels.DataSource == bsCeremonyLabels)
            {
                foreach (CeremonyLabel cl in ceremonyLabels)
                {
                    cl.Select = false;
                }
                bsCeremonyLabels.ResetBindings(false);
            }
            else if (dgvLabels.DataSource == bsEnvelopeLabels)
            {
                foreach (EnvelopeLabel el in envelopeLabels)
                {
                    el.Select = false;
                }
                bsEnvelopeLabels.ResetBindings(false);
            }

        }

        private void buttonLabelsPrint_Click(object sender, EventArgs e)
        {
            if (dgvLabels.DataSource == bsCeremonyLabels)
            {
                //MessageBox.Show("Ceremony");
                foreach (CeremonyLabel cl in ceremonyLabels)
                {
                    if (cl.Select == true)
                    {
                        printCeremonyLabel(cl);
                    }
                }
            }
            else if (dgvLabels.DataSource == bsEnvelopeLabels)
            {
                //MessageBox.Show("Envelope");
                foreach (EnvelopeLabel el in envelopeLabels)
                {
                    if (el.Select == true)
                    {
                        if (cbClubShortOnly.Checked)
                        {
                            printEnvelopeClubOnlyLabel(el);
                        }
                        else
                        {
                            printEnvelopeLabel(el);
                        }
                    }
                }
            }
        }

        private void printEnvelopeLabel(EnvelopeLabel el)
        {
            string templatePath = "obalka.lbx";

            bpac.DocumentClass doc = new bpac.DocumentClass();
            if (doc.Open(templatePath) != false)
            {
                doc.GetObject("ClubAbbr").Text = el.ClubAbbr;
                doc.GetObject("Club").Text = el.Club;
                doc.GetObject("FeeEntry").Text = el.Entry;
                doc.GetObject("FeeSI").Text = el.SI;
                doc.GetObject("FeeService").Text = el.Service;
                doc.GetObject("FeeTotal").Text = el.Total;
                doc.GetObject("Paid").Text = el.Paid;
                doc.GetObject("ToBePaid").Text = el.ToBePaid;

                // doc.SetMediaById(doc.Printer.GetMediaId(), true);
                doc.StartPrint("", PrintOptionConstants.bpoNoCut);
                doc.PrintOut(1, PrintOptionConstants.bpoDefault);
                doc.EndPrint();
                doc.Close();
            }
            else
            {
                MessageBox.Show("Open() Error: " + doc.ErrorCode);
            }
        }

        private void printEnvelopeClubOnlyLabel(EnvelopeLabel el)
        {
            string templatePath = "obalkaClub.lbx";

            bpac.DocumentClass doc = new bpac.DocumentClass();
            if (doc.Open(templatePath) != false)
            {
                doc.GetObject("ClubAbbr").Text = el.ClubAbbr;
                doc.GetObject("Club").Text = el.Club;
                
                // doc.SetMediaById(doc.Printer.GetMediaId(), true);
                doc.StartPrint("", PrintOptionConstants.bpoNoCut);
                doc.PrintOut(1, PrintOptionConstants.bpoDefault);
                doc.EndPrint();
                doc.Close();
            }
            else
            {
                MessageBox.Show("Open() Error: " + doc.ErrorCode);
            }
        }

        private void printCeremonyLabel(CeremonyLabel cl)
        {
            string templatePath = "vyhlaseni.lbx";

            bpac.DocumentClass doc = new bpac.DocumentClass();
            if (doc.Open(templatePath) != false)
            {
                doc.GetObject("Class").Text = cl.ClassName;
                doc.GetObject("Place").Text = cl.Place;

                doc.StartPrint("", PrintOptionConstants.bpoNoCut);
                doc.PrintOut(1, PrintOptionConstants.bpoDefault);
                doc.EndPrint();
                doc.Close();
            }
            else
            {
                MessageBox.Show("Open() Error: " + doc.ErrorCode);
            }
        }

        private void buttonPokladna_Click(object sender, EventArgs e)
        {
            testDownload();    
        }

        private void testDownload()
        {
            string servicesURL = "https://oris.orientacnisporty.cz/Exporty?agenda=Sluzby&comp=" + nudOrisID.Value.ToString();
            string result = webClient.DownloadString(servicesURL);
            if (result != null)
            {
                Log(result);
            }    
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.activeTab = tabControl.SelectedIndex;
        }

        private void buttonClubsAndServicesFeesFromOris_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                string clubsURL = "https://oris.orientacnisporty.cz/ExportIOF30?agenda=clubs&eventkey=" + tbOrisSecKey.Text;
                string result = webClient.DownloadString(clubsURL);
                if (result != null)
                {
                    string clubsTempFile = textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "clubs_" + nudOrisID.Value.ToString() + ".xml";
                    File.WriteAllText(clubsTempFile, result);
                    loadClubs(clubsTempFile);
                    
                    string servicesURL = "https://oris.orientacnisporty.cz/Exporty?agenda=Platby&comp=" + nudOrisID.Value.ToString();
                    result = webClient.DownloadString(servicesURL);
                    if (result != null)
                    {
                        string servicesTempFile = textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "platby_" + nudOrisID.Value.ToString() + ".csv";
                        File.WriteAllText(servicesTempFile, result, Encoding.GetEncoding(1250));
                        loadSluzbyFees(servicesTempFile);                        
                    }
                    else
                    {
                        MessageBox.Show("Chyba download sluzby a platby");
                    }
                }
                else
                {
                    MessageBox.Show("Chyba download kluby");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error message:" + ex.ToString());
            }
            Cursor.Current = Cursors.Default;
        }

        private void buttonDownloadRunners_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                string clubsURL = "https://oris.orientacnisporty.cz/ExportPrihlasek?mode=oe2010&encoding=win1250&id=" + nudOrisID.Value.ToString();
                string result = webClient.DownloadString(clubsURL);
                if (result != null)
                {
                    string runnersTempFile = textBoxWorkingDir.Text + Path.DirectorySeparatorChar + "export_entry_" + nudOrisID.Value.ToString() + ".csv";
                    File.WriteAllText(runnersTempFile, result,Encoding.GetEncoding(1250));
                    loadOE2010Runners(runnersTempFile);
                }
                else
                {
                    MessageBox.Show("Chyba download přihlášky-závodníci");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error message:" + ex.ToString());
            }
            Cursor.Current = Cursors.Default;
        }

        private void tbOrganizerIdentification_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.organizerIdentificationL1 = tbOrganizerIdentificationLine1.Text;
        }

        private void nudRentSIFee_ValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.rentSIFee = nudRentSIFee.Value;
        }

        private void tbOrganizerIdentificationLine2_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.organizerIdentificationL2 = tbOrganizerIdentificationLine2.Text;
        }

        private void tbOrganizerIdentificationLine3_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.organizerIdentificationL3 = tbOrganizerIdentificationLine3.Text;
        }

        private void tbEventName_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.eventName = tbEventName.Text;
        }

        private void tbOrganizerIdentificationAbbr_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.organizerIdentificationAbbr = tbOrganizerIdentificationAbbr.Text;
        }

        private void tbEventDate_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.eventDate = tbEventDate.Text;
        }

        private void rbOE2010_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.fromOE = rbOE2010.Checked;
        }
    }
}
