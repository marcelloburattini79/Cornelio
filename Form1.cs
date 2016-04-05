using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using MySql.Data;
using System.IO;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;
using Microsoft.Win32;
using System.Diagnostics;
using System.Net.Mail;
using ComponentFactory.Krypton.Toolkit;
using Newtonsoft.Json;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections;



namespace Cornelio
{
    public partial class Schedule : Form
    {
        //dichiarazione oggetti
        #region
        
        //stringhe per server centrale
        //public static MySqlConnection Georgia = new MySqlConnection(@"SERVER=192.168.1.200; Database=carolinapanthers2016; User ID=OGL; Password=ogl; Connection Timeout = 120");
        //public static MySqlConnection Wyoming = new MySqlConnection(@"SERVER=192.168.1.200; Database=carolinapanthers2016; User ID=OGL; Password=ogl");

        //stringhe per PC06
        //public static MySqlConnection Georgia = new MySqlConnection(@"SERVER=192.168.1.3; Database=carolinapanthers2016; User ID=marcello; Password=jesussss@79; Connection Timeout = 120");
        //public static MySqlConnection Wyoming = new MySqlConnection(@"SERVER=192.168.1.3; Database=carolinapanthers2016; User ID=marcello; Password=jesussss@79");

        //stringhe per RdTL
        public static MySqlConnection Georgia = new MySqlConnection(@"SERVER=192.168.1.250; Database=carolinapanthers2016; User ID=marcello; Password=jesussss@79; Connection Timeout = 120");
        public static MySqlConnection Wyoming = new MySqlConnection(@"SERVER=192.168.1.250; Database=carolinapanthers2016; User ID=marcello; Password=jesussss@79");

        MySqlDataAdapter Adattatore;
        MySqlDataAdapter adattatoreDettagli;
        public DataTable Tavoletta;
        MySqlCommandBuilder Renfresca;
        MySqlCommandBuilder dettagliRenfresca;
        public BindingSource Sorgente;
        DataSet cassapanca;

        public Utente UtenteAttivo;
        public FiltriAvanzati FiltriAvanzati1;
        public StatoCampioni statoCampioni1;

        public StreamWriter FileLog;
        public bool VistaCampioni = false;
        public bool sediciNoni = true;
        
        bool Ancorato = true;
        bool FiltriGBCliccato = false;
        int deltaX = 0;
        int deltaY = 0;
        int RegolaX = 0;

        public Excel.Application ExcApp;
        public Excel.Workbook OGL;
        public Excel.Worksheet AC;

        static string Key = "gaftisnmmnsbhjksllopsspmsnsbgfdr";
        static string IV = "qwmmclopplcngsda";

        bool ballControl = false;

        int provaScadenza = 0;
        public int pretrattamentoRow = 0;
        public DateTime ?DataScadenza = null;

        public bool chiuso = false;

        public bool cambiamento = false;

        bool fattore = false;

        bool allargato = false;

        public bool messaggio = true;

        AggiornaColonne AggiornaColonne1;
        public SlotImpostazioni SlotImpostazioni1;

        string copiaSt = "";

        CampoNote campoNote1;

        public string chiamaStato = "";

        public Hashtable reverse = new Hashtable();

        MySqlCommand caricaAD = null;

        public int idSelezionato = 0;
        
        #endregion

        public Schedule()
        {
            InitializeComponent();
        }

        private void Schedule_Load(object sender, EventArgs e)
        {

            //LoadToDoListGV();

            //LoadComboFilter(sender, e);


            SelettivaCB.Checked = true;
            DynamicRB.Checked = true;

            Decifra(@"C:\Geminus\FileLogCfr.txt", @"C:\Geminus\FileLog.txt");
            FileLog = new StreamWriter(@"C:\Geminus\FileLog.txt", true);

            reverse["Parametro"] = false;
            reverse["Accettazione"] = false;
            reverse["Area"] = false;
            reverse["Metodo"] = false;
            reverse["Firma"] = false;
            reverse["AssegnatoA"] = false;
            reverse["StatoCampione"] = false;
            reverse["RapportoDiProva"] = false;
            reverse["Matrice"] = false;

            vediTuttoTT.SetToolTip(vediTuttoBT, "Mostra tutte le colonne della griglia");
            AggiornaTT.SetToolTip(AggiornaBT, "Aggiorna la tabella con i risultati inseriti da altri utenti");
            salvaDatiTT.SetToolTip(SalvaDatiBT, "Salva dati inseriti");
            stampaTT.SetToolTip(StampaBT, "Stampa dati in tabella");
            dettagliTT.SetToolTip(DettagliBT, "Mostra i dettagli della prova. Se si tiene premuto il tasto Ctr verranno visualizzati i dettagli del campione");
            rubricaTT.SetToolTip(RubricaBT, "Visualizza le schede del personale");
            mascheraUtentiTT.SetToolTip(MascheraUtenteBT, "Mostra il profilo utente");
            statoCampioniTT.SetToolTip(StatoCampioniBT, "Mostra lo stato di avanzamento dei campioni");
            stampatoTT.SetToolTip(StampatoBT, "Mostra i campioni certificati");
            completatoTT.SetToolTip(CompletatoBT, "Mostra in attesa di essere certificati");
            inAnalisiTT.SetToolTip(InAnalisiBT, "Mostra i campioni in analisi");
            scegliTecnicoTT.SetToolTip(ScegliTecnicoBT, "Assegna le prove selezionate all'utente indicato nella combobox");
            lucchettoTT.SetToolTip(LucchettoBT, "Vincola l'esecuzone delle prove selezionate all'utente indicato nella combobox");
            sincronizzaTT.SetToolTip(SincronizzaOGLBT, "Inserisce in tabella i parametri dei campioni non accettati al momento del login");
            monitorTT.SetToolTip(MonitorBT, "Passa alla modalità di visualizzazione per monitor in 4/4");
            filtriAvanzatiTT.SetToolTip(FiltriAvanzatiBT, "Apre la schermata dei Filtri Avanzati");
            salvaVistaTT.SetToolTip(SalvaVistaBT, "Visualizza la schermata per la gestione dei filtri");
            loadVistaTT.SetToolTip(LoadVistaBT, "Carica le impostazioni visuali di default impostate dall'utente");
            wordTT.SetToolTip(wordBT, "Apre il file word relativo al campione selezionato");
            excelTT.SetToolTip(excelBT, "Apre il foglio di calcolo relativo al campione selezionato");
            logTT.SetToolTip(VediFileLogBT, "Apre il file di log relativo alle operazioni compiute su questo calcolatore");
            zipTT.SetToolTip(ZipBT, "Tenere premuto e trascinare per ridimensionare la zona dei filtri");
            nascondiTuttoTT.SetToolTip(nascondiTutto, "Nasconde tutte le colonne della tabella");
            filtriTT.SetToolTip(FiltriBT, "Filtra secondo i criteri scelti nelle combobox");
            togliFiltriTT.SetToolTip(TogliFiltriBT, "Cancella tutti i criteri impostati nelle combobox");
            salvaFilterTT.SetToolTip(SalvaFilterBT, "Mostra la schermata per la gestione dei filtri");
            caricaFilterTT.SetToolTip(SalvaFiltroBT, "Carica i filtri impostati di default");
            disconettiTT.SetToolTip(DisconnettiBT, "Disconnette l'utente registrato");
            anteprimaTT.SetToolTip(AnteprimaStampaBT, "Crea una pagina pdf con i dati in tabella");

            Georgia.Open();
            MySqlCommand dataAggiornamento = new MySqlCommand("select * from Data", Georgia);
            MySqlDataReader WDTVDA = dataAggiornamento.ExecuteReader();
            while (WDTVDA.Read())
            {
                AggiornatoLB.Text = WDTVDA[0].ToString();
            }
            Georgia.Close();

            Login Login1 = new Login();
            Login1.ShowDialog();

            if (!chiuso)
            {
                Inizializzazione inizializzazione1 = new Inizializzazione();
                inizializzazione1.ShowDialog();

                if (UtenteAttivo.Nome != "Pasquale" && UtenteAttivo.Nome != "Adamo" && UtenteAttivo.Nome != "Marcello" && UtenteAttivo.Nome != "Francesco")
                {
                    vectorBT.Enabled = false;
                    vectorBT.Visible = false;
                }

                ToDoListGV.FirstDisplayedScrollingRowIndex = ToDoListGV.Rows.Count - 1;
            }

            
        }

        public void LoadComboFilter(object sender, EventArgs e)
        {

            // Carica combobox speed

            Georgia.Open();

            MySqlCommand speedCB = new MySqlCommand("select Accettazione, Parametro, Area, Metodo, ScadenzaAnalisi, DataArrivo, DataAnalisi, Scadenza, Firma, AssegnatoA, Preparativa, Determinazione, Quantificazione, Matrice, StatoCampione, RapportoDiProva, DataRdP from caricolavoro", Georgia);

            MySqlDataReader speedDR = speedCB.ExecuteReader();

            while (speedDR.Read())
            {
                if (!AccettazioneCB.Items.Contains(speedDR["Accettazione"]))
                {
                    AccettazioneCB.Items.Add(speedDR["Accettazione"]);
                }
                if (!ParametroCB.Items.Contains(speedDR["Parametro"]))
                {
                    ParametroCB.Items.Add(speedDR["Parametro"]);
                }
                if (!AreaCB.Items.Contains(speedDR["Area"]))
                {
                    AreaCB.Items.Add(speedDR["Area"]);
                }
                if (!MetodoCB.Items.Contains(speedDR["Metodo"]))
                {
                    MetodoCB.Items.Add(speedDR["Metodo"]);
                }
                if (!ScadenzaAnalisiCB.Items.Contains(speedDR["ScadenzaAnalisi"]))
                {
                    ScadenzaAnalisiCB.Items.Add(speedDR["ScadenzaAnalisi"]);
                }
                if (!DalCB.Items.Contains(speedDR["Accettazione"]))
                {
                    DalCB.Items.Add(speedDR["Accettazione"]);
                }
                if (!AlCB.Items.Contains(speedDR["Accettazione"]))
                {
                    AlCB.Items.Add(speedDR["Accettazione"]);
                }
                if (!DataArrivoCB.Items.Contains(speedDR["DataArrivo"]))
                {
                    DataArrivoCB.Items.Add(speedDR["DataArrivo"]);
                }
                if (!DataAnalisiCB.Items.Contains(speedDR["DataAnalisi"]))
                {
                    DataAnalisiCB.Items.Add(speedDR["DataAnalisi"]);
                }
                if (!DalScadenzaCB.Items.Contains(speedDR["Scadenza"]))
                {
                    DalScadenzaCB.Items.Add(speedDR["Scadenza"]);
                }
                if (!AlScadenzaCB.Items.Contains(speedDR["Scadenza"]))
                {
                    AlScadenzaCB.Items.Add(speedDR["Scadenza"]);
                }
                if (!ScadenzaCB.Items.Contains(speedDR["Scadenza"]))
                {
                    ScadenzaCB.Items.Add(speedDR["Scadenza"]);
                }
                if (!FirmaCB.Items.Contains(speedDR["Firma"]))
                {
                    FirmaCB.Items.Add(speedDR["Firma"]);
                }
                if (!AssegnatoACB.Items.Contains(speedDR["AssegnatoA"]))
                {
                    AssegnatoACB.Items.Add(speedDR["AssegnatoA"]);
                }
                if (!PreparativaCB.Items.Contains(speedDR["Preparativa"]))
                {
                    PreparativaCB.Items.Add(speedDR["Preparativa"]);
                }
                if (!DeterminazioneCB.Items.Contains(speedDR["Determinazione"]))
                {
                    DeterminazioneCB.Items.Add(speedDR["Determinazione"]);
                }
                if (!QuantificazioneCB.Items.Contains(speedDR["Quantificazione"]))
                {
                    QuantificazioneCB.Items.Add(speedDR["Quantificazione"]);
                }
                if (!MatriceCB.Items.Contains(speedDR["Matrice"]))
                {
                    MatriceCB.Items.Add(speedDR["Matrice"]);
                }
                if (!StatoCampioneCB.Items.Contains(speedDR["StatoCampione"]))
                {
                    StatoCampioneCB.Items.Add(speedDR["StatoCampione"]);
                }
                if (!RapportoDiProvaCB.Items.Contains(speedDR["RapportoDiProva"]))
                {
                    RapportoDiProvaCB.Items.Add(speedDR["RapportoDiProva"]);
                }
                if (!DataRdPCB.Items.Contains(speedDR["DataRdP"]))
                {
                    DataRdPCB.Items.Add(speedDR["DataRdP"]);
                }
                if (!DalRdPCB.Items.Contains(speedDR["RapportoDiProva"]))
                {
                    DalRdPCB.Items.Add(speedDR["RapportoDiProva"]);
                }
                if (!AlRdPCB.Items.Contains(speedDR["RapportoDiProva"]))
                {
                    AlRdPCB.Items.Add(speedDR["RapportoDiProva"]);
                }
                if (!DalDataRdPCB.Items.Contains(speedDR["DataRdP"]))
                {
                    DalDataRdPCB.Items.Add(speedDR["DataRdP"]);
                }
                if (!AlDataRdPCB.Items.Contains(speedDR["DataRdP"]))
                {
                    AlDataRdPCB.Items.Add(speedDR["DataRdP"]);
                }
            }

            if(speedDR != null)
            {
                speedDR.Close();
            }

            foreach (Control CB in FiltriGB.Panel.Controls)
            {
                if (CB is KryptonComboBox)
                {
                    (CB as KryptonComboBox).Sorted = true;
                }
            }

            Georgia.Close();

            Georgia.Open();

            MySqlCommand tecnicoMyS = new MySqlCommand("select Cognome from personale", Georgia);

            MySqlDataReader tecnicoDR = tecnicoMyS.ExecuteReader();

            while (tecnicoDR.Read())
            {
                if (!SceltaTecnicoCB.Items.Contains(tecnicoDR["Cognome"]))
                {
                    SceltaTecnicoCB.Items.Add(tecnicoDR["Cognome"]);
                }
            }

            SceltaTecnicoCB.Sorted = true;

            StatoParametroCB.SelectedIndex = 0;

            FiltriBT_Click(sender, e);

            Georgia.Close();
        }

        public void LoadToDoListGV()
        {

            Tavoletta = new DataTable();
            Sorgente = new BindingSource();
            BindingSource dettagliSorgente = new BindingSource();

            ToDoListGV.DataSource = Sorgente;
            DettagliGV.DataSource = dettagliSorgente;

            cassapanca = new DataSet();

            Adattatore = new MySqlDataAdapter("select * from caricolavoro", Georgia);
            Adattatore.Fill(cassapanca, "caricolavoro");

            caricaAD = new MySqlCommand();
            caricaAD.CommandTimeout = 240;
            caricaAD.CommandText = "select * from dettagli";
            caricaAD.Connection = Georgia;

            adattatoreDettagli = new MySqlDataAdapter();
            adattatoreDettagli.SelectCommand = caricaAD;
            adattatoreDettagli.Fill(cassapanca, "dettagli");

            DataRelation papaFigli = new DataRelation
                ("famiglia", cassapanca.Tables["caricolavoro"].Columns["ID"], cassapanca.Tables["dettagli"].Columns["figlioDi"]);
            cassapanca.Relations.Add(papaFigli);
            
            Sorgente.DataSource = cassapanca;
            Sorgente.DataMember = "caricolavoro";

            dettagliSorgente.DataSource = Sorgente;
            dettagliSorgente.DataMember = "famiglia";

            Adattatore.Update(cassapanca, "caricolavoro");

            //Formattazione DataGridView
            ToDoListGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
            ToDoListGV.RowsDefaultCellStyle.BackColor = Color.LightBlue;
            ToDoListGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSkyBlue;
            ToDoListGV.Columns["ID"].Visible = false;
            ToDoListGV.Columns["AccettazioneInNumero"].Visible = false;
            ToDoListGV.Columns["DataPreparativa"].Visible = false;
            ToDoListGV.Columns["DataQuantificazione"].Visible = false;
            ToDoListGV.Columns["DataDeterminazione"].Visible = false;
            ToDoListGV.Columns["TecnicoPreparativa"].Visible = false;
            ToDoListGV.Columns["DataDeterminazione"].Visible = false;
            ToDoListGV.Columns["TecnicoDeterminazione"].Visible = false;
            ToDoListGV.Columns["RiferimentoLocked"].Visible = false;
            ToDoListGV.Columns["RiferimentoDeperibile"].Visible = false;
            ToDoListGV.Columns["RiferimentoAccredia"].Visible = false;
            ToDoListGV.Columns["NumeroAccredia"].Visible = false;
            ToDoListGV.Columns["RdPInNumero"].Visible = false;
            ToDoListGV.Columns["Stato"].Visible = false;
            ToDoListGV.Columns["SoloIntestazione"].Visible = false;
            ToDoListGV.Columns["Strumento"].Visible = false;
            ToDoListGV.Columns["idParametro"].Visible = false;
            ToDoListGV.Columns["CodiceFamiglia"].Visible = false;
            ToDoListGV.Columns["Accettazione"].Width = 75;
            ToDoListGV.Columns["Preparativa"].Width = 75;
            ToDoListGV.Columns["Determinazione"].Width = 90;
            ToDoListGV.Columns["Quantificazione"].Width = 90;
            ToDoListGV.Columns["Stato"].Width = 75;
            ToDoListGV.Columns["DataAnalisi"].Width = 75;
            ToDoListGV.Columns["Firma"].Width = 105;
            ToDoListGV.Columns["Deperibile"].Width = 60;
            ToDoListGV.Columns["LimiteA"].Width = 60;
            ToDoListGV.Columns["LimiteB"].Width = 60;


            ToDoListGV.Columns["ScadenzaAnalisi"].HeaderText = "Scadenza deperibili";
            ToDoListGV.Columns["DataAnalisi"].HeaderText = "Data di analisi";
            ToDoListGV.Columns["AssegnatoA"].HeaderText = "Assegnato a...";
            ToDoListGV.Columns["StatoCampione"].HeaderText = "Stato del campione";
            ToDoListGV.Columns["DataArrivo"].HeaderText = "Data di arrivo";
            ToDoListGV.Columns["RapportoDiProva"].HeaderText = "Rapporto di prova";
            ToDoListGV.Columns["DataArrivo"].HeaderText = "Data di inserimento";

            
            DataGridViewCellStyle StileFirma = new DataGridViewCellStyle();
            StileFirma.Font = new System.Drawing.Font("Lucida Handwriting", 8);
            StileFirma.ForeColor = Color.Blue;
            ToDoListGV.Columns["Firma"].DefaultCellStyle = StileFirma;

            ToDoListGV.Sort(ToDoListGV.Columns["Accettazione"], ListSortDirection.Ascending);

            //Formattazione DataGridView dettagli
            DettagliGV.RowsDefaultCellStyle.BackColor = Color.LightBlue;
            DettagliGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSkyBlue;
            DettagliGV.Columns["ID"].Visible = false;
            DettagliGV.Columns["AccettazioneInNumero"].Visible = false;
            DettagliGV.Columns["DataPreparativa"].Visible = false;
            DettagliGV.Columns["DataQuantificazione"].Visible = false;
            DettagliGV.Columns["DataDeterminazione"].Visible = false;
            DettagliGV.Columns["TecnicoPreparativa"].Visible = false;
            DettagliGV.Columns["DataDeterminazione"].Visible = false;
            DettagliGV.Columns["TecnicoDeterminazione"].Visible = false;
            DettagliGV.Columns["RiferimentoLocked"].Visible = false;
            DettagliGV.Columns["RiferimentoDeperibile"].Visible = false;
            DettagliGV.Columns["RiferimentoAccredia"].Visible = false;
            DettagliGV.Columns["NumeroAccredia"].Visible = false;
            DettagliGV.Columns["RdPInNumero"].Visible = false;
            DettagliGV.Columns["Stato"].Visible = false;
            DettagliGV.Columns["SoloIntestazione"].Visible = false;
            DettagliGV.Columns["Strumento"].Visible = false;
            DettagliGV.Columns["figlioDi"].Visible = false;
            DettagliGV.Columns["idParametro"].Visible = false;
            DettagliGV.Columns["CodiceFamiglia"].Visible = false;

            DettagliGV.Columns["Accettazione"].Width = 75;
            DettagliGV.Columns["Preparativa"].Width = 75;
            DettagliGV.Columns["Determinazione"].Width = 90;
            DettagliGV.Columns["Quantificazione"].Width = 90;
            DettagliGV.Columns["Stato"].Width = 75;
            DettagliGV.Columns["DataAnalisi"].Width = 75;
            DettagliGV.Columns["Firma"].Width = 105;
            DettagliGV.Columns["Deperibile"].Width = 60;
            DettagliGV.Columns["LimiteA"].Width = 60;
            DettagliGV.Columns["LimiteB"].Width = 60;


            DettagliGV.Columns["ScadenzaAnalisi"].HeaderText = "Scadenza deperibili";
            DettagliGV.Columns["DataAnalisi"].HeaderText = "Data di analisi";
            DettagliGV.Columns["AssegnatoA"].HeaderText = "Assegnato a...";
            DettagliGV.Columns["StatoCampione"].HeaderText = "Stato del campione";
            DettagliGV.Columns["DataArrivo"].HeaderText = "Data di arrivo";
            DettagliGV.Columns["RapportoDiProva"].HeaderText = "Rapporto di prova";

            DettagliGV.Columns["Firma"].DefaultCellStyle = StileFirma;

            
        }

        private void CambiaPasswordBT_Click(object sender, EventArgs e)
        {
            MascheraUtenti MascheraUtente1 = new MascheraUtenti();
            MascheraUtente1.modificaAbilitata = false;
            MascheraUtente1.ShowDialog();
        }

        public void SalvaDatiBT_Click(object sender, EventArgs e)
        {
            try
            {
                Renfresca = new MySqlCommandBuilder(Adattatore);
                Adattatore.Update(cassapanca, "caricolavoro");
            }

            catch
            {
                foreach (DataRow riga in cassapanca.Tables["caricolavoro"].Rows)
                {
                    if (riga.HasErrors)
                    {
                        MessageBox.Show("dataRow"[0] + "\n" + riga.RowError);
                    }
                }
            }

            dettagliRenfresca = new MySqlCommandBuilder(adattatoreDettagli);
            adattatoreDettagli.Update(cassapanca, "dettagli");

            if (DynamicRB.Checked == true)
            {
                FileLog.WriteLine("Le modifiche sono state salvate da " + UtenteAttivo.Nome + " " + UtenteAttivo.Cognome + " in data " + DateTime.Now.ToString());
            }
            else
            {
                FileLog.WriteLine("Le modifiche sono state salvate da " + UtenteAttivo.Nome + " " + UtenteAttivo.Cognome + " in data " + DataAnalisiDP.Value.ToString());
            }
            MessageBox.Show("Dati aggiornati");

            cambiamento = false;
        }

        public void FiltriBT_Click(object sender, EventArgs e)
        {
            
            
            string FiltroScadenza = "";
            string FiltroPar = "";
            string FiltroAre = "";
            string FiltroAss = "";
            string FiltroFir = "";
            string FiltroDDA = "";
            string FiltroADA = "";
            string FiltroDNA = "";
            string FiltroANA = "";
            string FiltroDAr = "";
            string FiltroDAn = "";
            string FiltroPrp = "";
            string FiltroDet = "";
            string FiltroQnt = "";
            string FiltroAll = "";
            string FiltroLoc = "";
            string FiltroAcr = "";
            string FiltroMat = "";
            string FiltroStC = "";
            string FiltroRdP = "";
            string FiltroDRP = "";
            string FiltroRPD = "";
            string FiltroRPA = "";
            string FiltroDRD = "";
            string FiltroDRA = "";
            string FiltroSPA = "";
            string FiltroMet = "";
            string FiltroUrg = "";
            string FiltroScA = "";
            string FiltroAcc = "";



            if (ScadenzaCB.SelectedIndex == -1)
            { FiltroScadenza = ""; }
            else
            { FiltroScadenza = "and [Scadenza] = '" + (ScadenzaCB.SelectedItem as System.Data.DataRowView)["Scadenza"].ToString().Substring(0, 10) + "'"; }

            if (DalScadenzaCB.SelectedIndex == -1)
            { FiltroDDA = ""; }
            else
            { FiltroDDA = "and [Scadenza] >= '" + (DalScadenzaCB.SelectedItem as System.Data.DataRowView)["Scadenza"].ToString().Substring(0, 10) + "'"; }

            if (AlScadenzaCB.SelectedIndex == -1)
            { FiltroADA = ""; }
            else
            { FiltroADA = "and [Scadenza] <= '" + (AlScadenzaCB.SelectedItem as System.Data.DataRowView)["Scadenza"].ToString().Substring(0, 10) + "'"; }

            if (DalCB.SelectedIndex == -1)
            { FiltroDNA = ""; }
            else
            { FiltroDNA = "and [Accettazione] >='" + DalCB.Text + "'"; }

            if (AlCB.SelectedIndex == -1)
            { FiltroANA = ""; }
            else
            { FiltroANA = "and [Accettazione] <='" + AlCB.Text + "'"; }

            if (DataArrivoCB.SelectedIndex == -1 || (DataArrivoCB.SelectedItem as System.Data.DataRowView)["DataArrivo"].ToString() == "")
            { FiltroDAr = ""; }
            else
            { FiltroDAr = "and [DataArrivo] = '" + (DataArrivoCB.SelectedItem as System.Data.DataRowView)["DataArrivo"].ToString().Substring(0, 10) + "'"; }

            if (DataAnalisiCB.SelectedIndex == -1 || (DataAnalisiCB.SelectedItem as System.Data.DataRowView)["DataAnalisi"].ToString() == "")
            { FiltroDAn = ""; }
            else
            { FiltroDAn = "and [DataAnalisi] = '" + (DataAnalisiCB.SelectedItem as System.Data.DataRowView)["DataAnalisi"].ToString() + "'"; }

            if (PreparativaCB.SelectedIndex == -1 && PreparativaCB.Text == "")
            { FiltroPrp = ""; }
            else
            { FiltroPrp = "and [Preparativa] like '%" + PreparativaCB.Text + "%'"; }

            if (DeterminazioneCB.SelectedIndex == -1 && DeterminazioneCB.Text == "")
            { FiltroDet = ""; }
            else
            { FiltroDet = "and [Determinazione] like '%" + DeterminazioneCB.Text + "%'"; }

            if (QuantificazioneCB.SelectedIndex == -1 && QuantificazioneCB.Text == "")
            { FiltroQnt = ""; }
            else
            { FiltroQnt = "and [Quantificazione] like '%" + QuantificazioneCB.Text + "%'"; }

            if (MatriceCB.SelectedIndex == -1 && MatriceCB.Text == "")
            { FiltroMat = ""; }
            else if (!(bool)reverse["Matrice"])
            { FiltroMat = "and [Matrice] like '%" + MatriceCB.Text + "%'"; }
            else
            { FiltroMat = "and [Matrice] <> '" + MatriceCB.Text + "'"; }

            if (DeperibileCB.SelectedIndex == -1 && DeperibileCB.Text == "")
            { FiltroAll = ""; }
            else if (DeperibileCB.SelectedIndex == 0)
            { FiltroAll = "and [RiferimentoDeperibile] = " + 1; }
            else if (DeperibileCB.SelectedIndex == 1)
            { FiltroAll = "and [RiferimentoDeperibile] = " + 0; }

            if (LockedCB.SelectedIndex == -1 && LockedCB.Text == "")
            { FiltroLoc = ""; }
            else if (LockedCB.Text == "Esclusivo")
            { FiltroLoc = "and [RiferimentoLocked] = " + 1; }
            else if (LockedCB.Text == "Non esclusivo")
            { FiltroLoc = "and [RiferimentoLocked] = " + 0; }

            if (AccrediaCB.SelectedIndex == -1 && AccrediaCB.Text == "")
            { FiltroAcr = ""; }
            else if (AccrediaCB.SelectedIndex == 0)
            { FiltroAcr = "and [RiferimentoAccredia] = " + 1; }
            else if (AccrediaCB.SelectedIndex == 1)
            { FiltroAcr = "and [RiferimentoAccredia] = " + 0; }

            if (StatoCampioneCB.SelectedIndex == -1 && StatoCampioneCB.Text == "")
            { FiltroStC = ""; }
            else if (!(bool)reverse["StatoCampione"])
            { FiltroStC = "and [StatoCampione] like '%" + StatoCampioneCB.Text + "%'"; }
            else
            { FiltroStC = "and [StatoCampione] <> '" + StatoCampioneCB.Text + "'"; }

            if (RapportoDiProvaCB.SelectedIndex == -1 && RapportoDiProvaCB.Text == "")
            { FiltroRdP = ""; }
            else if(!(bool)reverse["RapportoDiProva"])
            { FiltroDet = "and [RapportoDiProva] like '%" + RapportoDiProvaCB.Text + "%'"; }
            else
            { FiltroDet = "and [RapportoDiProva] <> '" + RapportoDiProvaCB.Text + "'"; }

            if (DataRdPCB.SelectedIndex == -1 || (DataRdPCB.SelectedItem as System.Data.DataRowView)["DataRdP"].ToString() == "")
            { FiltroDRP = ""; }
            else
            { FiltroDRP = "and [DataRdP] = '" + (DataRdPCB.SelectedItem as System.Data.DataRowView)["DataRdP"].ToString().Substring(0, 10) + "'"; }

            if (DalDataRdPCB.SelectedIndex == -1 || (DalDataRdPCB.SelectedItem as System.Data.DataRowView)["DataRdP"].ToString() == "")
            { FiltroDRD = ""; }
            else
            { FiltroDRD = "and [DataRdP] >= '" + (DalDataRdPCB.SelectedItem as System.Data.DataRowView)["DataRdP"].ToString().Substring(0, 10) + "'"; }

            if (AlDataRdPCB.SelectedIndex == -1 || (AlDataRdPCB.SelectedItem as System.Data.DataRowView)["DataRdP"].ToString() == "")
            { FiltroDRA = ""; }
            else
            { FiltroDRA = "and [DataRdP] <= '" + (AlDataRdPCB.SelectedItem as System.Data.DataRowView)["DataRdP"].ToString().Substring(0, 10) + "'"; }

            if (DalRdPCB.SelectedIndex == -1 && DalRdPCB.Text == "")
            { FiltroRPD = ""; }
            else
            { FiltroRPD = "and [RapportoDiProva] >= '" + DalRdPCB.Text + "'"; }

            if (AlRdPCB.SelectedIndex == -1 && AlRdPCB.Text == "")
            { FiltroRPA = ""; }
            else
            { FiltroRPA = "and [RapportoDiProva] <= '" + AlRdPCB.Text + "'"; }

            if (StatoParametroCB.SelectedIndex == -1 && StatoParametroCB.Text == "")
            { FiltroSPA = ""; }
            else
            { FiltroSPA = "[StatoParametro] like '%" + StatoParametroCB.Text + "%'"; }

            if (MetodoCB.SelectedIndex == -1 && MetodoCB.Text == "")
            { FiltroMet = ""; }
            else if(!(bool)reverse["Metodo"])
            { FiltroMet = "and [Metodo] like '%" + MetodoCB.Text + "%'"; }
            else
            { FiltroMet = "and [Metodo] <> '" + MetodoCB.Text + "'"; }

            if (UrgenzaCB.SelectedIndex == -1 && UrgenzaCB.Text == "")
            { FiltroUrg = ""; }
            else if (UrgenzaCB.SelectedIndex == 0)
            { FiltroUrg = "and [Urgenza] is not null"; }
            else if(UrgenzaCB.SelectedIndex == 1)
            { FiltroUrg = "and [Urgenza] is null"; }

            if (ScadenzaAnalisiCB.SelectedIndex == -1 || (ScadenzaAnalisiCB.SelectedItem as System.Data.DataRowView)["ScadenzaAnalisi"].ToString() == "")
            { FiltroScA = ""; }
            else
            { FiltroScA = "and [ScadenzaAnalisi] = '" + (ScadenzaAnalisiCB.SelectedItem as System.Data.DataRowView)["ScadenzaAnalisi"].ToString() + "'"; }

            if (ParametroCB.SelectedIndex == -1 && ParametroCB.Text == "")
            { FiltroPar = ""; }
            else if (!(bool)reverse["Parametro"])
            { FiltroPar = "and [Parametro] like '%" + ParametroCB.Text + "%'"; }
            else
            { FiltroPar = "and [Parametro] <> '" + ParametroCB.Text + "'"; }

            if (AreaCB.SelectedIndex == -1 && AreaCB.Text == "")
            { FiltroAre = ""; }
            else if (!(bool)reverse["Area"])
            { FiltroAre = "and [Area] like '%" + AreaCB.Text + "%'"; }
            else
            { FiltroAre = "and [Area] <> '" + AreaCB.Text + "'"; }

            if (AssegnatoACB.SelectedIndex == -1 && AssegnatoACB.Text == "")
            { FiltroAss = ""; }
            else if (!(bool)reverse["AssegnatoA"])
            { FiltroAss = "and [AssegnatoA] like '%" + AssegnatoACB.Text + "%'"; }
            else
            { FiltroAss = "and [AssegnatoA] <> '" + AssegnatoACB.Text + "'"; }

            if (FirmaCB.SelectedIndex == -1 && FirmaCB.Text == "")
            { FiltroFir = ""; }
            else if (!(bool)reverse["Firma"])
            { FiltroFir = "and [Firma] like '%" + FirmaCB.Text + "%'"; }
            else
            { FiltroFir = "and [Firma] <> '" + FirmaCB.Text + "'"; }

            if (AccettazioneCB.SelectedIndex == -1 && AccettazioneCB.Text == "")
            { FiltroAcc = ""; }
            else if (!(bool)reverse["Accettazione"])
            { FiltroAcc = "and [Accettazione] like '%" + AccettazioneCB.Text + "%'"; }
            else
            { FiltroAcc = "and [Accettazione] <> '" + AccettazioneCB.Text + "'"; }

            string controllo = FiltroAcc +
                    FiltroPar + FiltroAre + FiltroAss + FiltroFir +
                    FiltroScadenza + FiltroDDA + FiltroADA + FiltroDNA + FiltroANA + FiltroDAr + FiltroDAn + FiltroPrp +
                    FiltroDet + FiltroQnt + FiltroAll + FiltroLoc + FiltroMat + FiltroAcr + FiltroStC + FiltroRdP + FiltroDRP +
                    FiltroDRD + FiltroDRA + FiltroRPD + FiltroRPA + FiltroSPA + FiltroMet + FiltroUrg + FiltroScA;

            if (controllo != "" && Sorgente != null)
            {
                    Sorgente.Filter = FiltroSPA + FiltroAcc +
                        FiltroPar + FiltroAre + FiltroAss + FiltroFir +
                        FiltroScadenza + FiltroDDA + FiltroADA + FiltroDNA + FiltroANA + FiltroDAr + FiltroDAn + FiltroPrp +
                        FiltroDet + FiltroQnt + FiltroAll + FiltroLoc + FiltroMat + FiltroAcr + FiltroStC + FiltroRdP + FiltroDRP +
                        FiltroDRD + FiltroDRA + FiltroRPD + FiltroRPA + FiltroMet + FiltroUrg + FiltroScA;
            }
        }

        private void SpuntaPrepBT_Click(object sender, EventArgs e)
        {
            if (CompletaCB.Checked == true)
            {
                int NumeroSelezioneC = ToDoListGV.Rows.Count;

                for (int i = NumeroSelezioneC - 1; i > -1; i--)
                {
                    DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(24, ToDoListGV.Rows[i].Index);
                    ToDoListGV_CellClick(sender, huaweii);
                }
            }
            else
            {
                int NumeroSelezioneS = ToDoListGV.SelectedRows.Count;
                if (NumeroSelezioneS == 0)
                {
                    MessageBox.Show("Non è stata selezionata nessuna riga.");
                }
                else
                {
                    for (int i = NumeroSelezioneS - 1; i > -1; i--)
                    {
                        DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(24, ToDoListGV.SelectedRows[i].Index);
                        ToDoListGV_CellClick(sender, huaweii);
                    }
                }
            }
            ToDoListGV.Select();
        }
        
        private void SpuntaDetBT_Click(object sender, EventArgs e)
        {
            DateTime DataLavoro;

            if (DynamicRB.Checked == true)
            { DataLavoro = DateTime.Now; }
            else
            { DataLavoro = DataAnalisiDP.Value; }

            if (CompletaCB.Checked == true)
            {
                int NumeroSelezioneC = ToDoListGV.Rows.Count;

                for (int i = NumeroSelezioneC - 1; i > -1; i--)
                {
                    DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(25, ToDoListGV.Rows[i].Index);
                    ToDoListGV_CellClick(sender, huaweii);
                }
            }
            else
            {
                int NumeroSelezioneS = ToDoListGV.SelectedRows.Count;

                if (NumeroSelezioneS == 0)
                {
                    MessageBox.Show("Non è stata selezionata nessuna riga.");
                }
                else
                {
                    for (int i = NumeroSelezioneS - 1; i > -1; i--)
                    {
                        DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(25, ToDoListGV.SelectedRows[i].Index);
                        ToDoListGV_CellClick(sender, huaweii);
                    }
                }
            }
            ToDoListGV.Select();
        }

        private void SpuntaQuantBT_Click(object sender, EventArgs e)
        {
            DateTime DataLavoro;

            if (DynamicRB.Checked == true)
            { DataLavoro = DateTime.Now; }
            else
            { DataLavoro = DataAnalisiDP.Value; }

            try
            {
                if (CompletaCB.Checked == true)
                {
                    int NumeroSelezioneC = ToDoListGV.Rows.Count;

                    for (int i = NumeroSelezioneC - 1; i > -1; i--)
                    {
                        DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(8, ToDoListGV.Rows[i].Index);
                        ToDoListGV_CellClick(sender, huaweii);
                    }
                }
                else
                {
                    int NumeroSelezioneS = ToDoListGV.SelectedRows.Count;

                    if (NumeroSelezioneS == 0)
                    {
                        MessageBox.Show("Non è stata selezionata nessuna riga.");
                    }
                    else
                    {
                        
                        for (int i = NumeroSelezioneS - 1; i > -1; i--)
                        {
                            DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(8, ToDoListGV.SelectedRows[i].Index);
                            ToDoListGV_CellClick(sender, huaweii);
                        }
                    }
                }
                ToDoListGV.Select();
            }
            catch
            {
                MessageBox.Show("Impossibile terminare l'operazione. Probabile conflitto tra i filtri selezionati. Il programma cerrà chiuso senza salvare.");

            }
        }
        
        public void TogliFiltriBT_Click(object sender, EventArgs e)
        {
            //Sorgente.Filter = "[StatoParametro] ='accettato'" ;
            
            foreach (Control ComboFiltro in FiltriGB.Panel.Controls)
            {
                if (ComboFiltro is ComponentFactory.Krypton.Toolkit.KryptonComboBox)
                {
                    (ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex = -1;
                }
                StatoParametroCB.SelectedIndex = 0;
            }

            foreach (Control ComboFiltro in FiltriGB.Panel.Controls)
            {
                if (ComboFiltro is ComponentFactory.Krypton.Toolkit.KryptonComboBox)
                {
                    (ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex = -1;
                }

                StatoParametroCB.SelectedIndex = 0;
            }
            
            //LoadComboFilter(sender, e);
            FiltriBT_Click(sender, e);
            ToDoListGV.FirstDisplayedScrollingRowIndex = ToDoListGV.Rows.Count - 1;
        }

        private void LoadFiltroBT_Click(object sender, EventArgs e)
        {
            int numeroSlotFS = 0;
            
            Georgia.Open();
            MySqlCommand qualeSlot = new MySqlCommand("select SlotDefaultSemplici from personale where ID = " + UtenteAttivo.ID, Georgia);
            MySqlDataReader WDTVqst = qualeSlot.ExecuteReader();
            while (WDTVqst.Read())
            {
                numeroSlotFS = (int)WDTVqst[0];
            }
            Georgia.Close();
            Georgia.Open();
            foreach (Control ComboFiltro in FiltriGB.Panel.Controls)
            {
                if (ComboFiltro is KryptonComboBox)
                {
                   
                    string Colonna = ComboFiltro.Name.Substring(0, ComboFiltro.Name.Length - 1);
                    MySqlCommand loadFilter = new MySqlCommand("select " + Colonna + ", NomeSlotSemplici from config" + numeroSlotFS + " where IDUtente = " + UtenteAttivo.ID, Georgia);
                    MySqlDataReader WDTVLF;
                    WDTVLF = loadFilter.ExecuteReader();
                    while (WDTVLF.Read())
                    {
                        if ((string)WDTVLF[1] != "Vuoto")
                        {
                            if (WDTVLF[0] != DBNull.Value && (string)WDTVLF[0] != "")
                            { ComboFiltro.Text = (string)WDTVLF[0]; }
                        }
                        else
                        {
                            MessageBox.Show("Attenzione! Lo slot impostato di default è vuoto!");
                            Georgia.Close();
                            return;
                        }
                    }
                    WDTVLF.Close();
                }
            }
            Georgia.Close();
            Georgia.Open();
            foreach (KryptonButton bottoneRv in panel2.Controls)
            {
                string Colonna = bottoneRv.AccessibleName + "Rv";
                MySqlCommand loadFilterRv = new MySqlCommand
                    ("select " + Colonna + " from config" + numeroSlotFS + " where IDUtente = " + UtenteAttivo.ID, Georgia);
                MySqlDataReader WDTVRF = loadFilterRv.ExecuteReader();
                while (WDTVRF.Read())
                {
                    reverse[bottoneRv.AccessibleName] = WDTVRF[0];
                    if ((bool)WDTVRF[0] == false)
                    {
                        bottoneRv.Values.Image = global::Cornelio.Properties.Resources.database_accept_icon;
                    }
                    else
                    {
                        bottoneRv.Values.Image = global::Cornelio.Properties.Resources.http_status_not_found_icon;
                    }
                }
                WDTVRF.Close();
            }
            Georgia.Close();
            FiltriBT_Click(sender, e);
            MessageBox.Show("filtri caricati");
        }

        private void FinitoBT_Click(object sender, EventArgs e)
        {
            Login Login1 = new Login();
            Login1.ShowDialog();

        }

        private void ToDoListGV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!VistaCampioni)
            {
                if (e.ColumnIndex != 34 && e.RowIndex != -1) // annullato
                {

                    if (ToDoListGV["StatoParametro", e.RowIndex].Value.ToString() != "annullato")
                    {
                        if ((bool)ToDoListGV["SoloIntestazione", e.RowIndex].Value == false)
                        {
                            DateTime DataLavoro;

                            if (DynamicRB.Checked == true)
                            { DataLavoro = DateTime.Now; }
                            else
                            { DataLavoro = DataAnalisiDP.Value; }

                            if (ToDoListGV["ScadenzaAnalisi", e.RowIndex].Value != DBNull.Value)
                            {
                                DataScadenza = (DateTime)ToDoListGV["ScadenzaAnalisi", e.RowIndex].Value;
                            }
                            else
                            {
                                DataScadenza = null;
                            }

                            // timer
                            #region
                            if (e.ColumnIndex == 17 && e.RowIndex != -1) // timer
                            {
                                CountDownTimer.Stop();
                                provaScadenza = e.RowIndex;

                                if (ToDoListGV["ScadenzaAnalisi", provaScadenza].Value != DBNull.Value)
                                {
                                    DateTime ScaAnalisi = (DateTime)ToDoListGV["ScadenzaAnalisi", provaScadenza].Value;
                                    TimeSpan resto = ScaAnalisi.Subtract(DateTime.Now);
                                    long secondi = Convert.ToInt64(resto.TotalSeconds);

                                    if (DateTime.Now < DataScadenza && DataScadenza != null)
                                    {
                                        CountDownTimer.Start();
                                    }

                                    else if (DateTime.Now > DataScadenza)
                                    {
                                        ScadenzaAnalisiLB.Text = "Scaduto!";
                                    }
                                }
                                else if (ToDoListGV["ScadenzaAnalisi", provaScadenza].Value == DBNull.Value)
                                {
                                    ScadenzaAnalisiLB.Text = "00:00:00";
                                }
                            }
                            #endregion 

                            // Preparativa
                            #region
                            if (e.ColumnIndex == 24 && e.RowIndex != -1) 
                            {
                                cambiamento = true;
                                if (ToDoListGV["Preparativa", e.RowIndex].Value.ToString() == "--")
                                {

                                    if (DataLavoro < DataScadenza || DataScadenza == null)
                                    {
                                        ToDoListGV["DataPreparativa", e.RowIndex].Value = DataLavoro;
                                        ToDoListGV["TecnicoPreparativa", e.RowIndex].Value = UtenteAttivo.Nome + " " + UtenteAttivo.Cognome;
                                        ToDoListGV["Preparativa", e.RowIndex].Value = "Completato";
                                    }

                                    else
                                    {
                                        pretrattamentoRow = e.RowIndex;
                                        Pretrattamento Pretrattamento1 = new Pretrattamento();
                                        Pretrattamento1.Height = 237;
                                        Pretrattamento1.ShowDialog();
                                    }

                                }

                                else if (ToDoListGV["Preparativa", e.RowIndex].Value.ToString() == "Completato"
                                    && ToDoListGV["Determinazione", e.RowIndex].Value.ToString() == "--"
                                    && (ToDoListGV["TecnicoPreparativa", e.RowIndex].Value.ToString() == UtenteAttivo.Nome + " " + UtenteAttivo.Cognome
                                    || UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio"))
                                {
                                    ToDoListGV["DataPreparativa", e.RowIndex].Value = DBNull.Value;
                                    ToDoListGV["TecnicoPreparativa", e.RowIndex].Value = "--";
                                    ToDoListGV["Preparativa", e.RowIndex].Value = "--";
                                    ToDoListGV.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                                }

                                if ((ToDoListGV["CodiceFamiglia", e.RowIndex].Value as string).Length == 3)
                                {

                                    int NumeroSelezioneC = DettagliGV.Rows.Count;

                                    for (int i = NumeroSelezioneC - 1; i > -1; i--)
                                    {
                                        DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(23, DettagliGV.Rows[i].Index);
                                        DettagliGV_CellClick(sender, huaweii);
                                    }
                                }
                            }
                            #endregion

                            // Determinazione
                            #region
                            if (e.ColumnIndex == 25 && e.RowIndex != -1) 
                            {
                                cambiamento = true;
                                if (ToDoListGV["Determinazione", e.RowIndex].Value.ToString() == "--")
                                {

                                    if (ToDoListGV["Preparativa", e.RowIndex].Value.ToString() == "--")
                                    {
                                        DataGridViewCellEventArgs ePrep = new DataGridViewCellEventArgs(24, e.RowIndex);
                                        ToDoListGV_CellClick(sender, ePrep);
                                    }

                                    ToDoListGV["DataDeterminazione", e.RowIndex].Value = DataLavoro;
                                    ToDoListGV["TecnicoDeterminazione", e.RowIndex].Value = UtenteAttivo.Nome + " " + UtenteAttivo.Cognome;
                                    ToDoListGV["Determinazione", e.RowIndex].Value = "Completato";
                                }

                                else if (ToDoListGV["Determinazione", e.RowIndex].Value.ToString() == "Completato"
                                    && ToDoListGV["Quantificazione", e.RowIndex].Value.ToString() == "--"
                                    && (ToDoListGV["TecnicoDeterminazione", e.RowIndex].Value.ToString() == UtenteAttivo.Nome + " " + UtenteAttivo.Cognome ||
                                    UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio"))
                                {
                                    ToDoListGV["DataDeterminazione", e.RowIndex].Value = DBNull.Value;
                                    ToDoListGV["TecnicoDeterminazione", e.RowIndex].Value = "--";
                                    ToDoListGV["Determinazione", e.RowIndex].Value = "--";
                                }

                                if ((ToDoListGV["CodiceFamiglia", e.RowIndex].Value as string).Length == 3)
                                {

                                    int NumeroSelezioneC = DettagliGV.Rows.Count;

                                    for (int i = NumeroSelezioneC - 1; i > -1; i--)
                                    {
                                        DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(25, DettagliGV.Rows[i].Index);
                                        DettagliGV_CellClick(sender, huaweii);
                                    }
                                }
                            }
                            #endregion

                            // Risultato

                            #region
                            //if (e.ColumnIndex == 6 && e.RowIndex != -1 && (UtenteAttivo.Nome != "Pasquale" && UtenteAttivo.Nome != "Adamo" && UtenteAttivo.Nome != "Marcello"))
                            //{
                            //    cambiamento = true;

                            //    #region 
                            //    if (ToDoListGV["RiferimentoLocked", e.RowIndex].Value.ToString() == "0" || ToDoListGV["AssegnatoA", e.RowIndex].Value.ToString() == UtenteAttivo.Cognome)
                            //    {
                            //        string AccettazioneChamp = ToDoListGV["Accettazione", e.RowIndex].Value.ToString();

                            //        if (ToDoListGV["RiferimentoAccredia", e.RowIndex].Value.ToString() == "0" || UtenteAttivo.ProveAccreditate[(int)(ToDoListGV["NumeroAccredia", e.RowIndex].Value)] == 1)
                            //        {
                            //            if (ToDoListGV["Quantificazione", e.RowIndex].Value.ToString() == "Completato")
                            //            {

                            //                string accettazione = ((string)ToDoListGV["Accettazione", ToDoListGV.CurrentRow.Index].Value).Substring(0, 4);
                            //                int foglioCalcolo = ((DateTime)ToDoListGV["DataArrivo", ToDoListGV.CurrentRow.Index].Value).Month;
                            //                string meseCalcolo = string.Format("{0:00}", foglioCalcolo);
                            //                string foglioCalc = ToDoListGV["FoglioCalcolo", e.RowIndex].Value as string;
                            //                string start = ToDoListGV["Start", e.RowIndex].Value as string;
                            //                string RifC = ToDoListGV["RifC", e.RowIndex].Value as string;
                            //                string Blocco = ToDoListGV["Blocco", e.RowIndex].Value as string;
                            //                string RifL = ToDoListGV["RifL", e.RowIndex].Value as string;

                            //                if (foglioCalc != null)
                            //                {
                            //                    Excel.Application foglioCalcoloAA = new Excel.Application();
                            //                    Excel.Workbook foglioCalcoloWBA = foglioCalcoloAA.Workbooks.Open(@"\\Server\DATI\Gestione\Calcoli\Calcoli\Archivio\" + meseCalcolo + "\\"
                            //                        + accettazione + "-16" + "\\" + accettazione + "-16.xls");
                            //                    Excel.Worksheet foglioCalcoloSA = foglioCalcoloWBA.Sheets[accettazione + " 16 " + foglioCalc];

                            //                    string risultato = "";
                            //                    string numerico = "";
                            //                    decimal decimoMeridio = Convert.ToDecimal(foglioCalcoloSA.Cells[22, 7].Value);

                            //                    Inizializzazione variabili
                            //                    int LinioStart = 1;
                            //                    int LinioStop = 1;
                            //                    int ColunnaBlocco = 1;
                            //                    int ColunnaStart = 1;

                            //                    Serie di While per stabilire il range di valori che andrà copiato
                            //                    while ((foglioCalcoloSA.Cells[LinioStart, 1].Value as string) != start)
                            //                    {
                            //                        LinioStart++;
                            //                        if (LinioStart > 100) // interrompe il sottoprogramma se non trova il riferimento nel foglio excel
                            //                        {
                            //                            MessageBox.Show("Sono stati riscontrati dei problemi nel trovare la parola chiave Start. Controllare.");
                            //                            return;
                            //                        }
                            //                    }

                            //                    while ((foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value) != RifC)
                            //                    {
                            //                        ColunnaStart++;
                            //                        if (ColunnaStart > 155)
                            //                        {
                            //                            MessageBox.Show("Sono stati riscontrati dei problemi nel trovare la parola chiave RifC. Controllare.");
                            //                            return;
                            //                        }
                            //                    }

                            //                    while ((foglioCalcoloSA.Cells[LinioStart, ColunnaBlocco].Value) != Blocco)
                            //                    {
                            //                        ColunnaBlocco++;
                            //                        if (ColunnaBlocco > 155)
                            //                        {
                            //                            MessageBox.Show("Sono stati riscontrati dei problemi nel trovare la parola chiave Blocco. Controllare.");
                            //                            return;
                            //                        }
                            //                    }

                            //                    while ((foglioCalcoloSA.Cells[LinioStart, ColunnaBlocco].Value) != RifL)
                            //                    {
                            //                        LinioStart++;
                            //                        if (LinioStart > 200)
                            //                        {
                            //                            MessageBox.Show("Sono stati riscontrati dei problemi nel trovare la parola chiave RifL. Controllare.");
                            //                            return;
                            //                        }
                            //                    }



                            //                    if (foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value < 1)
                            //                    {
                            //                        numerico = foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value.ToString("G2");
                            //                        if (haiTagliato(numerico))
                            //                        {
                            //                            numerico = numerico + "0";
                            //                        }
                            //                    }

                            //                    else if (foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value >= 1 && foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value < 100)
                            //                    {
                            //                        numerico = foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value.ToString("G2");
                            //                    }

                            //                    else if (foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value >= 100 && foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value < 1000)
                            //                    {
                            //                        numerico = foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value.ToString("G3");
                            //                    }

                            //                    else if (foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value >= 1000)
                            //                    {
                            //                        numerico = foglioCalcoloSA.Cells[LinioStart, ColunnaStart].Value.ToString("G4");
                            //                    }

                            //                    if (foglioCalcoloSA.Cells[LinioStart, ColunnaStart - 1].Value != "")
                            //                    {
                            //                        risultato = (string)foglioCalcoloSA.Cells[LinioStart, ColunnaStart - 1].Value + " " + numerico;
                            //                    }
                            //                    else
                            //                    {
                            //                        risultato = numerico;
                            //                    }

                            //                    ToDoListGV["Risultato", e.RowIndex].Value = risultato;

                            //                    foglioCalcoloWBA.Close();
                            //                    foglioCalcoloAA.Quit();
                            //                }
                            //            }

                            //            else if (ToDoListGV["Quantificazione", e.RowIndex].Value.ToString() != "Completato")
                            //            {
                            //                MessageBox.Show("L'analisi deve essere quantificata");
                            //            }
                            //        }
                            //        else
                            //        {
                            //            MessageBox.Show("Utente non abilitato alla prova.");
                            //        }
                            //    }
                            //    else
                            //    {
                            //        MessageBox.Show("L'analisi deve essere quantificata obligatoriamente da " + ToDoListGV["AssegnatoA", e.RowIndex].Value.ToString());
                            //    }
                            //    #endregion

                            //    ToDoListGV.EditMode = DataGridViewEditMode.EditProgrammatically;
                            //}
                            #endregion

                            // Quantificazione
                            #region
                            if (e.ColumnIndex == 8 && e.RowIndex != -1) 
                            {
                                cambiamento = true;
                                if (ToDoListGV["RiferimentoLocked", e.RowIndex].Value.ToString() == "0" || ToDoListGV["AssegnatoA", e.RowIndex].Value.ToString() == UtenteAttivo.Cognome)
                                {
                                    string AccettazioneChamp = ToDoListGV["Accettazione", e.RowIndex].Value.ToString();


                                    if (ToDoListGV["RiferimentoAccredia", e.RowIndex].Value.ToString() == "0" || UtenteAttivo.ProveAccreditate[(int)(ToDoListGV["NumeroAccredia", e.RowIndex].Value)] == 1)
                                    {
                                        if (ToDoListGV["Quantificazione", e.RowIndex].Value.ToString() == "--")
                                        {
                                            if (ToDoListGV["AssegnatoA", e.RowIndex].Value.ToString() == "--")
                                            {
                                                ToDoListGV["AssegnatoA", e.RowIndex].Value = UtenteAttivo.Cognome;
                                            }

                                            if (ToDoListGV["Preparativa", e.RowIndex].Value.ToString() == "--")
                                            {
                                                DataGridViewCellEventArgs ePrep = new DataGridViewCellEventArgs(24, e.RowIndex);
                                                ToDoListGV_CellClick(sender, ePrep);
                                            }


                                            if (ToDoListGV["Determinazione", e.RowIndex].Value.ToString() == "--")
                                            {
                                                DataGridViewCellEventArgs eDet = new DataGridViewCellEventArgs(25, e.RowIndex);
                                                ToDoListGV_CellClick(sender, eDet);
                                            }

                                            ToDoListGV["Stato", e.RowIndex].Value = "Completato";
                                            ToDoListGV["DataQuantificazione", e.RowIndex].Value = DataLavoro;
                                            ToDoListGV["DataAnalisi", e.RowIndex].Value = DataLavoro;
                                            ToDoListGV["Firma", e.RowIndex].Value = UtenteAttivo.Nome + " " + UtenteAttivo.Cognome;
                                            ToDoListGV["Quantificazione", e.RowIndex].Value = "Completato";

                                        }

                                        else if (ToDoListGV["Quantificazione", e.RowIndex].Value.ToString() == "Completato"
                                            && (ToDoListGV["Firma", e.RowIndex].Value.ToString() == UtenteAttivo.Nome + " " + UtenteAttivo.Cognome ||
                                            UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio"))
                                        {
                                            ToDoListGV["Stato", e.RowIndex].Value = "--";
                                            ToDoListGV["DataQuantificazione", e.RowIndex].Value = DBNull.Value;
                                            ToDoListGV["DataAnalisi", e.RowIndex].Value = DBNull.Value;
                                            ToDoListGV["Firma", e.RowIndex].Value = "--";
                                            ToDoListGV["Quantificazione", e.RowIndex].Value = "--";
                                        }

                                        if ((ToDoListGV["CodiceFamiglia", e.RowIndex].Value as string).Length == 3)
                                        {


                                            Georgia.Open();
                                            MySqlCommand figli = new MySqlCommand
                                                ("update Dettagli set Quantificazione = '" + ToDoListGV["Quantificazione", e.RowIndex].Value + "' where Accettazione = '" 
                                                + ToDoListGV["Accettazione", e.RowIndex].Value +
                                                "' and CodiceFamiglia = " + ToDoListGV["ID", e.RowIndex].Value, Georgia);
                                            figli.ExecuteNonQuery();
                                            Georgia.Close();
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Utente non abilitato alla prova.");
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("L'analisi deve essere quantificata obligatoriamente da " + ToDoListGV["AssegnatoA", e.RowIndex].Value.ToString());
                                }
                            }
#endregion

                            // assegnato a 
                            #region
                            if (e.ColumnIndex == 12 && e.RowIndex != -1) 
                            {
                                cambiamento = true;
                                if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica")
                                {
                                    if (ToDoListGV["AssegnatoA", e.RowIndex].Value.ToString() == "--")
                                    {
                                        if (SceltaTecnicoCB.SelectedIndex != -1)
                                        {
                                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                                            {
                                                DettagliGV["AssegnatoA", riga.Index].Value = (SceltaTecnicoCB.SelectedItem).ToString();
                                            }

                                            ToDoListGV["AssegnatoA", e.RowIndex].Value = (SceltaTecnicoCB.SelectedItem).ToString();
                                        }
                                        else
                                        {
                                            MessageBox.Show("Attenzione! Non è stato selezionato nessun tecnico.");
                                        }
                                    }

                                    else
                                    {
                                        foreach (DataGridViewRow riga in DettagliGV.Rows)
                                        {
                                            DettagliGV["AssegnatoA", riga.Index].Value = "--";
                                        }

                                        ToDoListGV["AssegnatoA", e.RowIndex].Value = "--";
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Utente non abilitato alla modifica");
                                }
                            }
                            #endregion

                            // Lucchetto
                            #region
                            if (e.ColumnIndex == 13 && e.RowIndex != -1) 
                            {
                                cambiamento = true;
                                if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica")
                                {
                                    if ((string)ToDoListGV["AssegnatoA", e.RowIndex].Value != "--")
                                    {
                                        int PidLocked = 0;
                                        string StatoLock = "";

                                        if ((int)ToDoListGV["RiferimentoLocked", e.RowIndex].Value == 0)
                                        {
                                            PidLocked = 1;
                                            StatoLock = "redlockicon.png";
                                        }

                                        else if ((int)ToDoListGV["RiferimentoLocked", e.RowIndex].Value == 1)
                                        {
                                            PidLocked = 0;
                                            StatoLock = "greenunlockicon.png";
                                        }

                                        byte[] Padlock = null;
                                        FileStream ScanImm = new FileStream(Path.GetDirectoryName(Application.ExecutablePath) + @"\" + StatoLock, FileMode.Open, FileAccess.Read);
                                        BinaryReader Binario = new BinaryReader(ScanImm);
                                        Padlock = Binario.ReadBytes((int)ScanImm.Length);

                                        foreach (DataGridViewRow riga in DettagliGV.Rows)
                                        {
                                            DettagliGV["RiferimentoLocked", riga.Index].Value = PidLocked;
                                            DettagliGV["Locked", riga.Index].Value = Padlock;
                                        }

                                        ToDoListGV["RiferimentoLocked", e.RowIndex].Value = PidLocked;
                                        ToDoListGV["Locked", e.RowIndex].Value = Padlock;
                                    }
                                    else
                                    {
                                        MessageBox.Show("Attenzione! La prova non è stata assegnata a nessun tecnico.");
                                    }
                                }
                            }
                            #endregion

                            // Urgenza
                            #region
                            if (e.ColumnIndex == 15 && e.RowIndex != -1) 
                            {
                                cambiamento = true;
                                if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica")
                                {
                                    ToDoListGV.EditMode = DataGridViewEditMode.EditOnEnter;
                                }
                                else
                                { MessageBox.Show("Utente non abilitato alla modifica"); }
                            }
                            #endregion

                            // Sato campione
                            #region
                            if (e.ColumnIndex == 23 && e.RowIndex != -1)
                            {
                                if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica")
                                {
                                    string nroAcce = ToDoListGV["Accettazione", e.RowIndex].Value as string;
                                    cambiamento = true;
                                    Georgia.Open();
                                    MySqlCommand cambiaStato = new MySqlCommand("update caricolavoro set StatoCampione = @StatoCampione where Accettazione = '" + nroAcce +"'", Georgia);

                                    MySqlParameter StatoCampioneP = new MySqlParameter();
                                    StatoCampioneP.Direction = ParameterDirection.Input;
                                    StatoCampioneP.DbType = DbType.String;
                                    if (ToDoListGV["StatoCampione", e.RowIndex].Value as string == "In analisi")
                                    { StatoCampioneP.Value = "Da approvare"; }
                                    else if (ToDoListGV["StatoCampione", e.RowIndex].Value as string == "Da approvare")
                                    { StatoCampioneP.Value = "Certificato"; }
                                    else
                                    { StatoCampioneP.Value = "In analisi"; }
                                    cambiaStato.Parameters.AddWithValue("@StatoCampione", StatoCampioneP.Value);
                                    Georgia.Close();
                                }
                                else
                                {
                                    KryptonMessageBox.Show("Utente non abilitato alla modifica.");
                                }
                            }


                            #endregion

                            // campo note
                            #region
                            if (e.ColumnIndex == 41 && e.RowIndex != -1)
                            {
                                if (campoNote1 == null)
                                {
                                    campoNote1 = new CampoNote();
                                }

                                campoNote1.ShowDialog();
                            }
                            #endregion
                        
                        }

                        else
                        {
                            MessageBox.Show("Il parametro è in sola lettura. Impossibile apportare modifiche.");
                        }
                    }
                    
                    else
                    {
                        MessageBox.Show("Il parametro è stato annullato. Impossibile apportare modifiche.");
                    }
                }
                
                    // annullare parametro
                #region
                else if (e.ColumnIndex == 34 && e.RowIndex != -1)
                {
                    cambiamento = true;
                    if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica" || UtenteAttivo.Qualifica == "Addetto commerciale" || UtenteAttivo.Qualifica == "Responsabile Area Microbiologia")
                    {
                        if ((string)ToDoListGV["Preparativa", e.RowIndex].Value == "--")
                        {
                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                            {
                                DettagliGV["Preparativa", riga.Index].Value = "annullato";
                            }

                            ToDoListGV["Preparativa", e.RowIndex].Value = "annullato";
                        }

                        else if ((string)ToDoListGV["Preparativa", e.RowIndex].Value == "annullato")
                        {
                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                            {
                                DettagliGV["Preparativa", riga.Index].Value = "--";
                            }

                            ToDoListGV["Preparativa", e.RowIndex].Value = "--";
                        }

                        if ((string)ToDoListGV["Determinazione", e.RowIndex].Value == "--")
                        {
                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                            {
                                DettagliGV["Determinazione", riga.Index].Value = "annullato";
                            }

                            ToDoListGV["Determinazione", e.RowIndex].Value = "annullato";
                        }

                        else if ((string)ToDoListGV["Determinazione", e.RowIndex].Value == "annullato")
                        {
                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                            {
                                DettagliGV["Determinazione", riga.Index].Value = "--";
                            }

                            ToDoListGV["Determinazione", e.RowIndex].Value = "--";
                        }

                        if ((string)ToDoListGV["Quantificazione", e.RowIndex].Value == "--")
                        {
                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                            {
                                DettagliGV["Quantificazione", riga.Index].Value = "annullato";
                            }

                            ToDoListGV["Quantificazione", e.RowIndex].Value = "annullato";
                        }

                        else if ((string)ToDoListGV["Quantificazione", e.RowIndex].Value == "annullato")
                        {
                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                            {
                                DettagliGV["Quantificazione", riga.Index].Value = "--";
                            }

                            ToDoListGV["Quantificazione", e.RowIndex].Value = "--";
                        }

                        if ((string)ToDoListGV["StatoParametro", e.RowIndex].Value == "accettato")
                        {
                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                            {
                                DettagliGV["StatoParametro", riga.Index].Value = "annullato";
                            }

                            ToDoListGV["StatoParametro", e.RowIndex].Value = "annullato";
                        }

                        else
                        {
                            foreach (DataGridViewRow riga in DettagliGV.Rows)
                            {
                                DettagliGV["StatoParametro", riga.Index].Value = "accettato";
                            }

                            ToDoListGV["StatoParametro", e.RowIndex].Value = "accettato";
                        }
                    }

                    else
                    {
                        MessageBox.Show("Non si dispongono dei permessi necessari per annullare il parametro.");
                    }

                #endregion
                }
            }
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            if (FiltriAvanzati1 == null)
                FiltriAvanzati1 = new FiltriAvanzati();

            FiltriAvanzati1.Show(); 
        }

        private void printToDoList_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bm = new Bitmap(this.ToDoListGV.Width, this.ToDoListGV.Height);
            System.Drawing.Rectangle rc = new System.Drawing.Rectangle(0, 0, this.ToDoListGV.Width, this.ToDoListGV.Height);
            ToDoListGV.DrawToBitmap(bm, rc);
            e.Graphics.DrawImage(bm, 0, 0);
        }

        public void StampaBT_Click(object sender, EventArgs e)
        {
            goodWorkPdf(ToDoListGV);
            string applicazione = "";

            Process process = new Process();
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.StartInfo.Verb = "print";
            string pdfFileName = Path.GetDirectoryName(Application.ExecutablePath) + "\\" + "fileStampa\\" + UtenteAttivo.Cognome + ".pdf";

            if (File.Exists(@"C:\Programmi\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe"))
            {
                applicazione = @"C:\Programmi\Adobe\Acrobat 11.0\Acrobat\Acrobat.exe";
            }
            else if (File.Exists(@"C:\Programmi\Adobe\Acrobat 5.0\Reader\AcroRd32.exe"))
            {
                applicazione = @"C:\Programmi\Adobe\Acrobat 5.0\Reader\AcroRd32.exe";
            }
            else if (File.Exists(@"C:\Programmi\Adobe\reader 11.0\Reader\AcroRd32.exe"))
            {
                applicazione = @"C:\Programmi\Adobe\reader 11.0\Reader\AcroRd32.exe";
            }
            else
            {
                MessageBox.Show("PdfReader non configurato. Avvisare il RTdL,");
                return;
            }

            process.StartInfo.FileName = applicazione; ;

            process.StartInfo.Arguments = @"/p /h " + pdfFileName;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.Start();
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            process.EnableRaisingEvents = true;
            process.CloseMainWindow();

            MessageBox.Show("Confermare al termine della stampa");
            process.Kill();
        }

        private void AnteprimaStampaBT_Click(object sender, EventArgs e)
        {
            goodWorkPdf(ToDoListGV);
            Process.Start(Path.GetDirectoryName(Application.ExecutablePath) + "\\" + "fileStampa\\" + UtenteAttivo.Cognome + ".pdf");
            
        }

        public void AggiornaBT_Click(object sender, EventArgs e)
        {
            if (cambiamento)
            {
                DialogResult scelta;
                scelta = MessageBox.Show("Salvare le modifiche apportate?", "OGL", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (scelta == DialogResult.Yes)
                {
                    SalvaDatiBT_Click(sender, e);
                    cambiamento = false;
                }
            }

            LoadToDoListGV();
            FiltriBT_Click(sender, e);
            try
            {
                ToDoListGV.FirstDisplayedScrollingRowIndex = ToDoListGV.Rows.Count - 1;
            }
            catch{}
        }

        private void ScegliTecnicoBT_Click(object sender, EventArgs e)
        {
            if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica")
            {

                if (SceltaTecnicoCB.SelectedIndex > -1)
                {
                    if (CompletaCB.Checked)
                    {
                        for (int i = ToDoListGV.RowCount -1 ; i >= 0; i--)
                        {
                                DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(12, ToDoListGV.Rows[i].Index);
                                ToDoListGV_CellClick(sender, huaweii);
                        }
                    }

                    else
                    {
                        int NumeroSelezioneS = ToDoListGV.SelectedRows.Count;
                        if (NumeroSelezioneS == 0)
                        {
                            MessageBox.Show("Non è stata selezionata nessuna riga.");
                        }
                        else
                        {
                            for (int i = NumeroSelezioneS - 1; i > -1; i--)
                            {
                                DataGridViewCellEventArgs huaweii = new DataGridViewCellEventArgs(12, ToDoListGV.SelectedRows[i].Index);
                                ToDoListGV_CellClick(sender, huaweii);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Attenzione! Nessun tecnico selezionato.");
                }
            }
            else
            {
                MessageBox.Show("Utente non abilitato per questa operazione");
            }
        }

        private void ParametroFT_Click(object sender, EventArgs e)
        {
            String NomeFiltro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3) + "CB";
            foreach (Control ComboFiltro in FiltriGB.Panel.Controls)
            {
                if (ComboFiltro.Name == NomeFiltro && ComboFiltro is ComponentFactory.Krypton.Toolkit.KryptonComboBox)
                {
                    (ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex = -1;
                    (ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex = -1;
                    (ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).Text = "";
                    FiltriBT_Click(sender, e);
                }
            }
        }

        private void ParametroFTPlus_Click(object sender, EventArgs e)
        {
            String NomeParametro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3);
            String NomeComboFiltro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3) + "CB";

            foreach (Control ComboFiltro in FiltriGB.Panel.Controls)
            {
                if (ComboFiltro.Name == NomeComboFiltro && ComboFiltro is ComponentFactory.Krypton.Toolkit.KryptonComboBox)
                {
                    Sorgente.Filter = "[" + NomeParametro + "] like'%" + ComboFiltro.Text + "%' and [StatoParametro] = 'Accettato'";
                }
            }
        }

        private void ParametroFTPlusData_Click(object sender, EventArgs e)
        {
            String NomeParametro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3);
            String NomeComboFiltro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3) + "CB";

            foreach (Control ComboFiltro in FiltriGB.Panel.Controls)
            {
                if (ComboFiltro.Name == NomeComboFiltro && ComboFiltro is ComponentFactory.Krypton.Toolkit.KryptonComboBox)
                {
                    if ((ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex != -1)
                    {
                        Sorgente.Filter = "[" + NomeParametro + "] ='" 
                            + ((ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedItem as System.Data.DataRowView)[NomeParametro].ToString() + "' and"+
                            " [StatoParametro] = 'Accettato'"; 
                    }
                }
            }
        }

        private void ParametroFTPlusBool_Click(object sender, EventArgs e)
        {
            String NomeParametro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3);
            String NomeComboFiltro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3) + "CB";

            foreach (Control ComboFiltro in FiltriGB.Panel.Controls)
            {
                if (ComboFiltro.Name == NomeComboFiltro && ComboFiltro is ComponentFactory.Krypton.Toolkit.KryptonComboBox)
                {
                    if ((ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex == 0)
                    {
                        Sorgente.Filter = "[Riferimento" + NomeParametro + "] =" + 1 +
                        "and [StatoParametro] = 'Accettato'"; }

                    else if ((ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex == 1)
                    {
                        Sorgente.Filter = "[Riferimento" + NomeParametro + "] =" + 0 +
                        "and [StatoParametro] = 'Accettato'";
                    }

                }
            }
        }

        private void ParametroFTPNumericNull_Click(object sender, EventArgs e)
        {
            String NomeParametro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3);
            String NomeComboFiltro = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 3) + "CB";

            foreach (Control ComboFiltro in FiltriGB.Panel.Controls)
            {
                if (ComboFiltro.Name == NomeComboFiltro && ComboFiltro is ComponentFactory.Krypton.Toolkit.KryptonComboBox)
                {
                    if ((ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex == 0)
                    {
                        Sorgente.Filter = "[" + NomeParametro + "] is not null" +
                        " and [StatoParametro] = 'Accettato'";
                    }
                    
                    else if ((ComboFiltro as ComponentFactory.Krypton.Toolkit.KryptonComboBox).SelectedIndex == 1)
                    {
                        Sorgente.Filter = "[" + NomeParametro + "] is null" +
                        " and [StatoParametro] = 'Accettato'";
                    }
                }
            }
        }

        private void LucchettoBT_Click(object sender, EventArgs e)
        {
            if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica")
            {
                if (CompletaCB.Checked == true)
                {
                    for (int i = 0; i < ToDoListGV.RowCount; i++)
                    {
                        DataGridViewCellEventArgs exvo = new DataGridViewCellEventArgs(12, i);
                        ToDoListGV_CellClick(sender, exvo);
                    }
                }
                else
                {
                    int NumeroSelezioneS = ToDoListGV.SelectedRows.Count;
                    if (NumeroSelezioneS == 0)
                    {
                        MessageBox.Show("Non è stata selezionata nessuna riga.");
                    }
                    else
                    {
                        for (int i = NumeroSelezioneS - 1; i > -1; i--)
                        {
                            DataGridViewCellEventArgs exvo = new DataGridViewCellEventArgs(12, ToDoListGV.SelectedRows[i].Index);
                            ToDoListGV_CellClick(sender, exvo);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Utente non abilitato per questa operazione");
            }
        }

        private void kryptonGroupBox1_Panel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void DettagliBT_Click(object sender, EventArgs e)
        {
            if (ballControl)
            {
                MascheraCampione DettagliCampione = new MascheraCampione();
                DettagliCampione.Show();
            }
            else
            {
                MascheraAnalita DettagliProva = new MascheraAnalita();
                DettagliProva.Show();
            }
        }

        private void ToDoListGV_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                if (DynamicRB.Checked == true || SemiStaticRB.Checked == true)
                {
                    FileLog.WriteLine("Modifica effettuta sul campione: " + ToDoListGV[1, e.RowIndex].Value + " - Parametro: " + ToDoListGV[2, e.RowIndex].Value + " ." +
                        ToDoListGV.Columns[e.ColumnIndex].Name + ": " + ToDoListGV[e.ColumnIndex, e.RowIndex].Value
                         + ". Modifica effettuata in data: " + DateTime.Now.ToString()
                        + " da " + UtenteAttivo.Nome + " " + UtenteAttivo.Cognome);
                }
                else
                {
                    FileLog.WriteLine("Modifica effettuta sul campione: " + ToDoListGV[1, e.RowIndex].Value + " - Parametro: " + ToDoListGV[2, e.RowIndex].Value + " ." +
                        ToDoListGV.Columns[e.ColumnIndex].Name + ": " + ToDoListGV[e.ColumnIndex, e.RowIndex].Value
                         + ". Modifica effettuata in data: " + DataAnalisiDP.Value.ToString()
                        + " da " + UtenteAttivo.Nome + " " + UtenteAttivo.Cognome);
                }
            }
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {

        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            if (DynamicRB.Checked == true)
            { SemiStaticRB.Checked = true; }
            else if (SemiStaticRB.Checked == true)
            { StaticRB.Checked = true; }
            else
            { DynamicRB.Checked = true; }
        }

        private void SelCompBT_Click(object sender, EventArgs e)
        {
            if (CompletaCB.Checked == true)
            { SelettivaCB.Checked = true; }
            else
            { CompletaCB.Checked = true; } 
        }

        private void StatoCampioniBT_Click(object sender, EventArgs e)
        {

        }

        private void VistaBT_Click(object sender, EventArgs e)
        {

        }

        private void kryptonPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void NascondiColonna_Click(object sender, EventArgs e)
        {
            string nomeColonna = (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Substring(0, (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Name.Length - 4);

            foreach (DataGridViewColumn Colonna in ToDoListGV.Columns)
            {
                if (Colonna.Name == nomeColonna)
                {
                    if (Colonna.Visible == true)
                    {
                        Colonna.Visible = false;
                        (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Close_2_icon;
                    }
                    else
                    {
                        Colonna.Visible = true;
                        (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Ok_iconthick;
                    }
                }
            }

            foreach (DataGridViewColumn Colonna in DettagliGV.Columns)
            {
                if (Colonna.Name == nomeColonna)
                {
                    if (Colonna.Visible == true)
                    {
                        Colonna.Visible = false;
                        (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Close_2_icon;
                    }
                    else
                    {
                        Colonna.Visible = true;
                        (sender as ComponentFactory.Krypton.Toolkit.KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Ok_iconthick;
                    }
                }
            }
        }

        private void FiltriGB_MouseDown(object sender, MouseEventArgs e)
        {
            if (Ancorato == false)
            {
                FiltriGBCliccato = true;
                FiltriGB.Dock = DockStyle.None;
                deltaX = e.X - FiltriGB.Location.X;
                deltaY = e.Y - FiltriGB.Location.Y;
                //MessageBox.Show(FiltriGBCliccato.ToString() + " " + deltaX + " " + deltaY); 
            }
        }

        private void FiltriGB_MouseMove(object sender, MouseEventArgs e)
        {
            if (FiltriGBCliccato == true)
            {
                this.FiltriPanel.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
        }

        private void FiltriGB_MouseUp(object sender, MouseEventArgs e)
        {
            if (FiltriGBCliccato == true)
            {
                this.FiltriPanel.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                FiltriGB.Location = new Point(e.X - deltaX, e.Y - deltaY);
            }
        }

        private void SalvaFilterBT_Click(object sender, EventArgs e)
        {
            if (SlotImpostazioni1 == null)
            {
                SlotImpostazioni1 = new SlotImpostazioni();
            }

            SlotImpostazioni1.ShowDialog();
        }

        private void checkAnnullatoBT_Click(object sender, EventArgs e)
        {
            AggiornaBT_Click(sender, e);
            ExcApp = ExcApp = new Excel.Application();
            OGL = ExcApp.Workbooks.Open(@"\\Server\DATI\Gestione\OGLnew.xls", null, true);
            AC = OGL.Sheets["All"];
            int Corridore = 10;

            string preparativaCheck = "Completato";
            string determinazioneCheck = "Completato";
            string quantificazioneCheck = "Completato";

            MySqlCommand CheckCarica = new MySqlCommand("select * from caricolavoro", Georgia);
            MySqlDataAdapter CheckAdattatore = new MySqlDataAdapter();
            DataTable CheckTavoletta = new DataTable();

            CheckAdattatore.SelectCommand = CheckCarica;
            CheckAdattatore.Fill(CheckTavoletta);
            CheckAdattatore.Update(CheckTavoletta);

            Georgia.Open();

            while (AC.Cells[Corridore, 30].Value != null)
            {
                int riga = Corridore - 8;
                int checkRiga = Corridore - 9;
                if (AC.Cells[Corridore, 30].Value == "annullato")
                {
                    if ((string)CheckTavoletta.Rows[checkRiga]["Preparativa"] != "Completato")
                    { preparativaCheck = "annullato"; }
                    if ((string)CheckTavoletta.Rows[checkRiga]["Determinazione"] != "Completato")
                    { determinazioneCheck = "annullato"; }
                    if ((string)CheckTavoletta.Rows[checkRiga]["Quantificazione"] != "Completato")
                    { quantificazioneCheck = "annullato"; }
                    
                    MySqlCommand annullaCm = new MySqlCommand
                        ("update caricolavoro set Preparativa = '" + preparativaCheck + "', Quantificazione = '" + quantificazioneCheck 
                        + "', Determinazione = '" + determinazioneCheck + "', StatoParametro = 'annullato' where ID = "
                        + riga, Georgia);

                    
                    annullaCm.ExecuteNonQuery();
                    
                }
                Corridore++;
            }
            Georgia.Close();
            CheckTavoletta = null;
            CheckTavoletta = null;
            OGL.Close();
            ExcApp.Quit();
        }

        private void SalvaVistaBT_Click(object sender, EventArgs e)
        {
            if (SlotImpostazioni1 == null)
            {
                SlotImpostazioni1 = new SlotImpostazioni();
            }

            SlotImpostazioni1.ShowDialog();
        }

        private void LoadVistaBT_Click(object sender, EventArgs e)
        {
            string ColonnaRes = "";
            string NomeColonna = "";
            string NomeBottone = "";
            string ColonnaOrder = "";
            int numeroSlotLV = 1;

            Georgia.Open();
            MySqlCommand qualeSlot = new MySqlCommand("select SlotDefaultVisuale from personale where ID = " + UtenteAttivo.ID, Georgia);
            MySqlDataReader WDTVqst = qualeSlot.ExecuteReader();
            while (WDTVqst.Read())
            {
                numeroSlotLV = (int)WDTVqst[0];
            }
            Georgia.Close();
            
            Georgia.Open();

            MySqlCommand LoadVista = new MySqlCommand("select * from config" + numeroSlotLV + " where ID = " + UtenteAttivo.ID, Georgia);
            MySqlDataReader WDTVLV = LoadVista.ExecuteReader();
            while (WDTVLV.Read())
            {
                if ((string)WDTVLV["NomeSlotVisuale"] != "Vuoto")
                {
                    foreach (DataGridViewColumn Colonna in ToDoListGV.Columns)
                    {
                        ColonnaRes = Colonna.Name + "R";
                        try
                        {
                            Colonna.Width = (int)WDTVLV[ColonnaRes];
                        }
                        catch { }

                        ColonnaOrder = Colonna.Name + "O";
                        try
                        {
                            Colonna.DisplayIndex = (int)WDTVLV[ColonnaOrder];
                        }
                        catch { }

                        try
                        {
                            NomeColonna = Colonna.Name + "V";
                            NomeBottone = Colonna.Name + "View";

                            foreach (Control bottone in FiltriGB.Panel.Controls)
                            {
                                if (bottone.Name == NomeBottone)
                                {
                                    if ((bool)WDTVLV[NomeColonna] == true)
                                    {
                                        (bottone as ComponentFactory.Krypton.Toolkit.KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Ok_iconthick;
                                    }
                                    else
                                    {
                                        (bottone as ComponentFactory.Krypton.Toolkit.KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Close_2_icon;
                                    }
                                }
                            }
                            Colonna.Visible = (bool)WDTVLV[NomeColonna];
                        }
                        catch { }
                    }

                    foreach (DataGridViewColumn Colonna in DettagliGV.Columns)
                    {
                        ColonnaRes = Colonna.Name + "R";
                        try
                        {
                            Colonna.Width = (int)WDTVLV[ColonnaRes];
                        }
                        catch { }

                        ColonnaOrder = Colonna.Name + "O";
                        try
                        {
                            Colonna.DisplayIndex = (int)WDTVLV[ColonnaOrder];
                        }
                        catch { }

                        try
                        {
                            NomeColonna = Colonna.Name + "V";
                            NomeBottone = Colonna.Name + "View";

                            foreach (Control bottone in FiltriGB.Panel.Controls)
                            {
                                if (bottone.Name == NomeBottone)
                                {
                                    if ((bool)WDTVLV[NomeColonna] == true)
                                    {
                                        (bottone as ComponentFactory.Krypton.Toolkit.KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Ok_iconthick;
                                    }
                                    else
                                    {
                                        (bottone as ComponentFactory.Krypton.Toolkit.KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Close_2_icon;
                                    }
                                }
                            }
                            Colonna.Visible = (bool)WDTVLV[NomeColonna];
                        }
                        catch { }
                    }

                    this.Height = (int)WDTVLV["AltezzaForm"];
                    this.Width = (int)WDTVLV["LarghezzaForm"];
                    kryptonPanel4.Width = (int)WDTVLV["DimensionPannExt"];

                    if ((bool)WDTVLV["sediciNoni"] == false)
                    {
                        MonitorBT_Click(sender, e);
                    }
                }
                else
                {
                    MessageBox.Show("Attenzione! Lo slot impostato di default è vuoto!");
                    Georgia.Close();
                    return;
                }
            }
 
            Georgia.Close();
        }

        private void LogoPB_Click(object sender, EventArgs e)
        {
            MascheraUtenti MascheraUtente1 = new MascheraUtenti();
            MascheraUtente1.modificaAbilitata = true;
            MascheraUtente1.AvantiBT.Visible = false;
            MascheraUtente1.IndietroBT.Visible = false;
            MascheraUtente1.pictureBox1.Visible = true;
            MascheraUtente1.MostraFacceBT.Visible = true;
            
            
            foreach (Control pulsante in MascheraUtente1.kryptonPanel1.Controls)
            {
                string finaleNome = pulsante.Name.Substring(pulsante.Name.Length - 2, 2);
                if (finaleNome == "MD")
                {
                    pulsante.Visible = true;
                }
            }

            foreach (Control pulsante in MascheraUtente1.kryptonPanel2.Controls)
            {
                string finaleNome = pulsante.Name.Substring(pulsante.Name.Length - 2, 2);
                if (finaleNome == "MD")
                {
                    pulsante.Visible = true;
                }
            }
   
            MascheraUtente1.ShowDialog();
        }

        private void kryptonNumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            
            if (!fattore)
            {
                SizeF fattoreR = new SizeF(0.97F, 0.97F);
                FiltriGB.Panel.Scale(fattoreR);
                label1.Location = new Point(label1.Location.X, label1.Location.Y - 10);
                fattore = true;
            }

            if (label1.Font.Size > 16)
            {

                
                foreach (Control controllo in FiltriGB.Panel.Controls)
                {
                    if (controllo.Name != "label1" && controllo.Name != "pictureBox1" && controllo.Name != ZipBT.Name)
                        controllo.Location = new Point(controllo.Location.X, controllo.Location.Y - 2);
                }

                label1.Font = new System.Drawing.Font(label1.Font.Name, label1.Font.Size - 0.5F, label1.Font.Style);
            }
        }

        public void MonitorBT_Click(object sender, EventArgs e)
        {
            if (sediciNoni)
            {
                sediciNoni = false;
                SelCompBT.Visible = false;
                CompletaCB.Location = new Point(5, CompletaCB.Location.Y);
                SelettivaCB.Location = new Point(5, SelettivaCB.Location.Y);
                BottonePirata.Visible = false;
                DynamicRB.Location = new Point(5, DynamicRB.Location.Y);
                SemiStaticRB.Location = new Point(5, SemiStaticRB.Location.Y);
                StaticRB.Location = new Point(5, StaticRB.Location.Y);
                panel5.Location = new Point(150, panel5.Location.Y);
                pictureBox2.Visible = false;
                panel3.Width = 270;
                SincronizzaOGLBT.Location = new Point(SincronizzaOGLBT.Location.X - 180, SincronizzaOGLBT.Location.Y);
                sincronizzaLB.Location = new Point(sincronizzaLB.Location.X - 180, sincronizzaLB.Location.Y);
                timerPB.Location = new Point(timerPB.Location.X - 180, timerPB.Location.Y);
                ScadenzaAnalisiLB.Location = new Point(ScadenzaAnalisiLB.Location.X - 180, ScadenzaAnalisiLB.Location.Y);
                MonitorBT.Location = new Point(MonitorBT.Location.X - 190, MonitorBT.Location.Y);
                ResizeUD.Location = new Point(ResizeUD.Location.X - 190, ResizeUD.Location.Y);
                kryptonLabel42.Location = new Point(kryptonLabel42.Location.X - 190, kryptonLabel42.Location.Y);
                panel1.Location = new Point(panel1.Location.X - 200, panel1.Location.Y);
                FiltriAvanzatiBT.Location = new Point(FiltriAvanzatiBT.Location.X - 230, FiltriAvanzatiBT.Location.Y);
                LoadVistaBT.Location = new Point(LoadVistaBT.Location.X - 230, LoadVistaBT.Location.Y);
                SalvaVistaBT.Location = new Point(SalvaVistaBT.Location.X - 230, SalvaVistaBT.Location.Y);
                VediFileLogBT.Location = new Point(VediFileLogBT.Location.X - 230, VediFileLogBT.Location.Y);
                excelBT.Location = new Point(excelBT.Location.X - 230, excelBT.Location.Y);
                wordBT.Location = new Point(wordBT.Location.X - 230, wordBT.Location.Y);

                
                DettagliBT.Location = new Point(229, 77);
                kryptonLabel31.Location = new Point(266, 82);
                AggiornaBT.Location = new Point(375, 10);
                kryptonLabel18.Location = new Point(368, 37);
                StampaBT.Location = new Point(375, 60);
                kryptonLabel19.Location = new Point(366, 88);
                SalvaDatiBT.Location = new Point(426, 10);
                kryptonLabel17.Location = new Point(424, 37);
                AnteprimaStampaBT.Location = new Point(426, 60);
                kryptonLabel20.Location = new Point(415, 88);
                MascheraUtenteBT.Location = new Point(484, 2);
                kryptonLabel40.Location = new Point(482, 38);
                RubricaBT.Location = new Point(488, 58);
                kryptonLabel41.Location = new Point(482, 90);
                StatoCampioniBT.Location = new Point(558, 14);
                kryptonLabel35.Location = new Point(528, 44);
                panelState.Location = new Point(536, 60);
                ScegliTecnicoBT.Location = new Point(ScegliTecnicoBT.Location.X - 300, ScegliTecnicoBT.Location.Y);
                SceltaTecnicoCB.Location = new Point(SceltaTecnicoCB.Location.X - 300, SceltaTecnicoCB.Location.Y);
                LucchettoBT.Location = new Point(LucchettoBT.Location.X - 300, LucchettoBT.Location.Y);


                kryptonLabel23.Location = new Point(kryptonLabel23.Location.X - 300, kryptonLabel23.Location.Y);
                kryptonPanel2.Height = 117;
                ToDoListGV.Location = new Point(ToDoListGV.Location.X, ToDoListGV.Location.Y + 20);
                ToDoListGV.Height = ToDoListGV.Height - 20;

                NomeUtenteLB.Font = new System.Drawing.Font(NomeUtenteLB.Font.Name, NomeUtenteLB.Font.Size - 6, NomeUtenteLB.Font.Style);

                this.Refresh();

            }
            else
            {
                this.MonitorBT.Location = new Point(MonitorBT.Location.X + 190, MonitorBT.Location.Y);
                this.ResizeUD.Location = new Point(ResizeUD.Location.X + 190, ResizeUD.Location.Y);
                this.kryptonLabel42.Location = new Point(kryptonLabel42.Location.X + 190, kryptonLabel42.Location.Y);
                this.timerPB.Location = new Point(timerPB.Location.X + 180, timerPB.Location.Y);
                this.ScadenzaAnalisiLB.Location = new Point (ScadenzaAnalisiLB.Location.X + 180, ScadenzaAnalisiLB.Location.Y);
                this.LoadVistaBT.Location = new Point(LoadVistaBT.Location.X + 230, LoadVistaBT.Location.Y);
                this.SalvaVistaBT.Location = new Point(SalvaVistaBT.Location.X + 230, SalvaVistaBT.Location.Y);
                this.excelBT.Location = new Point(excelBT.Location.X + 230, excelBT.Location.Y);
                this.wordBT.Location = new Point(wordBT.Location.X + 230, wordBT.Location.Y);
                this.SincronizzaOGLBT.Location = new Point(SincronizzaOGLBT.Location.X + 180, SincronizzaOGLBT.Location.Y);
                this.sincronizzaLB.Location = new Point(sincronizzaLB.Location.X + 180, sincronizzaLB.Location.Y);
                this.panel5.Location = new System.Drawing.Point(212, 17);
                this.panel5.Size = new System.Drawing.Size(186, 50);
                this.StaticRB.Location = new System.Drawing.Point(54, 32);
                this.BottonePirata.Visible = true;
                this.DynamicRB.Location = new System.Drawing.Point(54, 1);
                this.SemiStaticRB.Location = new System.Drawing.Point(54, 16);
                this.panel3.Location = new System.Drawing.Point(23, 15);
                this.panel3.Size = new System.Drawing.Size(437, 56);
                this.SelCompBT.Visible = true;
                this.SelettivaCB.Location = new System.Drawing.Point(57, 8);
                this.CompletaCB.Location = new System.Drawing.Point(57, 28);
                this.FiltriAvanzatiBT.Location = new Point(FiltriAvanzatiBT.Location.X + 230, FiltriAvanzatiBT.Location.Y);
                this.DataAnalisiDP.Location = new System.Drawing.Point(10, 26);
                this.kryptonLabel14.Location = new System.Drawing.Point(8, 1);
                this.panel1.Location = new System.Drawing.Point(728, 18);
                this.panel1.Size = new System.Drawing.Size(141, 52);
                this.VediFileLogBT.Location = new Point(VediFileLogBT.Location.X + 230, VediFileLogBT.Location.Y);


                this.kryptonPanel2.Size = new System.Drawing.Size(1043, 92);
                this.ToDoListGV.Location = new Point(ToDoListGV.Location.X, ToDoListGV.Location.Y - 20);
                this.ToDoListGV.Height = ToDoListGV.Height + 20;
                this.kryptonLabel41.Location = new System.Drawing.Point(702, 54);
                this.RubricaBT.Location = new System.Drawing.Point(710, 20);
                this.kryptonLabel40.Location = new System.Drawing.Point(757, 54);
                this.kryptonLabel35.Location = new System.Drawing.Point(852, 41);
                this.StatoCampioniBT.Location = new System.Drawing.Point(880, 11);
                this.DettagliBT.Location = new System.Drawing.Point(627, 25);
                this.kryptonLabel31.Location = new System.Drawing.Point(618, 54);
                this.LucchettoBT.Location = new System.Drawing.Point(1004, 11);
                this.kryptonLabel25.Location = new System.Drawing.Point(257, 57);
                this.SpuntaQuantBT.Location = new System.Drawing.Point(228, 56);
                this.kryptonLabel24.Location = new System.Drawing.Point(256, 36);
                this.SpuntaDetBT.Location = new System.Drawing.Point(228, 35);
                this.SceltaTecnicoCB.Location = new System.Drawing.Point(958, 46);
                this.ScegliTecnicoBT.Location = new System.Drawing.Point(966, 11);
                this.kryptonLabel23.Location = new System.Drawing.Point(960, 67);
                this.AnteprimaStampaBT.Location = new System.Drawing.Point(551, 24);
                this.kryptonLabel20.Location = new System.Drawing.Point(538, 53);
                this.StampaBT.Location = new System.Drawing.Point(495, 26);
                this.kryptonLabel19.Location = new System.Drawing.Point(487, 53);
                this.kryptonLabel18.Location = new System.Drawing.Point(363, 53);
                this.SalvaDatiBT.Location = new System.Drawing.Point(421, 26);
                this.kryptonLabel17.Location = new System.Drawing.Point(419, 53);
                this.kryptonLabel15.Location = new System.Drawing.Point(256, 16);
                this.SpuntaPrepBT.Location = new System.Drawing.Point(228, 14);
                this.NomeUtenteLB.Location = new System.Drawing.Point(90, 43);
                this.panelState.Location = new System.Drawing.Point(860, 57);
                this.AggiornaBT.Location = new System.Drawing.Point(375, 24);

                this.MascheraUtenteBT.Location = new System.Drawing.Point(756, 15);
                pictureBox2.Visible = true;

                NomeUtenteLB.Font = new System.Drawing.Font(NomeUtenteLB.Font.Name, NomeUtenteLB.Font.Size + 6, NomeUtenteLB.Font.Style);

                this.Refresh();
                sediciNoni = true;



            }

        }

        private void ZipBT_MouseDown(object sender, MouseEventArgs e)
        {
            RegolaX = e.X;
            ZipBT.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Standalone;
        }

        private void ZipBT_MouseUp(object sender, MouseEventArgs e)
        {
            ZipBT.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.LowProfile;
            RegolaX = RegolaX - e.X;
            //ToDoListGV.Width = ToDoListGV.Width - RegolaX;
            kryptonPanel4.Width = kryptonPanel4.Width - RegolaX;

        }

        private static string Encrypt(string text)
        {
            byte[] plaintextbytes = System.Text.ASCIIEncoding.ASCII.GetBytes(text);
            AesCryptoServiceProvider aes = new AesCryptoServiceProvider();
            aes.BlockSize = 128;
            aes.KeySize = 256;
            aes.Key = System.Text.ASCIIEncoding.ASCII.GetBytes(Key);
            aes.IV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV);
            aes.Padding = PaddingMode.PKCS7;
            aes.Mode = CipherMode.CBC;
            ICryptoTransform cripto = aes.CreateEncryptor(aes.Key, aes.IV);
            byte[] encrypted = cripto.TransformFinalBlock(plaintextbytes, 0, plaintextbytes.Length);
            cripto.Dispose();
            return Convert.ToBase64String(encrypted);
        }

        private static string Decrypt(string encripted)
        {
            byte[] encryptedBytes = Convert.FromBase64String(encripted);
            AesCryptoServiceProvider aes = new AesCryptoServiceProvider();
            aes.BlockSize = 128;
            aes.KeySize = 256;
            aes.Key = System.Text.ASCIIEncoding.ASCII.GetBytes(Key);
            aes.IV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV);
            aes.Padding = PaddingMode.PKCS7;
            aes.Mode = CipherMode.CBC;
            ICryptoTransform cripto = aes.CreateDecryptor(aes.Key, aes.IV);
            byte[] secret = cripto.TransformFinalBlock(encryptedBytes, 0, encryptedBytes.Length);
            cripto.Dispose();
            return System.Text.ASCIIEncoding.ASCII.GetString(secret);
        }

        public void Crittografa(string input, string output)
        {
            FileStream inputStream = new FileStream(input, FileMode.Open, FileAccess.Read);
            FileStream outputStream = new FileStream(output, FileMode.OpenOrCreate, FileAccess.Write);

            byte[] datiFile = new byte[inputStream.Length];
            inputStream.Read(datiFile, 0, (int)inputStream.Length);

            AesCryptoServiceProvider aes = new AesCryptoServiceProvider();
            aes.BlockSize = 128;
            aes.KeySize = 256;
            aes.Key = System.Text.ASCIIEncoding.ASCII.GetBytes(Key);
            aes.IV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV);
            aes.Padding = PaddingMode.PKCS7;
            aes.Mode = CipherMode.CBC;
            ICryptoTransform cripto = aes.CreateEncryptor(aes.Key, aes.IV);

            CryptoStream streamCifrato = new CryptoStream(outputStream, cripto, CryptoStreamMode.Write);

            streamCifrato.Write(datiFile, 0, datiFile.Length);

            streamCifrato.Close();
            inputStream.Close();
            outputStream.Close();
        }

        public void Decifra(string input, string output)
        {
            FileStream inputStream = new FileStream(input, FileMode.Open, FileAccess.Read);
            FileStream outputStream = new FileStream(output, FileMode.OpenOrCreate, FileAccess.Write);

            byte[] datiFile = new byte[inputStream.Length];
            inputStream.Read(datiFile, 0, (int)inputStream.Length);

            AesCryptoServiceProvider aes = new AesCryptoServiceProvider();
            aes.BlockSize = 128;
            aes.KeySize = 256;
            aes.Key = System.Text.ASCIIEncoding.ASCII.GetBytes(Key);
            aes.IV = System.Text.ASCIIEncoding.ASCII.GetBytes(IV);
            aes.Padding = PaddingMode.PKCS7;
            aes.Mode = CipherMode.CBC;
            ICryptoTransform cripto = aes.CreateDecryptor(aes.Key, aes.IV);

            CryptoStream streamCifrato = new CryptoStream(outputStream, cripto, CryptoStreamMode.Write);
            streamCifrato.Write(datiFile, 0, datiFile.Length);
            
            streamCifrato.Close();
            inputStream.Close();
            outputStream.Close();
        }

        private void VediFileLog_Click(object sender, EventArgs e)
        {
            FileStream LogStrea;
            string exePath = Path.GetDirectoryName(Application.ExecutablePath);
            
            if (ballControl)
            {

                Decifra(exePath + "\\passwordLogCfr.txt", exePath + "\\passwordLogXlett.txt");

                LogStrea = new FileStream(exePath + "\\passwordLogXlett.txt", FileMode.Open, FileAccess.Read);
            }
            else
            {
                Decifra(@"C:\Geminus\FileLogCfr.txt", @"C:\Geminus\FileLogXlett.txt");

                LogStrea = new FileStream(@"C:\Geminus\FileLogXlett.txt", FileMode.Open, FileAccess.Read);
            }
            
            StreamReader WDTVTxt = new StreamReader(LogStrea);

            FileDiLog FileDiLog1 = new FileDiLog();

            FileDiLog1.FileLogRTB.AppendText(WDTVTxt.ReadToEnd());

            FileDiLog1.Show();

            LogStrea.Close();

            if (File.Exists(@"C:\Geminus\FileLogXlett.txt"))
            {
                File.Delete(@"C:\Geminus\FileLogXlett.txt");
            }
            if (File.Exists(exePath + "\\passwordLogXlett.txt"))
            {
                File.Delete(exePath + "\\passwordLogXlett.txt");
            }

        }

        public void SincronizzaOGLBT_Click(object sender, EventArgs e)
        {

            if(Georgia.State == ConnectionState.Open)
            {
                Georgia.Close();
            }

            try
            {
                Tavoletta.DefaultView.RowFilter = "";
            }
            catch { }

            ExcApp = new Excel.Application();
            OGL = ExcApp.Workbooks.Open(@"\\Server\DATI\Gestione\OGLnew.xls", null, true);
            AC = OGL.Sheets["All"];

            if (!ballControl)
            {
                MySqlCommand contariga = new MySqlCommand("select finQuiOK from oglNewUltimo", Georgia);

                Georgia.Open();

                int Corridore = (int)contariga.ExecuteScalar();

                Georgia.Close();

                byte[] greenunlock = null;
                FileStream ScanImm = new FileStream(Path.GetDirectoryName(Application.ExecutablePath) + @"\greenunlockicon.png", FileMode.Open, FileAccess.Read);
                BinaryReader Binario = new BinaryReader(ScanImm);
                greenunlock = Binario.ReadBytes((int)ScanImm.Length);

                byte[] deperibileImm = null;
                FileStream ScanImmD = new FileStream(Path.GetDirectoryName(Application.ExecutablePath) + @"\DeperibileImg.png", FileMode.Open, FileAccess.Read);
                BinaryReader BinarioD = new BinaryReader(ScanImmD);
                deperibileImm = BinarioD.ReadBytes((int)ScanImmD.Length);

                byte[] AccrediaImm = null;
                FileStream ScanImmAc = new FileStream(Path.GetDirectoryName(Application.ExecutablePath) + @"\AccrediaImm.png", FileMode.Open, FileAccess.Read);
                BinaryReader BinarioAc = new BinaryReader(ScanImmAc);
                AccrediaImm = BinarioAc.ReadBytes((int)ScanImmAc.Length);

                byte[] immagineVuota = null;
                FileStream ScanImmIV = new FileStream(Path.GetDirectoryName(Application.ExecutablePath) + @"\immaginevuota.png", FileMode.Open, FileAccess.Read);
                BinaryReader BinarioIV = new BinaryReader(ScanImmIV);
                immagineVuota = BinarioIV.ReadBytes((int)ScanImmIV.Length);

                string areaLavoro = "";
                int idFamiglia = 0;
                int cody = 0;
                string Acce = "";
                string codyFam = "";

                Georgia.Open();


                while (AC.Cells[Corridore, 1].Value != null)
                {

                    if (AC.Cells[Corridore, 21].Value as string != null && (AC.Cells[Corridore, 21].Value as string).Length == 3)
                    {
                        cody = Corridore;
                        Acce = AC.Cells[Corridore, 28].Value as string;
                        codyFam = (AC.Cells[Corridore, 21].Value as string);
                    }
                    else if (AC.Cells[Corridore, 21].Value as string != null && (AC.Cells[Corridore, 21].Value as string).Length > 3 &&
                        (AC.Cells[Corridore, 21].Value as string).Substring(0, 3) == codyFam && AC.Cells[Corridore, 28].Value as string == Acce)
                    {
                        idFamiglia = cody - 8;
                    }

                    MySqlCommand InserisciRiga = new MySqlCommand("", Georgia);
                    InserisciRiga.CommandText = "insert into caricolavoro (Parametro, Accettazione, Metodo, Area, Preparativa, Determinazione, Quantificazione, Stato," +
                        " DataAnalisi, Firma, AssegnatoA, Deperibile, RiferimentoDeperibile, DataArrivo, Scadenza, DataPreparativa, DataDeterminazione, DataQuantificazione," +
                        " TecnicoPreparativa, TecnicoDeterminazione, AccettazioneInNumero, Locked, RiferimentoLocked, Matrice, Accredia, RiferimentoAccredia," +
                        " NumeroAccredia, StatoCampione, StatoParametro, SoloIntestazione, ScadenzaAnalisi, LimiteA, LimiteB, Strumento, Note, unitaMisura, CodiceFamiglia) values (@Parametro, @Accettazione, @Metodo, @Area, @Preparativa, @Determinazione, @Quantificazione, @Stato," +
                        " @DataAnalisi, @Firma, @AssegnatoA, @Deperibile, @RiferimentoDeperibile, @DataArrivo, @Scadenza, @DataPreparativa, @DataDeterminazione, @DataQuantificazione," +
                        " @TecnicoPreparativa, @TecnicoDeterminazione, @AccettazioneInNumero, @Locked, @RiferimentoLocked, @Matrice, @Accredia, @RiferimentoAccredia," +
                        " @NumeroAccredia, @StatoCampione, @StatoParametro, @SoloIntestazione, @ScadenzaAnalisi, @LimiteA, @LimiteB, @Strumento, @Note, @unitaMisura, @CodiceFamiglia)";

                    MySqlCommand InserisciRigaDett = new MySqlCommand("", Georgia);
                    InserisciRigaDett.CommandText = "insert into dettagli (Parametro, Accettazione, Metodo, Area, Preparativa, Determinazione, Quantificazione, Stato," +
                        " DataAnalisi, Firma, AssegnatoA, Deperibile, RiferimentoDeperibile, DataArrivo, Scadenza, DataPreparativa, DataDeterminazione, DataQuantificazione," +
                        " TecnicoPreparativa, TecnicoDeterminazione, AccettazioneInNumero, Locked, RiferimentoLocked, Matrice, Accredia, RiferimentoAccredia," +
                        " NumeroAccredia, StatoCampione, StatoParametro, SoloIntestazione, ScadenzaAnalisi, LimiteA, LimiteB, Strumento, Note, unitaMisura, CodiceFamiglia) values (@Parametro, @Accettazione, @Metodo, @Area, @Preparativa, @Determinazione, @Quantificazione, @Stato," +
                        " @DataAnalisi, @Firma, @AssegnatoA, @Deperibile, @RiferimentoDeperibile, @DataArrivo, @Scadenza, @DataPreparativa, @DataDeterminazione, @DataQuantificazione," +
                        " @TecnicoPreparativa, @TecnicoDeterminazione, @AccettazioneInNumero, @Locked, @RiferimentoLocked, @Matrice, @Accredia, @RiferimentoAccredia," +
                        " @NumeroAccredia, @StatoCampione, @StatoParametro, @SoloIntestazione, @ScadenzaAnalisi, @LimiteA, @LimiteB, @Strumento, @Note, @unitaMisura, @CodiceFamigliaD)";

                    if (AC.Cells[Corridore, 12].Value != null)
                    { areaLavoro = "Chimica"; }
                    else if (AC.Cells[Corridore, 13].Value != null)
                    { areaLavoro = "Cromatografia"; }
                    else if (AC.Cells[Corridore, 14].Value != null)
                    { areaLavoro = "Spettroscopia"; }
                    else if (AC.Cells[Corridore, 15].Value != null)
                    { areaLavoro = "Out sourcing"; }
                    else if (AC.Cells[Corridore, 16].Value != null)
                    { areaLavoro = "Biologia"; }

                    MySqlParameter ParametroP = new MySqlParameter();
                    ParametroP.Direction = ParameterDirection.Input;
                    ParametroP.DbType = DbType.String;
                    ParametroP.Value = AC.Cells[Corridore, 4].Value;
                    InserisciRiga.Parameters.AddWithValue("@Parametro", ParametroP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Parametro", ParametroP.Value);

                    MySqlParameter AccettazioneP = new MySqlParameter();
                    AccettazioneP.Direction = ParameterDirection.Input;
                    AccettazioneP.DbType = DbType.String;
                    AccettazioneP.Value = AC.Cells[Corridore, 28].Value;
                    InserisciRiga.Parameters.AddWithValue("@Accettazione", AccettazioneP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Accettazione", AccettazioneP.Value);

                    MySqlParameter MetodoP = new MySqlParameter();
                    MetodoP.Direction = ParameterDirection.Input;
                    MetodoP.DbType = DbType.String;
                    MetodoP.Value = AC.Cells[Corridore, 5].Value;
                    InserisciRiga.Parameters.AddWithValue("@Metodo", MetodoP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Metodo", MetodoP.Value);

                    MySqlParameter AreaP = new MySqlParameter();
                    AreaP.Direction = ParameterDirection.Input;
                    AreaP.DbType = DbType.String;
                    AreaP.Value = areaLavoro;
                    InserisciRiga.Parameters.AddWithValue("@Area", AreaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Area", AreaP.Value);

                    MySqlParameter PreparativaP = new MySqlParameter();
                    PreparativaP.Direction = ParameterDirection.Input;
                    PreparativaP.DbType = DbType.String;
                    if (AC.Cells[Corridore, 5].Value == null)
                    { PreparativaP.Value = "Sola lettura"; }
                    else
                    { PreparativaP.Value = "--"; }
                    InserisciRiga.Parameters.AddWithValue("@Preparativa", PreparativaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Preparativa", PreparativaP.Value);

                    MySqlParameter DeterminazioneP = new MySqlParameter();
                    DeterminazioneP.Direction = ParameterDirection.Input;
                    DeterminazioneP.DbType = DbType.String;
                    if (AC.Cells[Corridore, 5].Value == null)
                    { DeterminazioneP.Value = "Sola lettura"; }
                    else
                    { DeterminazioneP.Value = "--"; }
                    InserisciRiga.Parameters.AddWithValue("@Determinazione", DeterminazioneP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Determinazione", DeterminazioneP.Value);

                    MySqlParameter QuantificazioneP = new MySqlParameter();
                    QuantificazioneP.Direction = ParameterDirection.Input;
                    QuantificazioneP.DbType = DbType.String;
                    if (AC.Cells[Corridore, 5].Value == null)
                    { QuantificazioneP.Value = "Sola lettura"; }
                    else
                    { QuantificazioneP.Value = "--"; }
                    InserisciRiga.Parameters.AddWithValue("@Quantificazione", QuantificazioneP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Quantificazione", QuantificazioneP.Value);

                    MySqlParameter StatoP = new MySqlParameter();
                    StatoP.Direction = ParameterDirection.Input;
                    StatoP.DbType = DbType.String;
                    if (AC.Cells[Corridore, 5].Value == null)
                    { StatoP.Value = "Sola lettura"; }
                    else
                    { StatoP.Value = "--"; }
                    InserisciRiga.Parameters.AddWithValue("@Stato", StatoP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Stato", StatoP.Value);

                    MySqlParameter DataAnalisiP = new MySqlParameter();
                    DataAnalisiP.Direction = ParameterDirection.Input;
                    DataAnalisiP.DbType = DbType.Date;
                    DataAnalisiP.Value = DBNull.Value;
                    InserisciRiga.Parameters.AddWithValue("@DataAnalisi", DataAnalisiP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@DataAnalisi", DataAnalisiP.Value);

                    MySqlParameter DataQuantificazioneP = new MySqlParameter();
                    DataQuantificazioneP.Direction = ParameterDirection.Input;
                    DataQuantificazioneP.DbType = DbType.Date;
                    DataQuantificazioneP.Value = DBNull.Value;
                    InserisciRiga.Parameters.AddWithValue("@DataQuantificazione", DataQuantificazioneP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@DataQuantificazione", DataQuantificazioneP.Value);

                    MySqlParameter FirmaP = new MySqlParameter();
                    FirmaP.Direction = ParameterDirection.Input;
                    FirmaP.DbType = DbType.String;
                    if (AC.Cells[Corridore, 5].Value == null)
                    { FirmaP.Value = "Sola lettura"; }
                    else
                    { FirmaP.Value = "--"; }
                    InserisciRiga.Parameters.AddWithValue("@Firma", FirmaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Firma", FirmaP.Value);

                    MySqlParameter AssegnatoAP = new MySqlParameter();
                    AssegnatoAP.Direction = ParameterDirection.Input;
                    AssegnatoAP.DbType = DbType.String;
                    AssegnatoAP.Value = "--";
                    InserisciRiga.Parameters.AddWithValue("@AssegnatoA", AssegnatoAP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@AssegnatoA", AssegnatoAP.Value);

                    MySqlParameter DataArrivoP = new MySqlParameter();
                    DataArrivoP.Direction = ParameterDirection.Input;
                    DataArrivoP.DbType = DbType.Date;
                    DataArrivoP.Value = AC.Cells[Corridore, 31].Value;
                    InserisciRiga.Parameters.AddWithValue("@DataArrivo", DataArrivoP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@DataArrivo", DataArrivoP.Value);

                    MySqlParameter ScadenzaP = new MySqlParameter();
                    ScadenzaP.Direction = ParameterDirection.Input;
                    ScadenzaP.DbType = DbType.Date;
                    ScadenzaP.Value = AC.Cells[Corridore, 29].Value;
                    InserisciRiga.Parameters.AddWithValue("@Scadenza", ScadenzaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Scadenza", ScadenzaP.Value);

                    MySqlParameter DataPreparativaP = new MySqlParameter();
                    DataPreparativaP.Direction = ParameterDirection.Input;
                    DataPreparativaP.DbType = DbType.Date;
                    DataPreparativaP.Value = DBNull.Value;
                    InserisciRiga.Parameters.AddWithValue("@DataPreparativa", DataPreparativaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@DataPreparativa", DataPreparativaP.Value);

                    MySqlParameter DataDeterminazioneP = new MySqlParameter();
                    DataDeterminazioneP.Direction = ParameterDirection.Input;
                    DataDeterminazioneP.DbType = DbType.Date;
                    DataDeterminazioneP.Value = DBNull.Value;
                    InserisciRiga.Parameters.AddWithValue("@DataDeterminazione", DataDeterminazioneP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@DataDeterminazione", DataDeterminazioneP.Value);

                    MySqlParameter TecnicoPreparativaP = new MySqlParameter();
                    TecnicoPreparativaP.Direction = ParameterDirection.Input;
                    TecnicoPreparativaP.DbType = DbType.String;
                    TecnicoPreparativaP.Value = "--";
                    InserisciRiga.Parameters.AddWithValue("@TecnicoPreparativa", TecnicoPreparativaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@TecnicoPreparativa", TecnicoPreparativaP.Value);

                    MySqlParameter TecnicoDeterminazioneP = new MySqlParameter();
                    TecnicoDeterminazioneP.Direction = ParameterDirection.Input;
                    TecnicoDeterminazioneP.DbType = DbType.String;
                    TecnicoDeterminazioneP.Value = "--";
                    InserisciRiga.Parameters.AddWithValue("@TecnicoDeterminazione", TecnicoDeterminazioneP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@TecnicoDeterminazione", TecnicoDeterminazioneP.Value);

                    MySqlParameter AccettazioneInNumeroP = new MySqlParameter();
                    AccettazioneInNumeroP.Direction = ParameterDirection.Input;
                    AccettazioneInNumeroP.DbType = DbType.String;
                    AccettazioneInNumeroP.Value = (AC.Cells[Corridore, 28].Value as string).Substring(0, 4);
                    InserisciRiga.Parameters.AddWithValue("@AccettazioneInNumero", AccettazioneInNumeroP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@AccettazioneInNumero", AccettazioneInNumeroP.Value);

                    MySqlParameter MatriceP = new MySqlParameter();
                    MatriceP.Direction = ParameterDirection.Input;
                    MatriceP.DbType = DbType.String;
                    MatriceP.Value = AC.Cells[Corridore, 2].Value;
                    InserisciRiga.Parameters.AddWithValue("@Matrice", MatriceP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Matrice", MatriceP.Value);

                    MySqlParameter StatoCampioneP = new MySqlParameter();
                    StatoCampioneP.Direction = ParameterDirection.Input;
                    StatoCampioneP.DbType = DbType.String;
                    StatoCampioneP.Value = "In analisi";
                    InserisciRiga.Parameters.AddWithValue("@StatoCampione", StatoCampioneP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@StatoCampione", StatoCampioneP.Value);

                    MySqlParameter LockedP = new MySqlParameter();
                    LockedP.Direction = ParameterDirection.Input;
                    LockedP.DbType = DbType.Binary;
                    LockedP.Value = greenunlock;
                    InserisciRiga.Parameters.AddWithValue("@Locked", LockedP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Locked", LockedP.Value);

                    MySqlParameter RiferimentoLockedP = new MySqlParameter();
                    RiferimentoLockedP.Direction = ParameterDirection.Input;
                    RiferimentoLockedP.DbType = DbType.Int16;
                    RiferimentoLockedP.Value = 0;
                    InserisciRiga.Parameters.AddWithValue("@RiferimentoLocked", RiferimentoLockedP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@RiferimentoLocked", RiferimentoLockedP.Value);

                    MySqlParameter DeperibileP = new MySqlParameter();
                    DeperibileP.Direction = ParameterDirection.Input;
                    DeperibileP.DbType = DbType.Binary;

                    if (AC.Cells[Corridore, 24].Value != null)
                    {
                        DeperibileP.Value = deperibileImm;
                    }
                    else
                    {
                        DeperibileP.Value = immagineVuota;
                    }
                    InserisciRiga.Parameters.AddWithValue("@Deperibile", DeperibileP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Deperibile", DeperibileP.Value);

                    MySqlParameter RiferimentoDeperibileP = new MySqlParameter();
                    RiferimentoDeperibileP.Direction = ParameterDirection.Input;
                    RiferimentoDeperibileP.DbType = DbType.Int16;

                    if (AC.Cells[Corridore, 24].Value != null)
                    {
                        RiferimentoDeperibileP.Value = 1;
                    }
                    else
                    {
                        RiferimentoDeperibileP.Value = 0;
                    }
                    InserisciRiga.Parameters.AddWithValue("@RiferimentoDeperibile", RiferimentoDeperibileP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@RiferimentoDeperibile", RiferimentoDeperibileP.Value);

                    MySqlParameter AccrediaP = new MySqlParameter();
                    AccrediaP.Direction = ParameterDirection.Input;
                    AccrediaP.DbType = DbType.Binary;

                    if (AC.Cells[Corridore, 23].Value != null)
                    {
                        AccrediaP.Value = AccrediaImm;
                    }
                    else
                    {
                        AccrediaP.Value = immagineVuota;
                    }
                    InserisciRiga.Parameters.AddWithValue("@Accredia", AccrediaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Accredia", AccrediaP.Value);

                    MySqlParameter RiferimentoAccrediaP = new MySqlParameter();
                    RiferimentoAccrediaP.Direction = ParameterDirection.Input;
                    RiferimentoAccrediaP.DbType = DbType.Int16;

                    if (AC.Cells[Corridore, 23].Value == null)
                    {
                        RiferimentoAccrediaP.Value = 0;
                    }
                    else
                    {
                        RiferimentoAccrediaP.Value = 1;
                    }
                    InserisciRiga.Parameters.AddWithValue("@RiferimentoAccredia", RiferimentoAccrediaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@RiferimentoAccredia", RiferimentoAccrediaP.Value);

                    MySqlParameter NumeroAccrediaP = new MySqlParameter();
                    NumeroAccrediaP.Direction = ParameterDirection.Input;
                    NumeroAccrediaP.DbType = DbType.Int16;
                    if (AC.Cells[Corridore, 23].Value == 0)
                    {
                        NumeroAccrediaP.Value = 0;
                    }
                    else
                    {
                        NumeroAccrediaP.Value = AC.Cells[Corridore, 23].Value;
                    }
                    InserisciRiga.Parameters.AddWithValue("@NumeroAccredia", NumeroAccrediaP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@NumeroAccredia", NumeroAccrediaP.Value);

                    MySqlParameter StatoParametroP = new MySqlParameter();
                    StatoParametroP.Direction = ParameterDirection.Input;
                    StatoParametroP.DbType = DbType.String;
                    StatoParametroP.Value = AC.Cells[Corridore, 30].Value;
                    InserisciRiga.Parameters.AddWithValue("@StatoParametro", StatoParametroP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@StatoParametro", StatoParametroP.Value);

                    MySqlParameter SoloIntestazioneP = new MySqlParameter();
                    SoloIntestazioneP.Direction = ParameterDirection.Input;
                    SoloIntestazioneP.DbType = DbType.Boolean;
                    if (AC.Cells[Corridore, 5].Value == null)
                    { SoloIntestazioneP.Value = true; }
                    else
                    { SoloIntestazioneP.Value = false; }
                    InserisciRiga.Parameters.AddWithValue("@SoloIntestazione", SoloIntestazioneP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@SoloIntestazione", SoloIntestazioneP.Value);

                    MySqlParameter ScadenzaAnalisiP = new MySqlParameter();
                    ScadenzaAnalisiP.Direction = ParameterDirection.Input;
                    ScadenzaAnalisiP.DbType = DbType.Date;
                    if (AC.Cells[Corridore, 24].Value == null)
                    { ScadenzaAnalisiP.Value = DBNull.Value; }
                    else
                    {
                        DateTime scadenzaDeperibile = AC.Cells[Corridore, 31].Value;
                        int giorniDeperibile = (int)AC.Cells[Corridore, 24].value;
                        ScadenzaAnalisiP.Value = scadenzaDeperibile.AddDays(giorniDeperibile);
                    }
                    InserisciRiga.Parameters.AddWithValue("@ScadenzaAnalisi", ScadenzaAnalisiP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@ScadenzaAnalisi", ScadenzaAnalisiP.Value);

                    MySqlParameter LimiteAP = new MySqlParameter();
                    LimiteAP.Direction = ParameterDirection.Input;
                    LimiteAP.DbType = DbType.String;
                    if (AC.Cells[Corridore, 9].Value == null)
                    { LimiteAP.Value = "--"; }
                    else
                    { LimiteAP.Value = AC.Cells[Corridore, 9].Value; }
                    InserisciRiga.Parameters.AddWithValue("@LimiteA", LimiteAP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@LimiteA", LimiteAP.Value);

                    MySqlParameter LimiteBP = new MySqlParameter();
                    LimiteBP.Direction = ParameterDirection.Input;
                    LimiteBP.DbType = DbType.String;
                    if (AC.Cells[Corridore, 10].Value == null)
                    { LimiteBP.Value = "--"; }
                    else
                    { LimiteBP.Value = AC.Cells[Corridore, 10].Value; }
                    InserisciRiga.Parameters.AddWithValue("@LimiteB", LimiteBP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@LimiteB", LimiteBP.Value);

                    MySqlParameter StrumentoP = new MySqlParameter();
                    StrumentoP.Direction = ParameterDirection.Input;
                    StrumentoP.DbType = DbType.Int16;
                    StrumentoP.Value = idFamiglia;
                    InserisciRiga.Parameters.AddWithValue("@Strumento", StrumentoP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Strumento", StrumentoP.Value);

                    MySqlParameter NoteP = new MySqlParameter();
                    NoteP.Direction = ParameterDirection.Input;
                    NoteP.DbType = DbType.String;
                    NoteP.Value = "--";
                    InserisciRiga.Parameters.AddWithValue("@Note", NoteP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@Note", NoteP.Value);

                    MySqlParameter CodiceFamigliaP = new MySqlParameter();
                    CodiceFamigliaP.Direction = ParameterDirection.Input;
                    CodiceFamigliaP.DbType = DbType.String;
                    if (AC.Cells[Corridore, 21].Value == null)
                    { CodiceFamigliaP.Value = "--"; }
                    else
                    { CodiceFamigliaP.Value = AC.Cells[Corridore, 21].Value; }
                    InserisciRiga.Parameters.AddWithValue("@CodiceFamiglia", CodiceFamigliaP.Value);

                    MySqlParameter CodiceFamigliaPD = new MySqlParameter();
                    CodiceFamigliaPD.Direction = ParameterDirection.Input;
                    CodiceFamigliaPD.DbType = DbType.Int16;
                    if (idFamiglia != 0)
                    {
                        CodiceFamigliaPD.Value = idFamiglia;
                    }
                    InserisciRigaDett.Parameters.AddWithValue("@CodiceFamigliaD", CodiceFamigliaPD.Value);

                    MySqlParameter unitaMisuraP = new MySqlParameter();
                    unitaMisuraP.Direction = ParameterDirection.Input;
                    unitaMisuraP.DbType = DbType.String;
                    unitaMisuraP.Value = AC.Cells[Corridore, 7].Value;
                    InserisciRiga.Parameters.AddWithValue("@unitaMisura", unitaMisuraP.Value);
                    InserisciRigaDett.Parameters.AddWithValue("@unitaMisura", unitaMisuraP.Value);


                    InserisciRiga.ExecuteNonQuery();


                    //if ((AC.Cells[Corridore, 21].Value as string) != null && (AC.Cells[Corridore, 21].Value as string).Length > 3)
                    //{

                    //    InserisciRigaDett.ExecuteNonQuery();
                    //}

                    Corridore++;
                }

                MySqlCommand aggiornaData = new MySqlCommand("update Data set dataagg = @dataagg", Georgia);

                MySqlParameter dataAggP = new MySqlParameter();
                dataAggP.Direction = ParameterDirection.Input;
                dataAggP.DbType = DbType.Date;
                dataAggP.Value = DateTime.Now;
                aggiornaData.Parameters.AddWithValue("@dataagg", dataAggP.Value);

                aggiornaData.ExecuteNonQuery();


                MySqlCommand finQuiOKCm = new MySqlCommand("update oglNewUltimo set finQuiOk = " + Corridore, Georgia);

                finQuiOKCm.ExecuteNonQuery();

                Georgia.Close();
            }
            else
            {
                if (AggiornaColonne1 == null)
                { AggiornaColonne1 = new AggiornaColonne(); }
                AggiornaColonne1.ShowDialog();
            }

            OGL.Close();
            ExcApp.Quit();

            AggiornatoLB.Text = DateTime.Now.ToString();

            FiltriBT_Click(sender, e);

            if (messaggio)
            {
                MessageBox.Show("Database aggiornato");
            }
        }

        private void Schedule_KeyDown(object sender, KeyEventArgs e)
        {
            ballControl = e.Control;
        }

        private void Schedule_KeyUp(object sender, KeyEventArgs e)
        {
            ballControl = false;
        }

        private void UrgenzaBL_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (UrgenzaBL.SelectedItems[0].Index != 9)
            //{ ToDoListGV[15, rigaUrgenza].Value = UrgenzaBL.SelectedItems[0].Index + 1; }
            //UrgenzaBL.Visible = false;
        }

        private void CountDownTimer_Tick(object sender, EventArgs e)
        {

            DateTime ScaAnalisi = (DateTime)ToDoListGV["ScadenzaAnalisi", provaScadenza].Value;
            TimeSpan resto = ScaAnalisi.Subtract(DateTime.Now);
            long secondi = Convert.ToInt64(resto.TotalSeconds);

            long ore = secondi / 3600;
            long restoOre = secondi % 3600;
            long minuti = restoOre / 60;
            long restoMinuti = restoOre % 60;
            secondi = restoMinuti;

            string oreText = string.Format("{0:00}", ore);
            string minutiText = string.Format("{0:00}", minuti);
            string secondiText = string.Format("{0:00}", secondi);

            ScadenzaAnalisiLB.Text = "" + oreText + ":" + minutiText + ":" + secondiText;
        }

        private void excelBT_Click(object sender, EventArgs e)
        {
            string accettazione = ((string)ToDoListGV["Accettazione", ToDoListGV.CurrentRow.Index].Value).Substring(0,4);
            int foglioCalcolo = ((DateTime)ToDoListGV["DataArrivo", ToDoListGV.CurrentRow.Index].Value).Month;
            string meseCalcolo = string.Format("{0:00}", foglioCalcolo);
            Excel.Application foglioCalcoloApp = new Excel.Application();
            Excel.Workbook calcoliWB = foglioCalcoloApp.Workbooks.Open(@"\\Server\DATI\Gestione\Calcoli\Calcoli\Archivio\" + meseCalcolo + "\\"
                + accettazione + "-16" + "\\" + accettazione + "-16.xls");
            foglioCalcoloApp.Visible = true;
        }

        private void wordBT_Click(object sender, EventArgs e)
        {
            string[] ricerca = new string[10];
            if (ballControl)
            {
                if (ToDoListGV["RapportoDiProva", ToDoListGV.CurrentRow.Index].Value != DBNull.Value)
                {
                    string rapportoProva = ((string)ToDoListGV["RapportoDiProva", ToDoListGV.CurrentRow.Index].Value).Substring(0, 4);
                    ricerca = Directory.GetFiles(@"\\Server\DATI\Segreteria\Rapporti di prova\Rapp. di prova-2014", rapportoProva + "-14.doc", SearchOption.AllDirectories);
                }
            }
            else
            {
                string accettazione = ((string)ToDoListGV["Accettazione", ToDoListGV.CurrentRow.Index].Value).Substring(0, 4);
                ricerca = Directory.GetFiles(@"\\Server\DATI\Segreteria\Labora\Labora 2014", accettazione + "-14.doc", SearchOption.AllDirectories);
            }
            Word.Application laboraApp = new Word.Application();
            if (ricerca[0] != null)
            {
                Word.Document laboraDoc = laboraApp.Documents.Open(ricerca[0]);
                laboraApp.Visible = true;
            }
            else
            { MessageBox.Show("Il file richiesto non esiste.", "Attenzione!", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void Schedule_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (cambiamento)
            {
                DialogResult scelta;
                scelta = MessageBox.Show("Salvare le modifiche apportate?", "OGL", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                if (scelta == DialogResult.Yes)
                {
                    SalvaDatiBT_Click(sender, e);
                }

                else if (scelta == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
            }

            FileLog.Close();
            Crittografa(@"C:\Geminus\FileLog.txt", @"C:\Geminus\FileLogCfr.txt");
            File.Delete(@"C:\Geminus\FileLog.txt");
            this.Dispose();
        }

        private void salvaAutomatico()
        {
            Renfresca = new MySqlCommandBuilder(Adattatore);
            Adattatore.Update(Tavoletta);
        }

        private void Tooltip_MouseHover(object sender, EventArgs e)
        {
            string nomePulsante = (sender as KryptonButton).AccessibleName;

            foreach (Control spiega in kryptonPanel5.Controls)
            {
                if (spiega is Label && nomePulsante == spiega.AccessibleName)
                {
                    spiega.Visible = true;
                }
            }
        }

        private void Tooltip_MouseHoverFiltrPanel(object sender, EventArgs e)
        {

        }

        private void Tooltip_MouseHoverSopra(object sender, EventArgs e)
        {
            string nomePulsante = (sender as KryptonButton).AccessibleName;

            foreach (Control spiega in kryptonPanel2.Controls)
            {
                if (spiega is Label && nomePulsante == spiega.AccessibleName)
                {
                    spiega.Visible = true;
                }
            }
        }

        private void Tooltip_MouseLeave(object sender, EventArgs e)
        {
            string nomePulsante = (sender as KryptonButton).AccessibleName;

            foreach (Control spiega in kryptonPanel5.Controls)
            {
                if (spiega is Label && nomePulsante == spiega.AccessibleName)
                {
                    spiega.Visible = false;
                }
            }
        }

        private void Tooltip_MouseLeaveFiltriPanel(object sender, EventArgs e)
        {
            string nomePulsante = (sender as KryptonButton).AccessibleName;

            foreach (Control spiega in FiltriGB.Panel.Controls)
            {
                if (spiega is Label && nomePulsante == spiega.AccessibleName)
                {
                    spiega.Visible = false;
                }
            }
        }

        private void Tooltip_MouseLeaveSopra(object sender, EventArgs e)
        {
            string nomePulsante = (sender as KryptonButton).AccessibleName;

            foreach (Control spiega in kryptonPanel2.Controls)
            {
                if (spiega is Label && nomePulsante == spiega.AccessibleName)
                {
                    spiega.Visible = false;
                }
            }
        }

        private void MetodoCB_DoubleClick(object sender, EventArgs e)
        {
            if (!allargato)
            {
                MetodoCB.Width += 60;
                kryptonLabel3.Location = new Point(kryptonLabel3.Location.X + 60, kryptonLabel3.Location.Y);
                allargato = true;
            }
            else
            {
                MetodoCB.Width -= 60;
                kryptonLabel3.Location = new Point(kryptonLabel3.Location.X - 60, kryptonLabel3.Location.Y);
                allargato = false;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {



        }

        public void goodWorkPdf(DataGridView griglia)
        {
            Document doc = new Document(iTextSharp.text.PageSize.A4, 5, 5, 10, 10);
            PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(Path.GetDirectoryName(Application.ExecutablePath) +
                "\\" + "fileStampa\\" + UtenteAttivo.Cognome + ".pdf", FileMode.Create));
            doc.Open();

            int numeroColonne = 0;
            int p = 0;

            for (int m = 0; m < griglia.Columns.Count; m++)
            {
                if (griglia.Columns[m].Visible == true)
                {
                    numeroColonne++;
                }
            }

            float[] larghezze = new float[numeroColonne];

            for (int m = 0; m < griglia.Columns.Count; m++)
            {
                if (griglia.Columns[m].Visible == true)
                {
                    larghezze[p] = griglia.Columns[m].Width;
                    p++;
                }
            }


            PdfPTable goodWork = new PdfPTable(numeroColonne);
            goodWork.SetWidths(larghezze);

            for (int j = 0; j < griglia.Columns.Count; j++)
            {
                if (griglia.Columns[j].Visible == true)
                {
                    PdfPCell testataCell = new PdfPCell(new Phrase(griglia.Columns[j].HeaderText, new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 12f, iTextSharp.text.Font.BOLD)));
                    testataCell.HorizontalAlignment = 1;
                    testataCell.VerticalAlignment = 1;
                    testataCell.FixedHeight = 25;
                    goodWork.AddCell(testataCell);
                }
            }


            goodWork.HeaderRows = 1;

            for (int i = 0; i < griglia.Rows.Count; i++)
            {
                for (int k = 0; k < griglia.Columns.Count; k++)
                {
                    if (griglia.Columns[k].Visible == true && griglia[k, i].Value != null)
                    {
                        PdfPCell contenutoCell = new PdfPCell(new Phrase(griglia[k, i].Value.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8f, iTextSharp.text.Font.NORMAL)));
                        contenutoCell.VerticalAlignment = 1;
                        goodWork.AddCell(contenutoCell);
                    }
                }
            }

            doc.Add(goodWork);

            doc.Close();
        }

        private void reverseBT(object sender, EventArgs e)
        {
            string elemento = (sender as KryptonButton).AccessibleName;

            if (!(bool)reverse[elemento])
            {
                reverse[elemento] = true;
                (sender as KryptonButton).Values.Image = global::Cornelio.Properties.Resources.http_status_not_found_icon;
            }

            else
            {
                reverse[elemento] = false;
                (sender as KryptonButton).Values.Image = global::Cornelio.Properties.Resources.database_accept_icon;
            }
        }

        private void vediTuttoBT_Click(object sender, EventArgs e)
        {
            foreach (Control vistaBT in FiltriGB.Panel.Controls)
            {
                if (vistaBT.AccessibleName == "vista" && vistaBT is KryptonButton)
                {
                    string Colonna = vistaBT.Name.Substring(0, vistaBT.Name.Length - 4);
                    ToDoListGV.Columns[Colonna].Visible = true;
                    DettagliGV.Columns[Colonna].Visible = true;
                    (vistaBT as KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Ok_iconthick;
                }
            }
        }

        private void nascondiTutto_Click(object sender, EventArgs e)
        {
            foreach (Control vistaBT in FiltriGB.Panel.Controls)
            {
                if (vistaBT.AccessibleName == "vista" && vistaBT is KryptonButton)
                {
                    string Colonna = vistaBT.Name.Substring(0, vistaBT.Name.Length - 4);
                    ToDoListGV.Columns[Colonna].Visible = false;
                    DettagliGV.Columns[Colonna].Visible = false;
                    (vistaBT as KryptonButton).Values.Image = global::Cornelio.Properties.Resources.Close_2_icon;
                }
            }
        }

        private void statoCampione_Click(object sender, EventArgs e)
        {

            chiamaStato = (sender as KryptonButton).AccessibleName;

            statoCampioni1 = new StatoCampioni();
            statoCampioni1.Show();
        }

        private bool haiTagliato(string digitato)
        {
            int diversiDaZero = 0;
            bool tagliato = false;
            foreach (char carattere in digitato)
            {
                if (carattere != '0')
                {
                    diversiDaZero++;
                }
            }

            if (diversiDaZero < 3)
            {
                tagliato = true;
            }

            return tagliato;
        }

        private void DettagliGV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 34 && e.RowIndex != -1) // annullato
            {

                if (DettagliGV["StatoParametro", e.RowIndex].Value.ToString() != "annullato")
                {
                    if ((bool)DettagliGV["SoloIntestazione", e.RowIndex].Value == false)
                    {
                        DateTime DataLavoro;

                        if (DynamicRB.Checked == true)
                        { DataLavoro = DateTime.Now; }
                        else
                        { DataLavoro = DataAnalisiDP.Value; }

                        if (DettagliGV["ScadenzaAnalisi", e.RowIndex].Value != DBNull.Value)
                        {
                            DataScadenza = (DateTime)DettagliGV["ScadenzaAnalisi", e.RowIndex].Value;
                        }
                        else
                        {
                            DataScadenza = null;
                        }


                        if (e.ColumnIndex == 17 && e.RowIndex != -1) // timer
                        {
                            CountDownTimer.Stop();
                            provaScadenza = e.RowIndex;

                            if (DettagliGV["ScadenzaAnalisi", provaScadenza].Value != DBNull.Value)
                            {
                                DateTime ScaAnalisi = (DateTime)DettagliGV["ScadenzaAnalisi", provaScadenza].Value;
                                TimeSpan resto = ScaAnalisi.Subtract(DateTime.Now);
                                long secondi = Convert.ToInt64(resto.TotalSeconds);

                                if (DateTime.Now < DataScadenza && DataScadenza != null)
                                {
                                    CountDownTimer.Start();
                                }

                                else if (DateTime.Now > DataScadenza)
                                {
                                    ScadenzaAnalisiLB.Text = "Scaduto!";
                                }
                            }
                            else if (DettagliGV["ScadenzaAnalisi", provaScadenza].Value == DBNull.Value)
                            {
                                ScadenzaAnalisiLB.Text = "00:00:00";
                            }
                        }


                        if (e.ColumnIndex == 24 && e.RowIndex != -1) // Preparativa
                        {
                            cambiamento = true;
                            if (DettagliGV["Preparativa", e.RowIndex].Value.ToString() == "--")
                            {
                                if (DataLavoro < DataScadenza || DataScadenza == null)
                                {
                                    DettagliGV["DataPreparativa", e.RowIndex].Value = DataLavoro;
                                    DettagliGV["TecnicoPreparativa", e.RowIndex].Value = UtenteAttivo.Nome + " " + UtenteAttivo.Cognome;
                                    DettagliGV["Preparativa", e.RowIndex].Value = "Completato";
                                }
                                else
                                {
                                    pretrattamentoRow = e.RowIndex;
                                    Pretrattamento Pretrattamento1 = new Pretrattamento();
                                    Pretrattamento1.Height = 237;
                                    Pretrattamento1.ShowDialog();
                                }
                            }

                            else if (DettagliGV["Preparativa", e.RowIndex].Value.ToString() == "Completato"
                                && DettagliGV["Determinazione", e.RowIndex].Value.ToString() == "--"
                                && (DettagliGV["TecnicoPreparativa", e.RowIndex].Value.ToString() == UtenteAttivo.Nome + " " + UtenteAttivo.Cognome
                                || UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio"))
                            {
                                DettagliGV["DataPreparativa", e.RowIndex].Value = DBNull.Value;
                                DettagliGV["TecnicoPreparativa", e.RowIndex].Value = "--";
                                DettagliGV["Preparativa", e.RowIndex].Value = "--";
                                DettagliGV.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                            }
                        }

                        if (e.ColumnIndex == 25 && e.RowIndex != -1) // Determinazione
                        {
                            cambiamento = true;
                            if (DettagliGV["Determinazione", e.RowIndex].Value.ToString() == "--")
                            {

                                if (DettagliGV["Preparativa", e.RowIndex].Value.ToString() == "--")
                                {
                                    DataGridViewCellEventArgs ePrep = new DataGridViewCellEventArgs(24, e.RowIndex);
                                    DettagliGV_CellClick(sender, ePrep);
                                }

                                DettagliGV["DataDeterminazione", e.RowIndex].Value = DataLavoro;
                                DettagliGV["TecnicoDeterminazione", e.RowIndex].Value = UtenteAttivo.Nome + " " + UtenteAttivo.Cognome;
                                DettagliGV["Determinazione", e.RowIndex].Value = "Completato";
                            }

                            else if (DettagliGV["Determinazione", e.RowIndex].Value.ToString() == "Completato"
                                && DettagliGV["Quantificazione", e.RowIndex].Value.ToString() == "--"
                                && (DettagliGV["TecnicoDeterminazione", e.RowIndex].Value.ToString() == UtenteAttivo.Nome + " " + UtenteAttivo.Cognome ||
                                UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio"))
                            {
                                DettagliGV["DataDeterminazione", e.RowIndex].Value = DBNull.Value;
                                DettagliGV["TecnicoDeterminazione", e.RowIndex].Value = "--";
                                DettagliGV["Determinazione", e.RowIndex].Value = "--";
                            }
                        }

                        if (e.ColumnIndex == 8 && e.RowIndex != -1) // Quantificazione
                        {
                            cambiamento = true;
                            if (DettagliGV["RiferimentoLocked", e.RowIndex].Value.ToString() == "0" || DettagliGV["AssegnatoA", e.RowIndex].Value.ToString() == UtenteAttivo.Cognome)
                            {
                                string AccettazioneChamp = DettagliGV["Accettazione", e.RowIndex].Value.ToString();


                                if (DettagliGV["RiferimentoAccredia", e.RowIndex].Value.ToString() == "0" || UtenteAttivo.ProveAccreditate[(int)(DettagliGV["NumeroAccredia", e.RowIndex].Value)] == 1)
                                {
                                    if (DettagliGV["Quantificazione", e.RowIndex].Value.ToString() == "--")
                                    {
                                        if (DettagliGV["AssegnatoA", e.RowIndex].Value.ToString() == "--")
                                        {
                                            DettagliGV["AssegnatoA", e.RowIndex].Value = UtenteAttivo.Cognome;
                                        }

                                        if (DettagliGV["Preparativa", e.RowIndex].Value.ToString() == "--")
                                        {
                                            DataGridViewCellEventArgs ePrep = new DataGridViewCellEventArgs(24, e.RowIndex);
                                            DettagliGV_CellClick(sender, ePrep);
                                        }


                                        if (DettagliGV["Determinazione", e.RowIndex].Value.ToString() == "--")
                                        {
                                            DettagliGV["DataDeterminazione", e.RowIndex].Value = DataLavoro;
                                            DettagliGV["TecnicoDeterminazione", e.RowIndex].Value = UtenteAttivo.Cognome;
                                            DettagliGV["Determinazione", e.RowIndex].Value = "Completato";
                                        }

                                        DettagliGV["Stato", e.RowIndex].Value = "Completato";
                                        DettagliGV["DataQuantificazione", e.RowIndex].Value = DataLavoro;
                                        DettagliGV["DataAnalisi", e.RowIndex].Value = DataLavoro;
                                        DettagliGV["Firma", e.RowIndex].Value = UtenteAttivo.Nome + " " + UtenteAttivo.Cognome;
                                        DettagliGV["Quantificazione", e.RowIndex].Value = "Completato";

                                    }

                                    else if (DettagliGV["Quantificazione", e.RowIndex].Value.ToString() == "Completato"
                                        && (DettagliGV["Firma", e.RowIndex].Value.ToString() == UtenteAttivo.Nome + " " + UtenteAttivo.Cognome ||
                                        UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio"))
                                    {
                                        DettagliGV["Stato", e.RowIndex].Value = "--";
                                        DettagliGV["DataQuantificazione", e.RowIndex].Value = DBNull.Value;
                                        DettagliGV["DataAnalisi", e.RowIndex].Value = DBNull.Value;
                                        DettagliGV["Firma", e.RowIndex].Value = "--";
                                        DettagliGV["Quantificazione", e.RowIndex].Value = "--";

                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Utente non abilitato alla prova.");
                                }
                            }
                            else
                            {
                                MessageBox.Show("L'analisi deve essere quantificata obligatoriamente da " + DettagliGV["AssegnatoA", e.RowIndex].Value.ToString());
                            }
                        }



                        if (e.ColumnIndex == 12 && e.RowIndex != -1) // assegnato a 
                        {
                            cambiamento = true;
                            if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica")
                            {
                                if (DettagliGV["AssegnatoA", e.RowIndex].Value.ToString() == "--")
                                {
                                    if (SceltaTecnicoCB.SelectedIndex != -1)
                                    { DettagliGV["AssegnatoA", e.RowIndex].Value = (SceltaTecnicoCB.SelectedItem as System.Data.DataRowView)["Cognome"].ToString(); }
                                    else
                                    {
                                        MessageBox.Show("Attenzione! Non è stato selezionato nessun tecnico.");
                                    }
                                }

                                else
                                { DettagliGV["AssegnatoA", e.RowIndex].Value = "--"; }

                            }
                            else
                            {
                                MessageBox.Show("Utente non abilitato alla modifica");
                            }
                        }

                        if (e.ColumnIndex == 13 && e.RowIndex != -1) // Lucchetto
                        {
                            cambiamento = true;
                            if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica")
                            {
                                if ((string)DettagliGV["AssegnatoA", e.RowIndex].Value != "--")
                                {
                                    int PidLocked = 0;
                                    string StatoLock = "";

                                    if ((int)DettagliGV["RiferimentoLocked", e.RowIndex].Value == 0)
                                    {
                                        PidLocked = 1;
                                        StatoLock = "redlockicon.png";
                                    }

                                    else if ((int)DettagliGV["RiferimentoLocked", e.RowIndex].Value == 1)
                                    {
                                        PidLocked = 0;
                                        StatoLock = "greenunlockicon.png";
                                    }

                                    byte[] Padlock = null;
                                    FileStream ScanImm = new FileStream(Path.GetDirectoryName(Application.ExecutablePath) + @"\" + StatoLock, FileMode.Open, FileAccess.Read);
                                    BinaryReader Binario = new BinaryReader(ScanImm);
                                    Padlock = Binario.ReadBytes((int)ScanImm.Length);

                                    DettagliGV["RiferimentoLocked", e.RowIndex].Value = PidLocked;
                                    DettagliGV["Locked", e.RowIndex].Value = Padlock;
                                }
                                else
                                {
                                    MessageBox.Show("Attenzione! La prova non è stata assegnata a nessun tecnico.");
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Il parametro è in sola lettura. Impossibile apportare modifiche.");
                    }
                }

                else
                {
                    MessageBox.Show("Il parametro è stato annullato. Impossibile apportare modifiche.");
                }
            }

            else if (e.ColumnIndex == 34 && e.RowIndex != -1)
            {
                cambiamento = true;
                if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica" || UtenteAttivo.Qualifica == "Addetto commerciale")
                {
                    if ((string)DettagliGV["Preparativa", e.RowIndex].Value == "--")
                    { DettagliGV["Preparativa", e.RowIndex].Value = "annullato"; }
                    else if ((string)DettagliGV["Preparativa", e.RowIndex].Value == "annullato")
                    { DettagliGV["Preparativa", e.RowIndex].Value = "--"; }

                    if ((string)DettagliGV["Determinazione", e.RowIndex].Value == "--")
                    { DettagliGV["Determinazione", e.RowIndex].Value = "annullato"; }
                    else if ((string)DettagliGV["Determinazione", e.RowIndex].Value == "annullato")
                    { DettagliGV["Determinazione", e.RowIndex].Value = "--"; }

                    if ((string)DettagliGV["Quantificazione", e.RowIndex].Value == "--")
                    { DettagliGV["Quantificazione", e.RowIndex].Value = "annullato"; }
                    else if ((string)DettagliGV["Quantificazione", e.RowIndex].Value == "annullato")
                    { DettagliGV["Quantificazione", e.RowIndex].Value = "--"; }

                    if ((string)DettagliGV["StatoParametro", e.RowIndex].Value == "accettato")
                    { DettagliGV["StatoParametro", e.RowIndex].Value = "annullato"; }
                    else
                    { DettagliGV["StatoParametro", e.RowIndex].Value = "accettato"; }
                }

                else
                {
                    MessageBox.Show("Non si dispongono dei permessi necessari per annullare il parametro.");
                }
            }


        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dettagliCampioneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MascheraCampione DettagliCampione = new MascheraCampione();
            DettagliCampione.Show();
        }

        private void selettivaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelettivaCB.Checked = true;
        }

        private void completaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CompletaCB.Checked = true;
        }

        private void staticaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            StaticRB.Checked = true;
        }

        private void semistaticaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SemiStaticRB.Checked = true;
        }

        private void dinamicaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DynamicRB.Checked = true;
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            sediciNoni = false;
            MonitorBT_Click(sender, e);
        }

        private void quattroQuartiBT_Click(object sender, EventArgs e)
        {
            sediciNoni = true;
            MonitorBT_Click(sender, e);
        }

        private void importaMetalliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string numeroAcc = (string)ToDoListGV["Accettazione", ToDoListGV.CurrentRow.Index].Value;
            int foglioCalcolo = ((DateTime)ToDoListGV["DataArrivo", ToDoListGV.CurrentRow.Index].Value).Month;
            string meseCalcolo = string.Format("{0:00}", foglioCalcolo);

            ExcApp = new Excel.Application();
            OGL = ExcApp.Workbooks.Open(@"\\Server\DATI\Gestione\Calcoli\Calcoli\Archivio\" + meseCalcolo +  "\\" + numeroAcc.Substring(0,4) + "-16\\" + numeroAcc.Substring(0,4) + "-16.xls");
            Excel.Worksheet met;

            int riga = 1;
            int rigaBianco = 1;
            int colonnaMetallo = 5;
            int colonnaSimbolo = 1;
            int colonnaCstrumentale = 1;
            int rigaMetallo = 1;
            int rigaStart = 1;
            int puntoEmissione = 1;
            string nomePunto = "";
            char separatore = ' ';
            string[] qualeMetallo = new string[3];
            double risultato = 0.0;
            double bianco = 0.0;
            met = OGL.Sheets[numeroAcc.Substring(0, 4) + " 16 Metalli"];


            try
            {
                AC = OGL.Sheets["Sheet1"];
            }
            catch
            {
                MessageBox.Show("il foglio con i dati grezzi non è stato importato correttamente. Controllare.");
                OGL.Close(false);
                ExcApp.Quit();
                return;
            }


            if ((ToDoListGV["Matrice", ToDoListGV.CurrentRow.Index].Value as string).Substring(0, 4) == "Test")
            {
                met = OGL.Sheets[numeroAcc.Substring(0, 4) + " 16 Metalli(TC)"];
                numeroAcc = numeroAcc + " TC";
            }
            else if (ToDoListGV["Matrice", ToDoListGV.CurrentRow.Index].Value as string == "Emissioni in atmosfera")
            {
                nomePunto = numeroAcc.Substring(numeroAcc.Length - 2, 2);
            }

            while (AC.Cells[riga, 1].Value != numeroAcc && riga < 200)
            {
                riga++;
            }

            rigaBianco = riga;

            while (rigaBianco > 6 && (AC.Cells[rigaBianco, 1].Value as string) != null && (AC.Cells[rigaBianco, 1].Value as string).Substring(0, 3) != "bia")
            {
                rigaBianco--;
            }

            while (AC.Cells[6, colonnaMetallo].Value as string != null)
            {
                
                colonnaSimbolo = 1;
                colonnaCstrumentale = 1;
                rigaMetallo = 1;
                rigaStart = 1;

                if (rigaBianco < 4 || !(AC.Cells[rigaBianco, colonnaMetallo].Value is double))
                {
                    bianco = 0.0;
                }
                else if(AC.Cells[rigaBianco, colonnaMetallo].Value is double)
                {
                    
                    bianco = AC.Cells[rigaBianco, colonnaMetallo].Value;
                }



                qualeMetallo = (AC.Cells[6, colonnaMetallo].Value as string).Split(separatore);

                try
                {
                    risultato = AC.Cells[riga, colonnaMetallo].Value - bianco;

                    while (met.Cells[rigaStart, 1].Value != "Start")
                    {
                        rigaStart++;
                    }

                    while (met.Cells[rigaStart, colonnaSimbolo].Value != "simbolo" && colonnaSimbolo < 100)
                    {
                        colonnaSimbolo++;
                    }

                    while (met.Cells[rigaStart, colonnaCstrumentale].Value != "cStrumentale" && colonnaCstrumentale < 100)
                    {
                        colonnaCstrumentale++;
                    }

                    if (ToDoListGV["Matrice", ToDoListGV.CurrentRow.Index].Value as string == "Emissioni in atmosfera")
                    {
                        while (met.Cells[rigaStart, puntoEmissione].Value != "puntoEmissione" && puntoEmissione < 100)
                        {
                            puntoEmissione++;
                        }

                        while (met.Cells[rigaStart, puntoEmissione].Value != nomePunto && rigaStart < 100)
                        {
                            rigaStart++;
                        }
                        rigaMetallo = rigaStart;
                    }
                    while (met.Cells[rigaMetallo, colonnaSimbolo].Value != qualeMetallo[0] && rigaMetallo < 100)
                    {
                        rigaMetallo++;
                    }

                    if (rigaMetallo < 100 && colonnaSimbolo < 100 && rigaMetallo < 100)
                    {
                        met.Cells[rigaMetallo, colonnaCstrumentale].Value = risultato;
                    }
                }

                catch { }

                colonnaMetallo++;
            }
            OGL.Close(true);
            ExcApp.Quit();

            MessageBox.Show("Finito. Vedi un po'...");

        }

        private void aggiornaColonneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ballControl = true;
            SincronizzaOGLBT_Click(sender, e);
            ballControl = false;
        }

        private void cambiaStatoCampioneSB_Click(object sender, EventArgs e)
        {
            cambiaStatoCampione cambiaStatoCampione1 = new cambiaStatoCampione();
            cambiaStatoCampione1.ShowDialog();
        }

        private void chiudiConnessioniSB_Click(object sender, EventArgs e)
        {
            if (Georgia.State == ConnectionState.Open)
            {
                Georgia.Close();
            }
            if (MascheraAnalita.Georgia.State == ConnectionState.Open)
            {
                MascheraAnalita.Georgia.Close();
            }
            if (MascheraCampione.Georgia.State == ConnectionState.Open)
            {
                MascheraCampione.Georgia.Close();
            }
            if (SlotImpostazioni.Georgia.State == ConnectionState.Open)
            {
                SlotImpostazioni.Georgia.Close();
            }
            if (FiltriAvanzati.Georgia.State == ConnectionState.Open)
            {
                FiltriAvanzati.Georgia.Close();
            }
            if (AggiornaColonne.Georgia.State == ConnectionState.Open)
            {
                AggiornaColonne.Georgia.Close();
            }
            if (cambiaStatoCampione.Georgia.State == ConnectionState.Open)
            {
                cambiaStatoCampione.Georgia.Close();
            }

            MessageBox.Show("Connessioni chiuse.");
        }

        private void controllaStatoCampioniSB_Click(object sender, EventArgs e)
        {
            int stato = 1;
            int conta = 0;
            MySqlCommand daApprovare = new MySqlCommand("select distinct Accettazione from caricolavoro where StatoCampione = 'In analisi'", Georgia);

            Georgia.Open();
            MySqlDataReader WDTVFin = daApprovare.ExecuteReader();

            while (WDTVFin.Read())
            {
                conta = 0;

                MySqlCommand vediBene = new MySqlCommand
                    ("select Accettazione, Quantificazione from caricolavoro where Accettazione = '" + (string)WDTVFin["Accettazione"] + "' and Quantificazione = '--'", Wyoming);

                Wyoming.Open();

                MySqlDataReader WDTVVbn = vediBene.ExecuteReader();

                while (WDTVVbn.Read())
                {
                    conta++;
                }
                Wyoming.Close();

                if (conta > 0)
                {
                    MySqlCommand completa = new MySqlCommand
                        ("update caricolavoro set StatoCampione = 'Da approvare' where Accettazione = '" + (string)WDTVFin["Accettazione"] + "'", Wyoming);

                    Wyoming.Open();
                    completa.ExecuteNonQuery();
                    Wyoming.Close();
                }
            }
            Georgia.Close();
            MessageBox.Show("controllo terminato");
        }

        private void congelaColonnaSB_Click(object sender, EventArgs e)
        {

           int numeroColonna = ToDoListGV.SelectedCells[0].ColumnIndex;

           ToDoListGV.Columns[numeroColonna].Frozen = true;

        }

        private void sbloccaColonneSB_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ToDoListGV.Columns.Count; i++)
            {
                ToDoListGV.Columns[i].Frozen = false;
            }
        }

        private void InserisciNotaSB_Click(object sender, EventArgs e)
        {
            if (campoNote1 == null)
            {
                campoNote1 = new CampoNote();
            }
                campoNote1.ShowDialog();
        }

        private void DynamicRB_CheckedChanged(object sender, EventArgs e)
        {
            if (DynamicRB.Checked == true)
            {
                DataAnalisiDP.Enabled = false;
            }
            else
            {
                DataAnalisiDP.Enabled = true;
            }
        }

        private void ToDoListGV_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // risultato
            #region
            if ((e.ColumnIndex == 6 || e.ColumnIndex == 43) && e.RowIndex != -1 && (UtenteAttivo.Nome == "Pasquale" || UtenteAttivo.Nome == "Adamo" || UtenteAttivo.Nome == "Marcello"))
            {
                ToDoListGV.EditMode = DataGridViewEditMode.EditOnEnter;

                foreach (DataGridViewColumn colonna in ToDoListGV.Columns)
                {
                    if (colonna.Name != "Risultato" && colonna.Name != "Incertezza")
                    {
                        ToDoListGV.Columns[colonna.Name].ReadOnly = true;
                    }
                    else
                    {
                        ToDoListGV.Columns[colonna.Name].ReadOnly = false;
                    }
                }
            }
            #endregion
        }

        private void vectorBT_Click(object sender, EventArgs e)
        {

            Georgia.Open();
            MySqlCommand scriviNA = new MySqlCommand
                ("update data set Accettazione = '" + ToDoListGV["Accettazione", ToDoListGV.SelectedRows[0].Index].Value.ToString() +
                "' , meseAccettazione = '" + ToDoListGV["DataArrivo", ToDoListGV.SelectedRows[0].Index].Value.ToString().Substring(3, 2) + "'", Georgia);
            scriviNA.ExecuteNonQuery();
            Georgia.Close();

            Process.Start(@"\\Server\dati\Gestione\Vector 2016\bin\Release\Vector.exe");
        }

        private void esportaTabellaInExcelSB_Click(object sender, EventArgs e)
        {
            esportaXls(ToDoListGV);
        }

        private void editOnEnterSB_Click(object sender, EventArgs e)
        {

            if (UtenteAttivo.Qualifica == "Responsabile Tecnico di Laboratorio" || UtenteAttivo.Qualifica == "Responsabile Area Chimica" || UtenteAttivo.Qualifica == "Responsabile Area Microbiologia")
            {
                if (ToDoListGV.EditMode == DataGridViewEditMode.EditProgrammatically)
                {
                    ToDoListGV.EditMode = DataGridViewEditMode.EditOnEnter;
                }
                else if (ToDoListGV.EditMode == DataGridViewEditMode.EditOnEnter)
                {
                    ToDoListGV.EditMode = DataGridViewEditMode.EditProgrammatically;
                }
            }
        }

        private void copiaSB_Click(object sender, EventArgs e)
        {
            copiaSt = ToDoListGV["Note", ToDoListGV.SelectedCells[0].RowIndex].Value as string;
            MessageBox.Show("Valore copiato");
        }

        private void incollaSB_Click(object sender, EventArgs e)
        {
            int numeroSelezioneS = ToDoListGV.SelectedRows.Count;

            if (numeroSelezioneS == 0)
            {
                MessageBox.Show("Non è stata selezionata nessuna riga");
            }
            else if (copiaSt == "")
            {
                MessageBox.Show("Attenzione! Non è stato copiato alcun valore");
            }
            else
            {
                for (int i = numeroSelezioneS -1; i > -1; i--)
                {
                    ToDoListGV["Note", ToDoListGV.SelectedRows[i].Index].Value = copiaSt;
                }
            }
        }

        public void esportaXls (DataGridView griglia)
        {
            SaveFileDialog csvPath = new SaveFileDialog();

            if (csvPath.ShowDialog() == DialogResult.OK)
            {
                ExcApp = new Excel.Application();
                Excel.Workbook csv = ExcApp.Workbooks.Add();
                Excel.Worksheet csvSheet = csv.Sheets.Add();

                int rigaExc = 2;
                int colonnaExc = 2;

                for (int k = 0; k < griglia.Columns.Count; k++)
                {
                    if (griglia.Columns[k].Visible == true)
                    {
                        csvSheet.Cells[rigaExc, colonnaExc].Value = griglia.Columns[k].HeaderText;
                        Excel.Range grasse = csvSheet.Cells[rigaExc, colonnaExc];
                        grasse.Font.Bold = true;
                        colonnaExc++;
                    }
                }
                rigaExc = 3;
                colonnaExc = 2;

                for (int i = 0; i < griglia.Rows.Count; i++)
                {
                    for (int k = 0; k < griglia.Columns.Count; k++)
                    {
                        if (griglia.Columns[k].Visible == true)
                        {
                            csvSheet.Cells[rigaExc, colonnaExc].Value = griglia[k, i].Value;
                            colonnaExc++;
                        }
                    }
                    rigaExc++;
                    colonnaExc = 2;
                }


                csv.SaveAs(csvPath.FileName);
                csv.Close();
                ExcApp.Quit();
                MessageBox.Show("controll");
            }
        }

        private void backDoorBT_Click(object sender, EventArgs e)
        {
            #region 
            //string accettazione = ((string)ToDoListGV["Accettazione", ToDoListGV.CurrentRow.Index].Value).Substring(0, 4);
            //int foglioCalcolo = ((DateTime)ToDoListGV["DataArrivo", ToDoListGV.CurrentRow.Index].Value).Month;
            //string meseCalcolo = string.Format("{0:00}", foglioCalcolo);
            //string percorso = @"\\Server\DATI\Gestione\Calcoli\Calcoli\Archivio\" + meseCalcolo + "\\"
            //    + accettazione + "-15" + "\\" + accettazione + "-15.xls";
            //string foglioCalc = ToDoListGV["FoglioCalcolo", ToDoListGV.SelectedRows[0].Index].Value as string;



            //if (foglioCalc != null)
            //{
            //    Excel.Application foglioCalcoloApp = new Excel.Application();
            //    Excel.Workbook foglioCalcoloWB = foglioCalcoloApp.Workbooks.Open(percorso, null, true);
            //    Excel.Worksheet wsi = foglioCalcoloWB.Sheets[foglioCalc];

            //    int LinioStart = 1;
            //    int LinioStop = 1;
            //    int ColunnaBlocco = 1;
            //    int ColunnaStart = 1;

            //    while ((wsi.Cells[LinioStart, 1].Value) as string != (string)ToDoListGV["Start", ToDoListGV.CurrentRow.Index].Value)
            //    {
            //        LinioStart++;
            //        if (LinioStart > 200) // interrompe il sottoprogramma se non trova il riferimento nel foglio excel
            //        {
            //            MessageBox.Show("Riferimento Start non trovato. Controllare il foglio di calcolo.");
            //            return;
            //        }
            //    }
            //    while ((wsi.Cells[LinioStart, ColunnaStart].Value as string) != (string)ToDoListGV["RifC", ToDoListGV.CurrentRow.Index].Value)
            //    {
            //        ColunnaStart++;
            //        if (ColunnaStart > 150) // interrompe il sottoprogramma se non trova il riferimento nel foglio excel
            //        {
            //            MessageBox.Show("Riferimento RifC non trovato. Controllare il foglio di calcolo.");
            //            return;
            //        }
            //    }
            //    while ((wsi.Cells[LinioStart, ColunnaBlocco].Value as string) != (string)ToDoListGV["Blocco", ToDoListGV.CurrentRow.Index].Value)
            //    {
            //        ColunnaBlocco++;
            //        if (ColunnaBlocco > 150) // interrompe il sottoprogramma se non trova il riferimento nel foglio excel
            //        {
            //            MessageBox.Show("Riferimento BloccoOne non trovato. Controllare il foglio di calcolo.");
            //            return;
            //        }
            //    }
            //    while ((wsi.Cells[LinioStart, ColunnaBlocco].Value as string) != (string)ToDoListGV["RifL", ToDoListGV.CurrentRow.Index].Value)
            //    {
            //        LinioStart++;
            //        if (LinioStart > 150) // interrompe il sottoprogramma se non trova il riferimento nel foglio excel
            //        {
            //            MessageBox.Show("Riferimento RifL non trovato. Controllare il foglio di calcolo.");
            //            return;
            //        }
            //    }

            //    double valoreTrovato = wsi.Cells[LinioStart, ColunnaStart].Value;

            //    MessageBox.Show(valoreTrovato.ToString());

            //    foglioCalcoloApp.Visible = true;
            //}
            //else
            //{
            //    MessageBox.Show("Il percorso del foglio o il nome del foglio non sono stati registrati. Controllare");
            //}
            #endregion

            idSelezionato = (int) ToDoListGV["ID", ToDoListGV.CurrentRow.Index].Value;

            if (((string)ToDoListGV["matrice", ToDoListGV.CurrentRow.Index].Value).Substring(0, 4) == "Acqu")
            {
                dettagliBio dettagliBio1 = new dettagliBio();
                dettagliBio1.Show();
            }
            else if (((string)ToDoListGV["matrice", ToDoListGV.CurrentRow.Index].Value).Substring(0, 4).ToLower() == "tamp")
            {
                tamponi tamponi1 = new tamponi();
                tamponi1.Show();
            }

            else
            {
                alimenti alimenti1 = new alimenti();
                alimenti1.Show();
            }
            #region
            //int NumeroSelezioneS = ToDoListGV.SelectedRows.Count;

            //if (NumeroSelezioneS == 0)
            //{
            //    MessageBox.Show("Non è stata selezionata nessuna riga.");
            //}
            //else
            //{
            //    for (int i = NumeroSelezioneS - 1; i > -1; i--)
            //    {
            //        Georgia.Open();
            //        MySqlCommand figli = new MySqlCommand
            //            ("update Dettagli set Quantificazione = 'Completato' where Accettazione = '" + ToDoListGV["Accettazione", ToDoListGV.SelectedRows[i].Index].Value +
            //            "' and CodiceFamiglia = " + ToDoListGV["ID", ToDoListGV.SelectedRows[i].Index].Value, Georgia);
            //        figli.ExecuteNonQuery();
            //        Georgia.Close();
            //    }
            //}
            #endregion


            //MySqlConnection Illinois = new MySqlConnection(@"SERVER=192.168.1.250; Database=carolinapanthers2016; User ID=marcello; Password=jesussss@79");

            //Illinois.Open();
            //MySqlCommand caricaCm = new MySqlCommand
            //    ("SELECT colorado2016.campioni.accettazione, colorado2016.clientiProp.RagioneSociale, colorado2016.clientiComm.RagioneSociale, caricolavoro.Matrice " +
            //    "FROM colorado2016.campioni LEFT JOIN colorado2016.clientiProp ON colorado2016.campioni.idProprietario = colorado2016.clientiProp.IDCliente " +
            //    "LEFT JOIN colorado2016.clientiComm ON colorado2016.campioni.idCommittente = colorado2016.clientiComm.IDCliente " +
            //    "LEFT JOIN caricolavoro ON colorado2016.campioni.accettazione = caricolavoro.Accettazione", Illinois);

            //MySqlDataReader WDTVSC = caricaCm.ExecuteReader();
            //DataTable tavolettaIll = new DataTable();
            //tavolettaIll.Load(WDTVSC);

            //Illinois.Close();
        }

        private void helpBT_Click(object sender, EventArgs e)
        {
            Word.Application help = new Word.Application();
            Word.Document eagle = help.Documents.Open(@"\\Server\dati\Gestione\UNI EN ISO IEC 17025\Procedure\PG Chim 04 Accettazione campioni identificazione e conservazione\OGL.doc");
            help.Visible = true;
        }

        private void inserisciRisultatoBT_Click(object sender, EventArgs e)
        {
            Georgia.Open();

            foreach (DataGridViewRow riga in ToDoListGV.SelectedRows)
            {
                //Calcolo il mese e lo formatto per costruire il path del foglio di calcolo

                string accettazione = (riga.Cells["accettazione"].Value as string).Substring(0,4);
                int meseCalcoloInt = ((DateTime)riga.Cells["DataArrivo"].Value).Month;
                string meseCalcolo = string.Format("{0:00}", meseCalcoloInt);
                string risultatoStr = "";



                MySqlCommand trovaPar = new MySqlCommand
                ("SELECT * FROM colorado2016.parametri WHERE id = @id", Georgia);

                MySqlParameter idP = new MySqlParameter();
                idP.Direction = ParameterDirection.Input;
                idP.DbType = DbType.String;
                idP.Value = riga.Cells["idParametro"].Value;
                trovaPar.Parameters.AddWithValue("@id", idP.Value);

                MySqlDataReader vivo = trovaPar.ExecuteReader();

                vivo.Read();

                string nomeFoglio = vivo["nomeFoglio"] as string;
                string[] rif = (vivo["rif1"] as string).Split(',');
                int rigaEx = Convert.ToInt32(rif[0]);
                int colonnaEx = Convert.ToInt32(rif[1]);

                vivo.Dispose();

                Excel.Application risultatoEA = new Excel.Application();

                Excel.Workbook risultatoWB = risultatoEA.Workbooks.Open(@"\\Server\DATI\Gestione\Calcoli\Calcoli\Archivio\" + meseCalcolo + "\\"
                     + accettazione + "-16" + "\\" + accettazione + "-16.xls");

                Excel.Worksheet risultatoWs = risultatoWB.Sheets[accettazione + " 16 " + nomeFoglio];

                double decimoMeridio = Convert.ToDouble(risultatoWs.Cells[rigaEx,colonnaEx].Value);

                risultatoWB.Close();
                risultatoEA.Quit();

                risultatoStr = formatta.significative(decimoMeridio);

                if(formatta.haiTagliato(risultatoStr)==1)
                { risultatoStr = risultatoStr + "0"; }
                else if(formatta.haiTagliato(risultatoStr) == 2)
                { risultatoStr = risultatoStr + ",0"; }

                riga.Cells["Risultato"].Value = risultatoStr;

            }

            

            Georgia.Close();
        }

        private void creaReportBT_Click(object sender, EventArgs e)
        {


        }

        private void LogoPB_Click_1(object sender, EventArgs e)
        {

        }
    }
}