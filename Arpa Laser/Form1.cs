using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using WMPLib; // Includo la libreria di Windows Media Player per la riproduzione di suoni.

namespace Arpa_Laser
{
    public partial class Form1 : Form
    {
        // Dichiaro l'oggetto SerialPort.
        SerialPort sp;

        // Oggetto DataTime, lo utilizzo per sapere l'ora in cui scrivo nella TextBox.
        static DateTime dt = DateTime.Now;
        String time = dt.ToShortTimeString();
        
        // Variabile per contenere le porte dispobili.
        string[] ports;

        // Lettore musicale, serve per riprodurre i suoni.
        WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();


        public Form1()
        {
            InitializeComponent();
            
            // Setto le proprietà dei miei oggetti. 
            label2.Text = "";
            label2.Enabled = false;
            button2.Enabled = false;
            comboBox2.Enabled = false;

            // Permette di avere accesso alla form anche da diversi thread.
            CheckForIllegalCrossThreadCalls = false;

            // RichTextBox non editabile.
            richTextBox1.ReadOnly = true;
            richTextBox1.BackColor = System.Drawing.SystemColors.Window;
        }



        // Tutte quelle opzioni che vengono caricate all'avvio del programma.
        private void Form1_Load(object sender, EventArgs e)
        {
            ports = SerialPort.GetPortNames();

            // Aggiungo tutte le porte seriali disponibili alla comboBox.
            if (ports.Length != 0)
            {
                foreach (string port in ports)
                {
                    comboBox1.Items.Add(port);
                }
            }
            else
            {
                MessageBox.Show("Non ci sono porte seriali disponibili!", "Errore...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                comboBox1.Enabled = false;
                button1.Enabled = false;
            }


            //////////////
            // Carico gli elementi nella comboBox2 -> Inserisco i tipi di suoni.
            comboBox2.Items.Insert(0, "Arpa");
            comboBox2.Items.Insert(1, "Batteria");
            comboBox2.Items.Insert(2, "Chitarra");
            comboBox2.Items.Insert(3, "Pianoforte");
            // Imposto suono di default.
            comboBox2.SelectedItem = "Arpa";
        }



        // Dichiaro una stringa per contenere la variabile porta, e in caso di porta selezionata,
        // chiamo la funzione "connect" per connettermi alla relativa porta.
        string com;
        // ---> Al clic del bottone "Connect to Arduino".
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("Selezionare prima la porta seriale!", "Errore...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                button1.Enabled = false;
                button2.Enabled = true;
                com = comboBox1.Text.ToString();
                connect(com);
            }
        }



        // Con questa funzione assegno le proprietà all'oggetto stanziato prima -> SerialPort.
        // Chiamo il Thread "myport_DataReceived" per ricevere i messaggi.
        // Infine, apro la porta seriale per l'ascolto.
        void connect(string nome_porta)
        {
            sp = new SerialPort(nome_porta, 9600, Parity.None, 8, StopBits.One);
            // DTR -> Indica che l'oggetto sulla porta seriale è pronto sia a ricevere sia a scrivere.
            sp.DtrEnable = true;
            sp.DataReceived += new SerialDataReceivedEventHandler(myport_DataReceived);

            // Provo ad aprire la porta...
            try
            {
                // Apro la porta.
                sp.Open();

                richTextBox1.AppendText("[" + time + "] " + "CONNESSO!\n");
                comboBox1.Enabled = false;
                comboBox2.Enabled = true;
                label2.Enabled = true;
                label2.Text = "Connesso alla porta " + nome_porta;
                label4.Text = "CONNESSO";
                ovalShape1.BackColor = Color.Green;
            }
            // ... in caso negativo restituisco errore.
            catch
            {
                MessageBox.Show("Impossibile stabilire una connessione con quella porta!", "Errore...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        // Con readTo() leggo ogni linea che mi viene mandata dall'Arduino e la salvo in una stringa. LEGGO FINO AL CARATTERE DI FINE RIGA.
        // Copio la stringa nella TextBox per leggere il messaggio.
        void myport_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string msg = sp.ReadTo("\r\n");
            richTextBox1.AppendText("[" + time + "] " + msg + "\r\n");

            // A questo punto chiamo la funzione "sound", con la quale gestisco i suoni.
            sound(msg);
        }



        // Passo come parametro il valore della stringa in modo da gestirlo in questa funzione.
        // In base al suono selezionato e alla nota inviata dall'Arduino, genero un suono diverso.
        void sound(string sound)
        {
            // Se il suono selezionato è quello dell'arpa...
            if (String.Equals(comboBox2.SelectedItem, "Arpa"))
            {
                if (String.Equals(sound, "Nota A"))
                {
                    wplayer.URL = "Sounds/Harp/A.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota B"))
                {
                    wplayer.URL = "Sounds/Harp/B.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota C"))
                {
                    wplayer.URL = "Sounds/Harp/C.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota D"))
                {
                    wplayer.URL = "Sounds/Harp/D.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota E"))
                {
                    wplayer.URL = "Sounds/Harp/E.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota F"))
                {
                    wplayer.URL = "Sounds/Harp/F.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota G"))
                {
                    wplayer.URL = "Sounds/Harp/G.mp3";
                    wplayer.controls.play();
                }
            }

            // Se il suono selezionato è quello della batteria...
            if (String.Equals(comboBox2.SelectedItem, "Batteria"))
            {
                if (String.Equals(sound, "Nota A"))
                {
                    wplayer.URL = "Sounds/Drums/A.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota B"))
                {
                    wplayer.URL = "Sounds/Drums/B.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota C"))
                {
                    wplayer.URL = "Sounds/Drums/C.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota D"))
                {
                    wplayer.URL = "Sounds/Drums/D.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota E"))
                {
                    wplayer.URL = "Sounds/Drums/E.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota F"))
                {
                    wplayer.URL = "Sounds/Drums/F.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota G"))
                {
                    wplayer.URL = "Sounds/Drums/G.mp3";
                    wplayer.controls.play();
                }
            }

            // Se il suono selezionato è quello della chitarra...
            if (String.Equals(comboBox2.SelectedItem, "Chitarra"))
            {
                if (String.Equals(sound, "Nota A"))
                {
                    wplayer.URL = "Sounds/Guitar/A.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota B"))
                {
                    wplayer.URL = "Sounds/Guitar/B.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota C"))
                {
                    wplayer.URL = "Sounds/Guitar/C.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota D"))
                {
                    wplayer.URL = "Sounds/Guitar/D.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota E"))
                {
                    wplayer.URL = "Sounds/Guitar/E.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota F"))
                {
                    wplayer.URL = "Sounds/Guitar/F.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota G"))
                {
                    wplayer.URL = "Sounds/Guitar/G.mp3";
                    wplayer.controls.play();
                }
            }

            // Se il suono selezionato è quello di un pianoforte...
            if (String.Equals(comboBox2.SelectedItem, "Pianoforte"))
            {
                if (String.Equals(sound, "Nota A"))
                {
                    wplayer.URL = "Sounds/Piano/A.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota B"))
                {
                    wplayer.URL = "Sounds/Piano/B.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota C"))
                {
                    wplayer.URL = "Sounds/Piano/C.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota D"))
                {
                    wplayer.URL = "Sounds/Piano/D.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota E"))
                {
                    wplayer.URL = "Sounds/Piano/E.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota F"))
                {
                    wplayer.URL = "Sounds/Piano/F.mp3";
                    wplayer.controls.play();
                }
                if (String.Equals(sound, "Nota G"))
                {
                    wplayer.URL = "Sounds/Piano/G.mp3";
                    wplayer.controls.play();
                }
            }
        }
        

        // Chiudo la porta aperta in precedenza e termino di conseguenza la connessione con la seriale.
        // Inoltre attivo/disattivo gli oggetti in modo da ricreare la situazione iniziale in apertura di programma.
        // ---> Al clic del bottone "STOP".
        private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.Enabled = true;
            comboBox2.Enabled = false;
            button1.Enabled = true;
            button2.Enabled = false;
            label2.Text = "";
            label4.Text = "DISCONNESSO";
            ovalShape1.BackColor = Color.Red;

            sp.Close();

            richTextBox1.AppendText("[" + time + "] " + "DISCONNESSO!\n");
        }



        // ComboBox che contiene le porte seriali disponili.
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Tutte le impostazioni sono gia' state definite precedentemente.
        }



        // RichTextBox che contiene tutti i comandi e messaggi scambiati.
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
