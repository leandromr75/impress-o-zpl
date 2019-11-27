using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Etiquetas_White_Martins;

namespace Eiquetas_White_Martins
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        int cont_par = 0;
        int cont_impar = 0;
        string impressao = "";

        int loteParAtual = 0;
        int loteImparAtual = 0;

        string manual = "sim";

       // int carreiras = 0;

        string carreiras_par = "";
        string carreiras_impar = "";

        string temp_par = "";
        string temp_impar = "";
        string Lote = "";
        //margens

        int margem1 = 208;
        int margem2 = 223;
        int margem3 = 100;
        int margem4 = 99;
        int margem5 = 25;
        int margem6 = 31;
        int margem7 = 495;
        int margem8 = 510;
        int margem9 = 387;
        int margem10 = 386;
        int margem11 = 312;
        int margem12 = 318;
        int margem13 = 782;
        int margem14 = 797;
        int margem15 = 674;
        int margem16 = 673;
        int margem17 = 599;
        int margem18 = 605;
        //-------------------
        int margemPonto1 = 109;
        int margemPonto2 = 395;
        int margemPonto3 = 680;

        int alturaPonto1 = 48;
        int alturaPonto2 = 48;
        int alturaPonto3 = 48;

        //-------------------
        int altura1 = 50;
        int altura2 = 106;
        int altura3 = 106;
        int altura4 = 53;
        int altura5 = 45;
        int altura6 = 100;
        int altura7 = 50;
        int altura8 = 106;
        int altura9 = 106;
        int altura10 = 53;
        int altura11 = 45;
        int altura12 = 100;
        int altura13 = 50;
        int altura14 = 106;
        int altura15 = 106;
        int altura16 = 53;
        int altura17 = 45;
        int altura18 = 100;

        public void diminui_altura()
        {
            altura1 += 2;
            altura2 += 2;
            altura3 += 2;
            altura4 += 2;
            altura5 += 2;
            altura6 += 2;
            altura7 += 2;
            altura8 += 2;
            altura9 += 2;
            altura10 += 2;
            altura11 += 2;
            altura12 += 2;
            altura13 += 2;
            altura14 += 2;
            altura15 += 2;
            altura16 += 2;
            altura17 += 2;
            altura18 += 2;

            alturaPonto1 += 2;
            alturaPonto2 += 2;
            alturaPonto2 += 2;

        }
        public void aumenta_altura()
        {
            altura1 -= 2;
            altura2 -= 2;
            altura3 -= 2;
            altura4 -= 2;
            altura5 -= 2;
            altura6 -= 2;
            altura7 -= 2;
            altura8 -= 2;
            altura9 -= 2;
            altura10 -= 2;
            altura11 -= 2;
            altura12 -= 2;
            altura13 -= 2;
            altura14 -= 2;
            altura15 -= 2;
            altura16 -= 2;
            altura17 -= 2;
            altura18 -= 2;

            alturaPonto1 -= 2;
            alturaPonto2 -= 2;
            alturaPonto2 -= 2;

        }
        public void diminui_margem()
        {
            margem1 -= 2;
            margem2 -= 2;
            margem3 -= 2;
            margem4 -= 2;
            margem5 -= 2;
            margem6 -= 2;
            margem7 -= 2;
            margem8 -= 2;
            margem9 -= 2;
            margem10 -= 2;
            margem11 -= 2;
            margem12 -= 2;
            margem13 -= 2;
            margem14 -= 2;
            margem15 -= 2;
            margem16 -= 2;
            margem17 -= 2;
            margem18 -= 2;

            margemPonto1 -= 2;
            margemPonto2 -= 2;
            margemPonto3 -= 2;

        }
        public void aumenta_margem()
        {
            margem1 += 2;
            margem2 += 2;
            margem3 += 2;
            margem4 += 2;
            margem5 += 2;
            margem6 += 2;
            margem7 += 2;
            margem8 += 2;
            margem9 += 2;
            margem10 += 2;
            margem11 += 2;
            margem12 += 2;
            margem13 += 2;
            margem14 += 2;
            margem15 += 2;
            margem16 += 2;
            margem17 += 2;
            margem18 += 2;

            margemPonto1 += 2;
            margemPonto2 += 2;
            margemPonto3 += 2;

        }
        private void button1_Click(object sender, EventArgs e)
        {
           
            timer1.Stop();
            pictureBox3.Visible = false;
            pictureBox4.Visible = true;
            button4.Visible = true;
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            comboBox7.Text = "21";
            comboBox8.Text = "21";

            comboBox1.Text = "1";
            comboBox2.Text = "3";

            comboBox3.Text = "3";
            comboBox4.Text = "1";

            comboBox5.Text = "LPT1";
            comboBox6.Text = "LPT1";

            cont_par = Convert.ToInt32(textBox5.Text);
            cont_impar = Convert.ToInt32(textBox6.Text);

            loteParAtual = Convert.ToInt32(textBox1.Text);
            loteImparAtual = Convert.ToInt32(textBox4.Text);

            temp_par = comboBox7.Text;
            temp_impar = comboBox8.Text;

            

            if (String.IsNullOrEmpty(comboBox1.Text) == false && String.IsNullOrEmpty(comboBox4.Text) == false)
            {
                carreiras_par = comboBox1.Text;
                carreiras_impar = comboBox4.Text;
            }
            else
            {
                MessageBox.Show("Selecione a quantidade de linhas",
                                    "Contador", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);  
            }

        }
        
        int temp = 0;

        private void timer1_Tick(object sender, EventArgs e)
        {
            label16.Text = "Lote O";
            //progress
            // progressBar1.Step = Convert.ToInt32(100 / Convert.ToInt32(textBox5.Text));

            //timer1.Start();
            //impressao = "par";
            //pictureBox3.Visible = true;
            carreiras_par = comboBox1.Text;
            carreiras_impar = comboBox2.Text;

            temp_par = comboBox7.Text;
            temp_impar = comboBox8.Text;

            //label9.Text = "Imprimindo Lote Nº " + Convert.ToString(loteParAtual);
            label18.Text = Convert.ToString(loteParAtual);
            textBox1.Text = Convert.ToString(loteParAtual);

            Lote = Convert.ToString(loteParAtual);
            if (Convert.ToString(loteParAtual).Length == 3)
            {
                Lote = "0" + loteParAtual;
            }
            if (Convert.ToString(loteParAtual).Length == 2)
            {
                Lote = "00" + loteParAtual;
            }
            if (Convert.ToString(loteParAtual).Length == 1)
            {
                Lote = "000" + loteParAtual;
            }
            string command = "CT~~CD,~CC^~CT~" +
                             "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR6,6~SD" + temp_par + "^JUS^LRN^CI0^XZ" +
                             "^XA" +
                             "^MMR" +
                             "^PW831" +
                             "^LL0120" +
                             "^LS0" +
                             "^FT" + Convert.ToString(margem1) + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                             "^FT" + Convert.ToString(margem2) + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                             "^FT" + Convert.ToString(margem3) + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                             "^FT" + Convert.ToString(margem4) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                             "^FT" + Convert.ToString(margem5) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                             "^FT" + Convert.ToString(margem6) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +

                            "^FT" + Convert.ToString(margem7) + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                             "^FT" + Convert.ToString(margem8) + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                             "^FT" + Convert.ToString(margem9) + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                             "^FT" + Convert.ToString(margem10) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                             "^FT" + Convert.ToString(margem11) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                             "^FT" + Convert.ToString(margem12) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +

                             "^FT" + Convert.ToString(margem13) + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                             "^FT" + Convert.ToString(margem14) + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                             "^FT" + Convert.ToString(margem15) + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                             "^FT" + Convert.ToString(margem16) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                             "^FT" + Convert.ToString(margem17) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                             "^FT" + Convert.ToString(margem18) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +
                             "^PQ28,0,1,Y^XZ";
            // Create a buffer with the command
            Byte[] buffer = new byte[command.Length];
            buffer = System.Text.Encoding.ASCII.GetBytes(command);
            // Use the CreateFile external func to connect to the LPT1 port
            SafeFileHandle printer = CreateFile("LPT1:", FileAccess.ReadWrite, 0, IntPtr.Zero, FileMode.Open, 0, IntPtr.Zero);
            // Aqui verifico se a impressora é válida
            if (printer.IsInvalid == true)
            {
                return;
            }

            // Open the filestream to the lpt1 port and send the command
            FileStream lpt1 = new FileStream(printer, FileAccess.ReadWrite);

            for (int i = 0; i < Convert.ToInt32(comboBox1.Text); i++)
            {
                lpt1.Write(buffer, 0, buffer.Length); //1
            }


            // Close the FileStream connection
            lpt1.Close();
            //label9.Text = "Imprimindo Lote Nº " + Convert.ToString(loteParAtual);
            label18.Text = Convert.ToString(loteParAtual);
            //imprime separação


            timer1.Stop();
            //Confirma Impressão
            Form2 f2 = new Form2(temp_par, loteParAtual, manual);
            f2.ShowDialog();

            //atualiza valores
            temp = Convert.ToInt32(textBox5.Text) - 1;
            textBox5.Text = Convert.ToString(temp);
            loteParAtual = loteParAtual - 1;


            // progressBar1.PerformStep();


            /*if (temp <= 0)
            {
                textBox1.Text = Convert.ToString(loteParAtual);
                //timer1.Stop();
                MessageBox.Show("Fim da sequência de impressão",
                                "Contador", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                //pictureBox3.Visible = false;
                //pictureBox4.Visible = true;

                //progressBar1.Value = 0;
            }*/
            //timer1.Start();
            if (loteParAtual < Convert.ToInt32(textBox2.Text))
            {
                //timer1.Stop();
                MessageBox.Show("Fim da sequência de Lotes",
                                "Contador", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                timer1.Enabled = false;
                timer1.Stop();
                pictureBox5.Visible = true;
                pictureBox6.Visible = false;
                checkBox5.CheckState = CheckState.Unchecked;
                return;
                //pictureBox3.Visible = false;
                //pictureBox4.Visible = false;

                //progressBar1.Value = 0;
            }
            else
            {
                timer1.Start();
            }
           
            
            
        }

        private void sobreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Sistema para imprimir etiquetas com Lote intercaladas,\nespecífico para o cliente White Martins\n\nArtchik © 2017 [By LEX]",
            "Etiquetas", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            label16.Text = "Lote O";
            groupBox2.Visible = false;
            checkBox4.CheckState = CheckState.Checked;
            manual = "sim";
            checkBox5.CheckState = CheckState.Unchecked;
            //progress
           // progressBar1.Step = Convert.ToInt32(100 / Convert.ToInt32(textBox5.Text));
            
            //timer1.Start();
            //impressao = "par";
            //pictureBox3.Visible = true;
            carreiras_par = comboBox1.Text;
            carreiras_impar = comboBox2.Text;

            temp_par = comboBox7.Text;
            temp_impar = comboBox8.Text;

            //label9.Text = "Imprimindo Lote Nº " + Convert.ToString(loteParAtual);
            label18.Text = Convert.ToString(loteParAtual);
            textBox1.Text = Convert.ToString(loteParAtual);
            
            Lote = Convert.ToString(loteParAtual);
            if (Convert.ToString(loteParAtual).Length == 3)
            {
                Lote = "0" + loteParAtual;
            }
            if (Convert.ToString(loteParAtual).Length == 2)
            {
                Lote = "00" + loteParAtual;
            }
            if (Convert.ToString(loteParAtual).Length == 1)
            {
                Lote = "000" + loteParAtual;
            }
            string command = "CT~~CD,~CC^~CT~" +
                            "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR6,6~SD" + temp_par + "^JUS^LRN^CI0^XZ" +
                            "^XA" +
                            "^MMR" +
                            "^PW831" +
                            "^LL0120" +
                            "^LS0" +
                            "^FT" + Convert.ToString(margem1)  + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem2)  + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem3)  + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem4)  + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +                            
                            "^FT" + Convert.ToString(margem5)  + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem6)  + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +

                           "^FT" + Convert.ToString(margem7)   + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem8)  + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem9)  + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem10) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT" + Convert.ToString(margem11) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem12) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +

                            "^FT" + Convert.ToString(margem13) + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem14) + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem15) + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem16) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT" + Convert.ToString(margem17) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem18) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" + 
                            "^PQ28,0,1,Y^XZ";
            // Create a buffer with the command
            Byte[] buffer = new byte[command.Length];
            buffer = System.Text.Encoding.ASCII.GetBytes(command);
            // Use the CreateFile external func to connect to the LPT1 port
            SafeFileHandle printer = CreateFile("LPT1:", FileAccess.ReadWrite, 0, IntPtr.Zero, FileMode.Open, 0, IntPtr.Zero);
            // Aqui verifico se a impressora é válida
            if (printer.IsInvalid == true)
            {
                return;
            }

            // Open the filestream to the lpt1 port and send the command
            FileStream lpt1 = new FileStream(printer, FileAccess.ReadWrite);

            for (int i = 0; i < Convert.ToInt32(comboBox1.Text); i++)
            {
                lpt1.Write(buffer, 0, buffer.Length); //1
            }          


            // Close the FileStream connection
            lpt1.Close();
            //label9.Text = "Imprimindo Lote Nº " + Convert.ToString(loteParAtual);
            label18.Text = Convert.ToString(loteParAtual);
            //imprime separação

            

            //Confirma Impressão
            Form2 f2 = new Form2(temp_par,loteParAtual,manual);
            f2.ShowDialog();
        
            //atualiza valores
            temp = Convert.ToInt32(textBox5.Text) - 1;
            textBox5.Text = Convert.ToString(temp);
            loteParAtual = loteParAtual - 1;


           // progressBar1.PerformStep();


            /*if (temp <= 0)
            {
                textBox1.Text = Convert.ToString(loteParAtual);
                //timer1.Stop();
                MessageBox.Show("Fim da sequência de impressão",
                                "Contador", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                //pictureBox3.Visible = false;
                //pictureBox4.Visible = true;

                //progressBar1.Value = 0;
            }*/
            if (loteParAtual < Convert.ToInt32(textBox2.Text))
            {
                //timer1.Stop();
                MessageBox.Show("Fim da sequência de Lotes",
                                "Contador", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                //pictureBox3.Visible = false;
                //pictureBox4.Visible = false;

                //progressBar1.Value = 0;
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            label16.Text = "Lote I";
            //progress
            //progressBar1.Step = Convert.ToInt32(100 / Convert.ToInt32(textBox6.Text));

            //timer1.Start();
            //impressao = "impar";
            //pictureBox3.Visible = true;
            groupBox1.Visible = false;
            checkBox3.CheckState = CheckState.Checked;

            manual = "sim";
            checkBox6.CheckState = CheckState.Unchecked;
            
            carreiras_par = comboBox1.Text;
            carreiras_impar = comboBox2.Text;

            temp_par = comboBox7.Text;
            temp_impar = comboBox8.Text;

            //label9.Text = "Imprimindo Lote Nº " + Convert.ToString(loteImparAtual);
            label18.Text = Convert.ToString(loteImparAtual);
            textBox4.Text = Convert.ToString(loteImparAtual);
            Lote = Convert.ToString( loteImparAtual );
            if (Convert.ToString(loteImparAtual).Length == 3)
            {
                Lote = "0" + loteImparAtual;
            }
            if (Convert.ToString(loteImparAtual).Length == 2)
            {
                Lote = "00" + loteImparAtual;
            }
            if (Convert.ToString(loteImparAtual).Length == 1)
            {
                Lote = "000" + loteImparAtual;
            }
            // Command to be sent to the printer
            string command = "CT~~CD,~CC^~CT~" +
                            "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR6,6~SD" + temp_impar + "^JUS^LRN^CI0^XZ" +
                             "^XA" +
                            "^MMR" +
                            "^PW831" +
                            "^LL0120" +
                            "^LS0" +
                            "^FT" + Convert.ToString(margem1)  + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem2)  + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem3)  + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem4)  + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT" + Convert.ToString(margem5)  + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem6)  + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +

                           "^FT" + Convert.ToString(margem7)   + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem8)  + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem9)  + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem10) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT" + Convert.ToString(margem11) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem12) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +

                            "^FT" + Convert.ToString(margem13) + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem14) + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem15) + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem16) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT" + Convert.ToString(margem17) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem18) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" + 
                            "^PQ16,0,1,Y^XZ";
            // Create a buffer with the command
            Byte[] buffer = new byte[command.Length];
            buffer = System.Text.Encoding.ASCII.GetBytes(command);
            // Use the CreateFile external func to connect to the LPT1 port
            SafeFileHandle printer = CreateFile("LPT1:", FileAccess.ReadWrite, 0, IntPtr.Zero, FileMode.Open, 0, IntPtr.Zero);
            // Aqui verifico se a impressora é válida
            if (printer.IsInvalid == true)
            {
                return;
            }

            // Open the filestream to the lpt1 port and send the command
            FileStream lpt1 = new FileStream(printer, FileAccess.ReadWrite);

            for (int i = 0; i < Convert.ToInt32(comboBox4.Text); i++)
            {
                lpt1.Write(buffer, 0, buffer.Length); //1
            }
            
            /*lpt1.Write(buffer, 0, buffer.Length); //1
            lpt1.Write(buffer, 0, buffer.Length); //2
            lpt1.Write(buffer, 0, buffer.Length); //3
            lpt1.Write(buffer, 0, buffer.Length); //4

            lpt1.Write(buffer, 0, buffer.Length); //5
            lpt1.Write(buffer, 0, buffer.Length); //6
            lpt1.Write(buffer, 0, buffer.Length); //7
            lpt1.Write(buffer, 0, buffer.Length); //8

            lpt1.Write(buffer, 0, buffer.Length); //9
            lpt1.Write(buffer, 0, buffer.Length); //10
            lpt1.Write(buffer, 0, buffer.Length); //11
            lpt1.Write(buffer, 0, buffer.Length); //12

            lpt1.Write(buffer, 0, buffer.Length); //13
            lpt1.Write(buffer, 0, buffer.Length); //14
            lpt1.Write(buffer, 0, buffer.Length); //15
            lpt1.Write(buffer, 0, buffer.Length); //16*/

            // Close the FileStream connection
            lpt1.Close();
            //label9.Text = "Imprimindo Lote Nº " + Convert.ToString(loteImparAtual);
            label18.Text = Convert.ToString(loteImparAtual);
            //imprime separação

            

            //ok
            Form2 f3 = new Form2(temp_impar,loteImparAtual, manual);
            f3.ShowDialog();

            //atualiza valores
            temp = Convert.ToInt32(textBox6.Text) - 1;
            textBox6.Text = Convert.ToString(temp);
            loteImparAtual = loteImparAtual - 1;


            //progressBar1.PerformStep();

            /*if (temp <= 0)
            {
                textBox4.Text = Convert.ToString(loteImparAtual);
                //timer1.Stop();
                MessageBox.Show("Fim da sequência de impressão",
                                "Contador", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                //pictureBox3.Visible = false;
                //pictureBox4.Visible = true;

                //progressBar1.Value = 0;
            }*/
            if (loteImparAtual < Convert.ToInt32(textBox3.Text))
            {
                //timer1.Stop();
                MessageBox.Show("Fim da sequência de Lotes",
                                "Contador", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                //pictureBox3.Visible = false;
                //pictureBox4.Visible = false;

                //progressBar1.Value = 0;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            /*if (Convert.ToInt32(textBox1.Text) % 2 > 0)
            {
                MessageBox.Show("Favor Digite Nº PAR");
                textBox1.Focus();
                return;
            }
            if (Convert.ToInt32(textBox2.Text) % 2 > 0)
            {
                MessageBox.Show("Favor Digite Nº PAR");
                textBox2.Focus();
                return;
            }*/
            if (Convert.ToInt32(textBox1.Text) <= Convert.ToInt32(textBox2.Text))
            {
                MessageBox.Show("Intervalo Inválido");
                return;
            }
            loteParAtual = Convert.ToInt32(textBox1.Text);
            MessageBox.Show("Gravado na mémoria novo lote inicial",
            "Lote Inicial", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void pictureBox1_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            /*if (Convert.ToInt32(textBox4.Text) % 2 == 0)
            {
                MessageBox.Show("Favor Digite Nº ÍMPAR");
                textBox4.Focus();
                return;
            }
            if (Convert.ToInt32(textBox3.Text) % 2 == 0)
            {
                MessageBox.Show("Favor Digite Nº ÍMPAR");
                textBox3.Focus();
                return;
            }*/
            if (Convert.ToInt32(textBox4.Text) <= Convert.ToInt32(textBox3.Text))
            {
                MessageBox.Show("Intervalo Inválido");
                return;
            }
            loteImparAtual = Convert.ToInt32(textBox4.Text);
            MessageBox.Show("Gravado na mémoria novo lote inicial",
            "Lote Inicial", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                groupBox1.Visible = false;
                label16.Text = "Lote I";
            }
            if (checkBox3.Checked == false)
            {
                groupBox1.Visible = true;
                label16.Text = "Lote X";
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                groupBox2.Visible = false;
                label16.Text = "Lote O";
            }
            if (checkBox4.Checked == false)
            {
                groupBox2.Visible = true;
                label16.Text = "Lote X";
            }
        }
        string utexto = "";
        private void label16_Click(object sender, EventArgs e)
        {
            textBox7.Text = label16.Text;
            utexto = "label16"; 
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (utexto == "label16")
            {
                label16.Text = textBox7.Text;
            }
            if (utexto == "label17")
            {
                label17.Text = textBox7.Text;
            }
            if (utexto == "label19")
            {
                label19.Text = textBox7.Text;
            }
            if (utexto == "label20")
            {
                label20.Text = textBox7.Text;
            }
            if (utexto == "label21")
            {
                label21.Text = textBox7.Text;
            }
        }

        private void label17_Click(object sender, EventArgs e)
        {
            textBox7.Text = label17.Text;
            utexto = "label17"; 
        }

        private void label19_Click(object sender, EventArgs e)
        {
            textBox7.Text = label19.Text;
            utexto = "label19"; 
        }

        private void label20_Click(object sender, EventArgs e)
        {
            textBox7.Text = label20.Text;
            utexto = "label20"; 
        }

        private void button7_Click(object sender, EventArgs e)
        {
            aumenta_margem();
            label16.Location = new Point(label16.Location.X + 2, label16.Location.Y);
            label17.Location = new Point(label17.Location.X + 2, label17.Location.Y);
            label18.Location = new Point(label18.Location.X + 2, label18.Location.Y);
            label19.Location = new Point(label19.Location.X + 2, label19.Location.Y);
            label20.Location = new Point(label20.Location.X + 2, label20.Location.Y);
            label21.Location = new Point(label21.Location.X + 2, label21.Location.Y);
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            diminui_margem();
            label16.Location = new Point(label16.Location.X - 2, label16.Location.Y);
            label17.Location = new Point(label17.Location.X - 2, label17.Location.Y);
            label18.Location = new Point(label18.Location.X - 2, label18.Location.Y);
            label19.Location = new Point(label19.Location.X - 2, label19.Location.Y);
            label20.Location = new Point(label20.Location.X - 2, label20.Location.Y);
            label21.Location = new Point(label21.Location.X - 2, label21.Location.Y);
            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            aumenta_altura();
            label16.Location = new Point(label16.Location.X, label16.Location.Y - 2);
            label17.Location = new Point(label17.Location.X, label17.Location.Y - 2);
            label18.Location = new Point(label18.Location.X, label18.Location.Y - 2);
            label19.Location = new Point(label19.Location.X, label19.Location.Y - 2);
            label20.Location = new Point(label20.Location.X, label20.Location.Y - 2);
            label21.Location = new Point(label21.Location.X, label21.Location.Y - 2);
           
        }

        private void button9_Click(object sender, EventArgs e)
        {
            diminui_altura();
            label16.Location = new Point(label16.Location.X, label16.Location.Y + 2);
            label17.Location = new Point(label17.Location.X, label17.Location.Y + 2);
            label18.Location = new Point(label18.Location.X, label18.Location.Y + 2);
            label19.Location = new Point(label19.Location.X, label19.Location.Y + 2);
            label20.Location = new Point(label20.Location.X, label20.Location.Y + 2);
            label21.Location = new Point(label21.Location.X, label21.Location.Y + 2);
            
        }

        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            groupBox2.Visible = false;
            checkBox4.CheckState = CheckState.Checked;
            if (Convert.ToInt32(textBox1.Text) % 2 > 0)
            {
                MessageBox.Show("Favor Digite Nº PAR");
                textBox1.Focus();
                checkBox5.CheckState = CheckState.Unchecked;
                return;
            }
            if (Convert.ToInt32(textBox2.Text) % 2 > 0)
            {
                MessageBox.Show("Favor Digite Nº PAR");
                textBox2.Focus();
                checkBox5.CheckState = CheckState.Unchecked;
                return;
            }
            if (Convert.ToInt32(textBox1.Text) <= Convert.ToInt32(textBox2.Text))
            {
                MessageBox.Show("Intervalo Inválido");
                checkBox5.CheckState = CheckState.Unchecked;
                return;
            }
            if (checkBox5.Checked == true)
            {
                timer1.Enabled = true;
                pictureBox6.Visible = true;
                pictureBox5.Visible = false;
                manual = "não";
            }
            if (checkBox5.Checked == false)
            {
                pictureBox6.Visible = false;
                pictureBox5.Visible = true;
                manual = "sim";
                timer1.Enabled = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            checkBox3.CheckState = CheckState.Checked;
            if (Convert.ToInt32(textBox4.Text) % 2 == 0)
            {
                MessageBox.Show("Favor Digite Nº ÍMPAR");
                checkBox6.CheckState = CheckState.Unchecked;
                textBox4.Focus();
                return;
            }
            if (Convert.ToInt32(textBox3.Text) % 2 == 0)
            {
                MessageBox.Show("Favor Digite Nº ÍMPAR");
                checkBox6.CheckState = CheckState.Unchecked;
                textBox3.Focus();
                return;
            }
            if (Convert.ToInt32(textBox4.Text) <= Convert.ToInt32(textBox3.Text))
            {
                MessageBox.Show("Intervalo Inválido");
                checkBox6.CheckState = CheckState.Unchecked;
                return;
            }
            if (checkBox6.Checked == true)
            {
                timer2.Enabled = true;
                pictureBox7.Visible = true;
                pictureBox8.Visible = false;
                manual = "não";
            }
            if (checkBox6.Checked == false)
            {
                timer2.Enabled = false;
                pictureBox7.Visible = false;
                pictureBox8.Visible = true;
                manual = "sim";
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            label16.Text = "Lote I";
            //progress
            //progressBar1.Step = Convert.ToInt32(100 / Convert.ToInt32(textBox6.Text));

            //timer1.Start();
            //impressao = "impar";
            //pictureBox3.Visible = true;
            carreiras_par = comboBox1.Text;
            carreiras_impar = comboBox2.Text;

            temp_par = comboBox7.Text;
            temp_impar = comboBox8.Text;

            //label9.Text = "Imprimindo Lote Nº " + Convert.ToString(loteImparAtual);
            label18.Text = Convert.ToString(loteImparAtual);
            textBox4.Text = Convert.ToString(loteImparAtual);
            Lote = Convert.ToString(loteImparAtual);
            if (Convert.ToString(loteImparAtual).Length == 3)
            {
                Lote = "0" + loteImparAtual;
            }
            if (Convert.ToString(loteImparAtual).Length == 2)
            {
                Lote = "00" + loteImparAtual;
            }
            if (Convert.ToString(loteImparAtual).Length == 1)
            {
                Lote = "000" + loteImparAtual;
            }
            // Command to be sent to the printer
            string command = "CT~~CD,~CC^~CT~" +
                            "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR6,6~SD" + temp_impar + "^JUS^LRN^CI0^XZ" +
                             "^XA" +
                            "^MMR" +
                            "^PW831" +
                            "^LL0120" +
                            "^LS0" +
                            "^FT" + Convert.ToString(margem1) + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem2) + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem3) + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem4) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT" + Convert.ToString(margem5) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem6) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +

                            "^FT" + Convert.ToString(margem7) + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem8) + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem9) + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem10) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT" + Convert.ToString(margem11) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem12) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +

                            "^FT" + Convert.ToString(margem13) + "," + Convert.ToString(altura1) + "^A0N,39,38^FH\\^FD" + label17.Text + "^FS" +
                            "^FT" + Convert.ToString(margem14) + "," + Convert.ToString(altura2) + "^A0N,48,48^FH\\^FD" + label21.Text + "^FS" +
                            "^FT" + Convert.ToString(margem15) + "," + Convert.ToString(altura3) + "^A0N,48,48^FH\\^FD" + label20.Text + "^FS" +
                            "^FT" + Convert.ToString(margem16) + "," + Convert.ToString(altura4) + "^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT" + Convert.ToString(margem17) + "," + Convert.ToString(altura5) + "^A0N,25,24^FH\\^FD" + label16.Text + "^FS" +
                            "^FT" + Convert.ToString(margem18) + "," + Convert.ToString(altura6) + "^A0N,31,31^FH\\^FD" + label19.Text + "^FS" +
                            "^PQ16,0,1,Y^XZ";
            // Create a buffer with the command
            Byte[] buffer = new byte[command.Length];
            buffer = System.Text.Encoding.ASCII.GetBytes(command);
            // Use the CreateFile external func to connect to the LPT1 port
            SafeFileHandle printer = CreateFile("LPT1:", FileAccess.ReadWrite, 0, IntPtr.Zero, FileMode.Open, 0, IntPtr.Zero);
            // Aqui verifico se a impressora é válida
            if (printer.IsInvalid == true)
            {
                return;
            }

            // Open the filestream to the lpt1 port and send the command
            FileStream lpt1 = new FileStream(printer, FileAccess.ReadWrite);

            for (int i = 0; i < Convert.ToInt32(comboBox4.Text); i++)
            {
                lpt1.Write(buffer, 0, buffer.Length); //1
            }

            /*lpt1.Write(buffer, 0, buffer.Length); //1
            lpt1.Write(buffer, 0, buffer.Length); //2
            lpt1.Write(buffer, 0, buffer.Length); //3
            lpt1.Write(buffer, 0, buffer.Length); //4

            lpt1.Write(buffer, 0, buffer.Length); //5
            lpt1.Write(buffer, 0, buffer.Length); //6
            lpt1.Write(buffer, 0, buffer.Length); //7
            lpt1.Write(buffer, 0, buffer.Length); //8

            lpt1.Write(buffer, 0, buffer.Length); //9
            lpt1.Write(buffer, 0, buffer.Length); //10
            lpt1.Write(buffer, 0, buffer.Length); //11
            lpt1.Write(buffer, 0, buffer.Length); //12

            lpt1.Write(buffer, 0, buffer.Length); //13
            lpt1.Write(buffer, 0, buffer.Length); //14
            lpt1.Write(buffer, 0, buffer.Length); //15
            lpt1.Write(buffer, 0, buffer.Length); //16*/

            // Close the FileStream connection
            lpt1.Close();
            //label9.Text = "Imprimindo Lote Nº " + Convert.ToString(loteImparAtual);
            label18.Text = Convert.ToString(loteImparAtual);
            //imprime separação

            timer2.Stop();

            //ok
            Form2 f3 = new Form2(temp_impar, loteImparAtual, manual);
            f3.ShowDialog();

            //atualiza valores
            temp = Convert.ToInt32(textBox6.Text) - 1;
            textBox6.Text = Convert.ToString(temp);
            loteImparAtual = loteImparAtual - 1;


            //progressBar1.PerformStep();

            /*if (temp <= 0)
            {
                textBox4.Text = Convert.ToString(loteImparAtual);
                //timer1.Stop();
                MessageBox.Show("Fim da sequência de impressão",
                                "Contador", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                //pictureBox3.Visible = false;
                //pictureBox4.Visible = true;

                //progressBar1.Value = 0;
            }*/
            if (loteImparAtual < Convert.ToInt32(textBox3.Text))
            {
                //timer1.Stop();
                MessageBox.Show("Fim da sequência de Lotes",
                                "Contador", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                timer2.Enabled = false;
                timer2.Stop();
                pictureBox7.Visible = false;
                pictureBox8.Visible = true;
                checkBox6.CheckState = CheckState.Unchecked;
                return;
                //pictureBox3.Visible = false;
                //pictureBox4.Visible = false;

                //progressBar1.Value = 0;
            }
            else
            {
                timer2.Start();
            }
        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {
            textBox7.Text = label21.Text;
            utexto = "label21";
        }

       
    }
}
