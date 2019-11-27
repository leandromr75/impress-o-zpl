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

namespace Etiquetas_White_Martins
{
    public partial class Form2 : Form
    {
        string temp_par;
        int loteParAtual;
        string manu = "";
        public Form2(string temp, int lote, string Manual)
        {
            InitializeComponent();
            temp_par = temp;
            loteParAtual = lote;
            manu = Manual;
        }
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);


        private void button1_Click(object sender, EventArgs e)
        {
            // Command to be sent to the printer
            string command2 = "CT~~CD,~CC^~CT~" +
                            "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2~SD" + temp_par + "^JUS^LRN^CI0^XZ" +
                            "^XA" +
                            "^MMT" +
                            "^PW831" +
                            "^LL0120" +
                            "^LS0" +
                            "^FO56,54^GB162,0,8^FS" +
                            "^FO340,54^GB162,0,8^FS" +
                            "^FO623,54^GB162,0,8^FS" +
                            "^PQ1,0,1,Y^XZ";
            // Create a buffer with the command
            Byte[] buffer2 = new byte[command2.Length];
            buffer2 = System.Text.Encoding.ASCII.GetBytes(command2);
            // Use the CreateFile external func to connect to the LPT1 port
            SafeFileHandle printer2 = CreateFile("LPT1:", FileAccess.ReadWrite, 0, IntPtr.Zero, FileMode.Open, 0, IntPtr.Zero);
            // Aqui verifico se a impressora é válida
            if (printer2.IsInvalid == true)
            {
                return;
            }

            // Open the filestream to the lpt1 port and send the command
            FileStream lpt12 = new FileStream(printer2, FileAccess.ReadWrite);


            lpt12.Write(buffer2, 0, buffer2.Length);


            // Close the FileStream connection
            lpt12.Close();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
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
                            "^FT208,50^A0N,39,38^FH\\^FD/19^FS" +
                            "^FT223,106^A0N,48,48^FH\\^FD3^FS" +
                            "^FT100,106^A0N,48,48^FH\\^FD455^FS" +
                            "^FT99,53^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT25,45^A0N,25,24^FH\\^FDLote O^FS" + 
                            "^FT31,100^A0N,31,31^FH\\^FDUNP^FS" +

                            "^FT495,50^A0N,39,38^FH\\^FD/19^FS" +
                            "^FT510,106^A0N,48,48^FH\\^FD3^FS" +
                            "^FT387,106^A0N,48,48^FH\\^FD455^FS" +
                            "^FT386,53^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT312,45^A0N,25,24^FH\\^FDLote O^FS" +
                            "^FT318,100^A0N,31,31^FH\\^FDUNP^FS" +

                            "^FT782,50^A0N,39,38^FH\\^FD/19^FS" +
                            "^FT797,106^A0N,48,48^FH\\^FD3^FS" +
                            "^FT674,106^A0N,48,48^FH\\^FD455^FS" +
                            "^FT673,53^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT599,45^A0N,25,24^FH\\^FDLote O^FS" +
                            "^FT605,100^A0N,31,31^FH\\^FDUNP^FS" +
                            "^PQ1,0,1,Y^XZ";
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


            lpt1.Write(buffer, 0, buffer.Length); //1            


            // Close the FileStream connection
            lpt1.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
           
            button1.Focus();
            label2.Text = Convert.ToString(loteParAtual);
            if (manu == "não")
            {
                timer1.Enabled = true;
                label3.Visible = true;
                pictureBox7.Visible = true;
                button1.Visible = false;
                button2.Visible = false;
                button3.Visible = false;
            }
            if (manu == "sim")
            {
                timer1.Enabled = false;
                label3.Visible = false;
                pictureBox7.Visible = false;
                button1.Visible = true;
                button2.Visible = true;
                button3.Visible = true;
            }
        }
        string Lote = "";
        private void button3_Click(object sender, EventArgs e)
        {
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
                           "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR7,7~SD" + temp_par + "^JUS^LRN^CI0^XZ" +
                           "^XA" +
                           "^MMR" +
                           "^PW831" +
                           "^LL0120" +
                           "^LS0" +
                           "^FT208,50^A0N,39,38^FH\\^FD/19^FS" +
                            "^FT223,106^A0N,48,48^FH\\^FD3^FS" +
                            "^FT100,106^A0N,48,48^FH\\^FD455^FS" +
                            "^FT99,53^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT25,45^A0N,25,24^FH\\^FDLote I^FS" +
                            "^FT31,100^A0N,31,31^FH\\^FDUNP^FS" +

                            "^FT495,50^A0N,39,38^FH\\^FD/19^FS" +
                            "^FT510,106^A0N,48,48^FH\\^FD3^FS" +
                            "^FT387,106^A0N,48,48^FH\\^FD455^FS" +
                            "^FT386,53^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT312,45^A0N,25,24^FH\\^FDLote I^FS" +
                            "^FT318,100^A0N,31,31^FH\\^FDUNP^FS" +

                            "^FT782,50^A0N,39,38^FH\\^FD/19^FS" +
                            "^FT797,106^A0N,48,48^FH\\^FD3^FS" +
                            "^FT674,106^A0N,48,48^FH\\^FD455^FS" +
                            "^FT673,53^A0N,48,48^FH\\^FD" + Lote + "^FS" +
                            "^FT599,45^A0N,25,24^FH\\^FDLote I^FS" +
                            "^FT605,100^A0N,31,31^FH\\^FDUNP^FS" +
                           "^PQ1,0,1,Y^XZ";
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

            lpt1.Write(buffer, 0, buffer.Length); //1
            
            // Close the FileStream connection
            lpt1.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            // Command to be sent to the printer
            string command2 = "CT~~CD,~CC^~CT~" +
                            "^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2~SD" + temp_par + "^JUS^LRN^CI0^XZ" +
                            "^XA" +
                            "^MMT" +
                            "^PW831" +
                            "^LL0120" +
                            "^LS0" +
                            "^FO56,54^GB162,0,8^FS" +
                            "^FO340,54^GB162,0,8^FS" +
                            "^FO623,54^GB162,0,8^FS" +
                            "^PQ1,0,1,Y^XZ";
            // Create a buffer with the command
            Byte[] buffer2 = new byte[command2.Length];
            buffer2 = System.Text.Encoding.ASCII.GetBytes(command2);
            // Use the CreateFile external func to connect to the LPT1 port
            SafeFileHandle printer2 = CreateFile("LPT1:", FileAccess.ReadWrite, 0, IntPtr.Zero, FileMode.Open, 0, IntPtr.Zero);
            // Aqui verifico se a impressora é válida
            if (printer2.IsInvalid == true)
            {
                return;
            }

            // Open the filestream to the lpt1 port and send the command
            FileStream lpt12 = new FileStream(printer2, FileAccess.ReadWrite);


            lpt12.Write(buffer2, 0, buffer2.Length);


            // Close the FileStream connection
            lpt12.Close();
            timer1.Enabled = false;
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (pictureBox7.Visible == true)
            {
                timer1.Stop();
            }
        }
    }
}
