using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using ProcessMemoryReaderLib;
using System.Runtime.InteropServices;

namespace Memory_Hacking
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region write_memory
        public void write()
        {
            try
            {
                ProcessMemoryReaderLib.ProcessMemoryReader preader = new ProcessMemoryReaderLib.ProcessMemoryReader();
                ProcessMemoryReader Reader = new ProcessMemoryReader();
                Process[] MyProcess = Process.GetProcessesByName(textBox1 .Text );
                preader.ReadProcess = MyProcess[0];
                preader.OpenProcess();

                int byteswritten;
                long value;
                byte[] memory;
                string wadress = textBox2.Text;
                IntPtr writeadress = new IntPtr(Convert.ToInt32(wadress, 16));

                value = Convert.ToInt64(textBox3 .Text );
                memory = BitConverter.GetBytes(value);
                preader.WriteProcessMemory((IntPtr)writeadress, memory, out byteswritten);

                
                label5.ForeColor = Color.Green;
                label5.Text = "Değerleriniz Aktarıldı!...";
 
            }

            catch
            {
                label5.ForeColor = Color.Red;
                label5.Text = "Hata Oluştu!...";
            }
        }
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (Process P in Process.GetProcesses())
            {
                if (P.MainWindowTitle.Length > 1)
                {
                    listBox1.Items.Add(P.MainWindowTitle);
                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            foreach (Process P in Process.GetProcesses())
            {
                if (P.MainWindowTitle.Length > 1)
                {
                    
                    listBox1.Items.Add(P.MainWindowTitle);
                }

            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
             string process = listBox1.Text;
             textBox1.Text = listBox1.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            write();
        }



    }
}
