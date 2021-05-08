using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        string winDir = System.Environment.GetEnvironmentVariable("windir");
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //How to read a text file.
            //try...catch is to deal with a 0 byte file.
            this.listBox1.Items.Clear();// This clears the listbox
            StreamReader reader = new StreamReader("C:/Users/HRo/Desktop/TestFiles/KBTest.txt"); // This chooses the file to read and saves in reader
            try
            {
                do
                {
                    addListItem(reader.ReadLine()); // Reads the whole content of the file
                }
                while (reader.Peek() != -1); // No ideas what this condition means
            }
            catch
            {
                addListItem("File is empty"); // What to display if there is nothing in the file
            }
            finally
            {
                reader.Close(); // Stops reading
            }
            
        }
        private void addListItem(string value)
        {
            this.listBox1.Items.Add(value); // Things added to this list are displayed in the listbox
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Demonstrates how to create and write to a text file.
            StreamWriter writer = new StreamWriter("C:/Users/HRo/Desktop/TestFiles/KBTest.txt"); //This creates the file at this filepath
            writer.WriteLine("File created using StreamWriter class."); //This adds content to the file
            writer.Close(); //This stops adding content to the file
            this.listBox1.Items.Clear(); //This clears the listbox
            addListItem("File Written to C:/Users/HRo/Desktop/TestFiles/KBTest.txt"); //This is the message displayed in the listbox
        }
    }
}
