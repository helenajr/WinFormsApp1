﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Security;

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

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e) //I'm not sure this function is doing anything, but the app won't work if I remove it.
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
           this.listBox1.Items.Clear();// This clears the list box
           if (openFileDialog1.ShowDialog() == DialogResult.OK) // This Opens the file dialog
            {
                try
                {
                    var filePath = openFileDialog1.FileName; // var tells the computer to figure out variable type for itself
                    
                    {
                        StreamReader reader = new StreamReader(filePath); // This chooses the file to read and saves in reader
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
                }
                catch (SecurityException ex) // I don't know what this security exception is about, but I've left it in
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e) //Creates a new file and puts data from old file into it, overwrites existing file, rather than appends stuff to it.
        {
            this.listBox1.Items.Clear();// This clears the list box
            if (openFileDialog1.ShowDialog() == DialogResult.OK) // This Opens the file dialog
            {
                try // This try...catch handles a security exception
                {
                    var filePath = openFileDialog1.FileName; // Initialises filepath varible. var tells the computer to figure out variable type for itself
                    string fileName = Path.GetFileName(filePath);
                    {
                        StreamReader reader = new StreamReader(filePath); // This chooses the file to read and saves in reader
                        StreamWriter writer = new StreamWriter("C:/Users/HRo/Desktop/TestFiles/NewFiles/Copy_" + fileName); //Creates new file and saves it in writer
                        try
                        {
                            do
                            {
                                writer.WriteLine(reader.ReadLine()); // Writes the content of old file into new file line-by-line
                            }
                            while (reader.Peek() != -1); // Think this looks to see if there is  next line
                        }
                        catch
                        {
                            addListItem("File is empty"); // What to display if there is nothing in the file
                        }
                        finally
                        {
                            reader.Close(); // Stops reading
                            writer.Close(); // Stops writing
                            addListItem("File contents copied to new file in NewFiles folder"); //Message saying what has happened
                        }
                    }
                }
                catch (SecurityException ex) // I don't know what this security exception is about, but I've left it in
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }

        //For more about writing to files see https://docs.microsoft.com/en-us/dotnet/standard/io/how-to-write-text-to-a-file
        
    }
}
