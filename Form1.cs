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
using System.Security;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


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
        //For the excel stuff below see https://docs.microsoft.com/en-US/previous-versions/office/troubleshoot/office-developer/automate-excel-from-visual-c


        private void button5_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;


            this.listBox1.Items.Clear();// This clears the list box
            if (openFileDialog1.ShowDialog() == DialogResult.OK) // This Opens the file dialog
            {
                try // This try...catch handles a security exception
                {
                    var filePath = openFileDialog1.FileName; // Initialises filepath varible. var tells the computer to figure out variable type for itself
                    string fileName = Path.GetFileName(filePath); // Gets the filename
                    try // Again, I do not know what this try....catch is for, but I've left it here
                    {
                        //Start Excel and get Application object.
                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.Visible = true;

                        //Open an existing workbook at the first sheet.
                        oWB = oXL.Workbooks.Open(filePath);
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                        //Apply a formula to get a thrid column and give column a name
                        oSheet.Cells[1, 3] = "Double number";
                        oRng = oSheet.get_Range("C2", "C8");
                        oRng.Formula = "=B2 * 2";


                        //Make sure Excel is visible and give the user control
                        //of Microsoft Excel's lifetime.
                        oXL.Visible = true;
                        oXL.UserControl = true;
                    }
                    catch (Exception theException)
                    {
                        String errorMessage;
                        errorMessage = "Error: ";
                        errorMessage = String.Concat(errorMessage, theException.Message);
                        errorMessage = String.Concat(errorMessage, " Line: ");
                        errorMessage = String.Concat(errorMessage, theException.Source);

                        MessageBox.Show(errorMessage, "Error");
                    }
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }

        }

        private void button6_Click(object sender, EventArgs e) //This button displays the content of cell A1 in the listbox
        {          
            DisplayCell();
        }

        public void DisplayCell()
        {

            string fp = ChooseFilePath(); // Uses function below to find file path
            Excel excel = new Excel(fp , 1); // Creates a new instance of the Excel class

            addListItem(excel.ReadCell(0, 0)); // Uses the excel class function ReadCell and displays result in the listbox
        }
        public string ChooseFilePath() // Custom function to use whenever you need to obtain file path through the dialog box
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var filePath = openFileDialog1.FileName; // Initialises filepath varible. var tells the computer to figure out variable type for itself
                return filePath;
            }
            else
            {
                return "Error"; //All code paths need a string return for the function to work
            }


        }
        public string ChooseFileName() // Custom function to use whenever you need to obtain file name through the dialog box
        {
            string filePath = ChooseFilePath();
            string fileName = Path.GetFileName(filePath); // Gets the filename
            return fileName;
        }

        public void WriteData()
        {
            string fp = ChooseFilePath(); // Uses function below to find file path
            Excel excel = new Excel(fp, 1); // Creates a new instance of the Excel class

            excel.WriteToCell(6, 6, "BOOM");           
            excel.SaveAs(@"BOOM3.xlsx"); //The @ symbol adds the file path This PC/ My Documents

            excel.Close();
            addListItem("File with new content added saved in ThisPC/My Documents/ BOOM3.xlsx"); //Message saying what has happened
        }

        private void button7_Click(object sender, EventArgs e)
        {
            WriteData();
        }
    }

        

    
}
