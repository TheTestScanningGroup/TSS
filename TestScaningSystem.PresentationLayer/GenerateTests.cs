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
using TestScaningSystem.BusinessLayer;

namespace TestScaningSystem.PresentationLayer
{
    public partial class GenerateTests : Form
    {
        public GenerateTests()
        {
            InitializeComponent();
        }
        DataHandeler dh;
        DocumentHandler doch = new DocumentHandler();
        List<Student> students = new List<Student>();
        private void showBalloon(string title, string body)
        {
            NotifyIcon notifyIcon = new NotifyIcon();
            notifyIcon.Visible = true;

            if (title != null)
            {
                notifyIcon.BalloonTipTitle = title;
            }

            if (body != null)
            {
                notifyIcon.BalloonTipText = body;
            }
            notifyIcon.Icon = SystemIcons.Application;
            notifyIcon.ShowBalloonTip(30000);
            notifyIcon.Dispose();

        }
        private void btnbrowse_Click(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            openFileDialog1.FileName = null;
            openFileDialog1.Title = "Please select class list";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {

                string name = openFileDialog1.FileName;
                if (name.EndsWith(".xls") || name.EndsWith(".xlsx") || name.EndsWith(".xlsm"))
                {
                    showBalloon("Populating", "Please wait while we populate your subjects ");
                    edtlocation.Text = openFileDialog1.FileName;
                    //Creates a Data Handler object
                    dh = new DataHandeler(edtlocation.Text);
                    //Creates an array an inputs the information returned by GetSheetNames()
                    string[] sheetNames = dh.GetSheetNames();
                    //Adds the sheet names to the combox
                    for (int i = 0; i < sheetNames.Length; i++)
                    {
                        comboBox2.Items.Add(sheetNames[i]);
                    }
                    comboBox2.Enabled = true;

                }
                else
                {
                    DialogResult result2 = MessageBox.Show("Please select a valid excel file!", "Invalid file type", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);
                    if (result2 == DialogResult.Retry)
                    {
                        btnbrowse_Click(sender, e);
                    }
                }
                

            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            showBalloon("Populating", "Please wait while we populate the list of venues");
            List<string> venueNames = new List<string>();
            comboBox3.Items.Clear();
            //Creates a list that is filled by the venues returned by GetVenueNames()
            venueNames=dh.GetVenueNames(comboBox2.SelectedIndex);
            //Adds the venues stored in the list to the combobox
            foreach (string item in venueNames)
            {
                if (item != null)
                {
                    comboBox3.Items.Add(item);
                }
                
            }
            comboBox3.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool lineChecked = false;
            bool gridChecked = false;
            bool trueFalseChecked = false;
            bool monkeyChecked = false;
            bool matchABChecked = false;
            bool crosswordChecked = false;
            
            if (checkBox1.Checked == true)
            {
                lineChecked = true;
            }
            if (checkBox2.Checked == true)
            {
                trueFalseChecked = true;
            }
            if (checkBox3.Checked == true)
            {
                monkeyChecked = true;
            }
            if (checkBox4.Checked == true)
            {
                matchABChecked = true;
            }
            if (checkBox5.Checked == true)
            {
                gridChecked = true;
            }
            if (checkBox6.Checked == true)
            {
                crosswordChecked = true;
            }
            if (lineChecked == false && trueFalseChecked == false && monkeyChecked == false && matchABChecked == false && gridChecked == false && crosswordChecked == false)
            {
                MessageBox.Show("Please select a option.");
            }
            else
            {
                //Fills the List<Student> with the students returned by GetStudentByClass()
                students = dh.GetStudentsByClass(comboBox3.SelectedItem.ToString(), comboBox2.SelectedItem.ToString());
                //Creates a empty list of QR Codes
                List<Image> listOfQRCodes = new List<Image>();
                try
                {
                    QRConverter qrcon = new QRConverter();
                    //Fills the empty list of QR Codes with codes generated by GenerateQRCode()
                    listOfQRCodes = qrcon.GenerateQRCode(students, dateTimePicker1.Text);
                    //Creates and stores QR Codes to a file and stores their file paths into an array
                    string[] qrCodePaths = doch.SaveQRCodesToFile(students, listOfQRCodes);
                    int counter = 0;
                    //Iterates through the list of students and generates a document
                    foreach (Student student in students)
                    {
                        if (lineChecked == true)
                        {
                            doch.GenerateDocument(DocumentHandler.TempleteType.Lined, student, qrCodePaths[counter], numericUpDown1.Value.ToString(),dateTimePicker1.Value);
                        }
                        if (trueFalseChecked == true)
                        {
                            //doch.GenerateDocument(DocumentHandler.TempleteType.Crossword, student, qrCodePaths[counter], numericUpDown1.Value.ToString(),dateTimePicker1.Value);
                        }
                        if (monkeyChecked == true)
                        {
                            //doch.GenerateDocument(DocumentHandler.TempleteType.Grid, student, qrCodePaths[counter], numericUpDown1.Value.ToString(),dateTimePicker1.Value);
                        }
                        if (matchABChecked == true)
                        {
                            //doch.GenerateDocument(DocumentHandler.TempleteType.MatchAB, student, qrCodePaths[counter], numericUpDown1.Value.ToString(),dateTimePicker1.Value);
                        }
                        if (gridChecked == true)
                        {
                            //doch.GenerateDocument(DocumentHandler.TempleteType.Monkey, student, qrCodePaths[counter], numericUpDown1.Value.ToString(),dateTimePicker1.Value);
                        }
                        if (crosswordChecked == true)
                        {
                            //doch.GenerateDocument(DocumentHandler.TempleteType.TrueFalse, student, qrCodePaths[counter], numericUpDown1.Value.ToString(),dateTimePicker1.Value);
                        }
                        counter++;
                    }
                }
                catch (Exception er)
                {

                    MessageBox.Show(er.Message);
                }
            }             
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            groupBox1.Enabled = true;
            button1.Enabled = true;
            dateTimePicker1.Enabled = true;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                numericUpDown1.Enabled = true;
            }
            else
            {
                numericUpDown1.Enabled = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                numericUpDown2.Enabled = true;
            }
            else
            {
                numericUpDown2.Enabled = false;
            }
        }
    }
}
