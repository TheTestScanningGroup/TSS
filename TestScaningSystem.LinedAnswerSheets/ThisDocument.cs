using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;

namespace TestScaningSystem.LinedAnswerSheets
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }
        string name = "";
        string surname = "";
        public void  ThisDocument_AddData(string surname, string firstname)
        {
            lblName.Text = firstname;
            lblSurname.Text = surname;
        }
        
        public void AddQRCodeToDocument(List<Image> listOfQRCodes, List<string> listOfStudents)
        {
            int i = 0;
            foreach (Image item in listOfQRCodes)
            {
                string[] temp = listOfStudents[i].Split(';');
                
                lblSurname.Text = temp[0];
                lblName.Text = temp[1];
                pictureBox1.Image = item;
                object copies = "1";
                object pages = "";
                object range = Word.WdPrintOutRange.wdPrintAllDocument;
                object items = Word.WdPrintOutItem.wdPrintDocumentContent;
                object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                object oTrue = true;
                object oFalse = false;

                this.PrintOut(ref oTrue, ref oFalse, ref range, ref missing, ref missing, ref missing,
                    ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
                    ref missing, ref oFalse, ref missing, ref missing, ref missing, ref missing);
                i++;
            }
        }

        private void AddLabelInfo()
        {
            lblSurname.Text = surname;
            lblName.Text = name;
        }
        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
