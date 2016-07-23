using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Drawing.Printing;
using System.Threading;

namespace TestScaningSystem.BusinessLayer
{
    public class DocumentHandler
    {
        #region TempleteTypeEnum
        public enum TempleteType
        {
            Lined, Grid, TrueFalse, Monkey, MatchAB, Crossword,
        } 
        #endregion

        #region SaveQRCodesToFile
        public string[] SaveQRCodesToFile(List<Student> listOfStudents, List<Image> listOfQRCodes)
        {
            int listCounter = 0;
            int arrCounter = 0;
            string[] arrPaths = new string[listOfQRCodes.Count];
            foreach (Student student in listOfStudents)
            {
                string path = string.Format(@"C:\Tests\{0}{1}{2}{3}.jpeg", student.Surname, student.FirstName, student.ClassID, student.Venue);
                listOfQRCodes[listCounter].Save(path, System.Drawing.Imaging.ImageFormat.Jpeg);
                arrPaths[arrCounter] = path;
                arrCounter++;
                listCounter++;
            }
            return arrPaths;
        } 
        #endregion

        #region GenerateDocument
        public void GenerateDocument(TempleteType TT, Student student, string qrCodePath, string amountOfCopies)
        {
            MessageFilter.Register();
            Object oMissing = System.Reflection.Missing.Value;
            Object oTrue = true;
            Object oFalse = false;
            //Creates a blank word document
            Word.Application oWord = new Word.Application();
            Word.Document oWordDoc = new Word.Document();
            oWord.Visible = false;
            //Sets what document is going to be created
            object oTemplate = null;
            switch (TT)
            {
                case TempleteType.Lined:
                    oTemplate = @"C:\TestScannigSystem\Lined Answer Sheet.dotx";
                    break;
                case TempleteType.Grid:
                    oTemplate = @"C:\TestScannigSystem\Grid Answer Sheet.dotx";
                    break;
                case TempleteType.TrueFalse:
                    oTemplate = @"C:\TestScannigSystem\True or False Answer Sheet.dotx";
                    break;
                case TempleteType.Monkey:
                    oTemplate = @"C:\TestScannigSystem\Monkey puzzle Answer Sheet.dotx";
                    break;
                case TempleteType.MatchAB:
                    oTemplate = @"C:\TestScannigSystem\Match A to B Answer Sheet.dotx";
                    break;
                case TempleteType.Crossword:
                    oTemplate = @"C:\TestScannigSystem\Crossword Answer Sheet.dotx";
                    break;
            }
            //Creates a document based on the templete selected
            oWordDoc = oWord.Documents.Add(ref oTemplate);

            //Sets where the QR Code is going to sit
            object start = 20;
            object end = 23;
            Word.Range rng = oWordDoc.Range(ref start, ref end);
            rng.InlineShapes.AddPicture(qrCodePath);
            
            start = 195;
            end = 198;
            rng = oWordDoc.Range(ref start, ref end);
            //Places the QR Code
            rng.InlineShapes.AddPicture(qrCodePath);
            
            
            //Settings for how the document needs to be printed
            
            object copies = amountOfCopies;
            object pages = "";
            object range = Word.WdPrintOutRange.wdPrintAllDocument;
            object items = Word.WdPrintOutItem.wdPrintDocumentContent;
            object pageType = Word.WdPrintOutPages.wdPrintAllPages;

            //Prints the document
            oWordDoc.PrintOut(oTrue, oFalse, range, oMissing, oMissing, oMissing, items, copies, oMissing, pageType, oFalse, oTrue, oMissing, oFalse, oMissing, oMissing, oMissing, oMissing);
            MessageFilter.Revoke();
        } 
        #endregion
    }    
}
