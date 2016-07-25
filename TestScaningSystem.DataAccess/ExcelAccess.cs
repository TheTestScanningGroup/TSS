using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;


namespace TestScaningSystem.DataAccess
{
    public class ExcelAccess
    {
        //Creates a Excel Workbook object
        public string FileName;
        public Excel.Application xlApp;
        public Excel.Workbook xlWorkBook;
        public Excel.Worksheet xlWorkSheet;
        public Excel.Range xlRange;
        public int worksheet = 0;
        public ExcelAccess(string fileName)
        {
            FileName = fileName;
            xlApp = new Excel.Application();
        }
        
        public void OpenExcelConnection()
        {
            xlWorkBook = xlApp.Workbooks.Open(FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        }

        public void CloseExcelConnection()
        {
            xlApp.Workbooks.Close();
        }
        #region GetSheetNames
        public string[] GetSheetNames()
        {
            OpenExcelConnection();
            int amountOfSheets = xlWorkBook.Worksheets.Count;
            string[] sheetNames = new string[amountOfSheets];
            int i = 0;
            //Iterates through the excel workbook and adds the sheet names to the array
            foreach (Excel.Worksheet item in xlWorkBook.Worksheets)
            {
                sheetNames[i] = item.Name;
                i++;
            }
            CloseExcelConnection();
            return sheetNames;
        }
        #endregion

        #region GetVenueNames
        public List<string> GetVenueNames(int workSheet)
        {
            OpenExcelConnection();
            //Sets which worksheet has to be worked on
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[workSheet + 1];
            worksheet = workSheet + 1;
            xlRange = xlWorkSheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string> venueNames = new List<string>();

            bool venueExists = false;
            //Iterates through the worksheet and gets the venue names
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 4; j <= 4; j++)
                {
                    string temp = (string)(xlRange.Cells[i, j] as Excel.Range).Value2;
                    
                    foreach (string item in venueNames)
                    {
                        //Checks if the venue has already been added to the list
                        if (item == temp)
                        {
                            venueExists = true;
                            break;
                        }
                    }
                    if (venueExists == false)
                    {
                        venueNames.Add(temp);
                    }
                    venueExists = false;
                }
            }
            //Returns the list of venue names
            CloseExcelConnection();
            return venueNames;            
        }
        #endregion

        #region GetStudentByClass
        public List<string> GetStudentsByClass(string venue)
        {
            OpenExcelConnection();
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[worksheet];
            xlRange = xlWorkSheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string> students = new List<string>();
            int counter = 0;
            //Iterates through the cells on the worksheet to get student information
            for (int i = 2; i <= rowCount; i++)
            {
                string surname = "";
                string firstName = "";
                string classID = "";
                string tempVenue = "";
                for (int j = 1; j <= colCount; j++)
                {
                    if (j == 1)
                    {
                        surname = (string)(xlRange.Cells[i, j] as Excel.Range).Value2;
                    }
                    else if (j == 2)
                    {
                        firstName = (string)(xlRange.Cells[i, j] as Excel.Range).Value2;
                    }
                    else if (j == 3)
                    {
                        classID = (string)(xlRange.Cells[i, j] as Excel.Range).Value2;
                    }
                    else if (j == 4)
                    {
                        tempVenue = (string)(xlRange.Cells[i, j] as Excel.Range).Value2;
                    }
                }
                if (tempVenue == venue)
                {
                    students.Add(surname + ";" + firstName + ";" + classID);
                    counter++;
                }
            }
            //Returns the list of students
            CloseExcelConnection();
            return students;
        } 
        #endregion

    }
}
