using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestScaningSystem.DataAccess;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;


using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace TestScaningSystem.BusinessLayer
{
    public class DataHandeler
    {
        ExcelAccess ea;
        DatabaseAccess da;
        public DataHandeler()
        {
            da = new DatabaseAccess();
        }
        public DataHandeler(string filePath)
        {
            ea = new ExcelAccess(filePath);
            da = new DatabaseAccess();
        }
        public bool Login(string username, string password)
        {
            return da.Login(username, password);
        }
        public string[] GetSheetNames()
        {
            return ea.GetSheetNames();
        }
        public List<string> GetVenueNames(int workSheet)
        {
            return ea.GetVenueNames(workSheet);
        }
        public List<Student> GetStudentsByClass(string venue, string subject)
        {
            List<string> studentArr = ea.GetStudentsByClass(venue);
            List<Student> students = new List<Student>();
            int i = 0;
            foreach (string item in studentArr)            
            {
                string[] studentInfo = studentArr[i].Split(';');
                string[] info = new string[]
                {
                    studentInfo[0],
                    studentInfo[1],
                    subject
                };
                long id = da.GetStudentID(info);
                students.Add(new Student(id, studentInfo[0], studentInfo[1], studentInfo[2], venue, subject));
                i++;
            }
            
            return students;
        }

        
    }
}
