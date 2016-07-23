using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScaningSystem.BusinessLayer
{
    public class Student
    {
        private long idNumber;

        public long IDNumber
        {
            get { return idNumber; }
            set { idNumber = value; }
        }

        private string surname;

        public string Surname
        {
            get { return surname; }
            set { surname = value; }
        }
        private string firstName;

        public string FirstName
        {
            get { return firstName; }
            set { firstName = value; }
        }
        private string classID;

        public string ClassID
        {
            get { return classID; }
            set { classID = value; }
        }
        private string venue;

        public string Venue
        {
            get { return venue; }
            set { venue = value; }
        }
        private string subject;

        public string Subject
        {
            get { return subject; }
            set { subject = value; }
        }


        public Student(long id, string surname, string firstName, string classID, string venue, string subject)
        {
            IDNumber = id;
            Surname = surname;
            FirstName = firstName;
            ClassID = classID;
            Venue = venue;
            Subject = subject;
        }
    }
}
