using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using ThoughtWorks.QRCode;
namespace TestScaningSystem.BusinessLayer
{
    public class QRConverter
    {
        private Image qrCode;
        private string[] alldata;
        QRMedium medium = new QRMedium();
        private string AllDataString;
        List<Image> listOfQRCodes;
        public Image QrCode
        {
            get
            { return qrCode; }
            set
            { qrCode = value; }
        }

        public string[] Alldata
        {
            get{ return alldata; }
            set{ alldata = value; }
        }


        public QRConverter()
        {

        }

        #region GenerateQRCode
        public List<Image> GenerateQRCode(List<Student> listOfStudents, string testDate)
        {
            //Creates a new list of images to store the qr codes
            listOfQRCodes = new List<Image>();
            try
            {
                foreach (Student student in listOfStudents)
                {
                    //Puts the information needed in a specific format to be stored in
                    AllDataString = string.Format("{0};{1};{2};{3};{4};{5}", student.IDNumber, student.Surname, student.FirstName, student.ClassID, student.Subject, testDate);
                    //Generates the QR Code
                    QrCode = medium.Encode(Error_Correction.L, Encode_Mode.BYTE, 1, 7, AllDataString);
                    //Stores the QR Code in the image list
                    listOfQRCodes.Add(QrCode);
                }
                //Returns the list of QR Codes
                return listOfQRCodes;
            }
            catch (Exception)
            {
                return null;
            }
        } 
        #endregion
        public QRConverter(Image qrCode)
        {
            try
            {
                string data = medium.Decode(qrCode);
                Alldata = data.Split(';');
            }
            catch (Exception e)
            {

                throw new SystemException(e.Message);
            }
            
        }
        
        

    }
}
