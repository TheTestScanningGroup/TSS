using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using ThoughtWorks.QRCode.Codec;
using ThoughtWorks.QRCode.Codec.Data;


namespace ThoughtWorks.QRCode
{
    public enum Encode_Mode { ALPHA_NUMERIC, NUMERIC, BYTE };
    public enum Error_Correction { L, M, Q, H };
    public class QRMedium
    {
        QRCodeEncoder qrCodeE = new QRCodeEncoder();
        QRCodeDecoder qrCodeD = new QRCodeDecoder();
        /// <summary>
        /// Encoding Constructor
        /// </summary>
        public QRMedium()
        {

        }

        public Image Encode(Error_Correction Error_Correction, Encode_Mode Encode_Mode, int scale, int version, string Data)
        {
            try
            {
                qrCodeE.QRCodeEncodeMode = (QRCodeEncoder.ENCODE_MODE)Encode_Mode;
                qrCodeE.QRCodeScale = scale;
                qrCodeE.QRCodeVersion = version;
                qrCodeE.QRCodeErrorCorrect = (QRCodeEncoder.ERROR_CORRECTION)Error_Correction;
                return qrCodeE.Encode(Data);
            }
            catch (Exception e)
            {
                throw new SystemException("Error encoding image", e);
            }
            
        }
        public string Decode(Image QRCodeImage)
        {
            try
            {
                return qrCodeD.decode(new QRCodeBitmapImage(new Bitmap(QRCodeImage)));
            }
            catch (Exception e)
            {
                throw new SystemException("Error decoding image", e);
            }
        }
    }
}
