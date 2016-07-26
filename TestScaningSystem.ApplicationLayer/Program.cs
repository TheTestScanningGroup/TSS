using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using TestScaningSystem.PresentationLayer;

namespace TestScaningSystem.ApplicationLayer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Thread FirstRun = new Thread(TempleteCreater);
            FirstRun.Start();
            //Application.Run(new Login());
            Application.Run(new GenerateTests());
        }
        static void TempleteCreater()
        {
            string[] Templete = new string[5];

            Templete[0] = "Lined Answer Sheet.dotx";
            //Templete[1] = "Grid Answer Sheet.dotx";
            //Templete[2] = "True or False Answer Sheet.dotx";
            //Templete[3] = "Monkey puzzle Answer Sheet.dotx";
            //Templete[4] = "Match A to B Answer Sheet.dotx";
            //Templete[5] = "Crossword Answer Sheet.dotx";

            string MainDirectory = @"C:\TestScannigSystem\";
            if (!Directory.Exists(MainDirectory))
            {
                Directory.CreateDirectory(MainDirectory);
                foreach (string item in Templete)
                {
                    File.Copy(item, string.Format("{0}{1}", MainDirectory, item));
                }
            }
            else
            {
                string[] files = Directory.GetFiles(MainDirectory);
                foreach (string item in files)
                {
                    File.Delete(item);
                }
                foreach (string item in Templete)
                {
                    File.Copy(item, string.Format("{0}{1}", MainDirectory, item));
                }
            }
        }
    }
}
