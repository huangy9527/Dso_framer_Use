using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace dsoFramerTestUse
{
    static class Program
    {
        public static KeyboardHook kh;  //Declare hook obj
        public static MouseHook mh;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (kh = new KeyboardHook()) //create keyboard hook obj, diff from C++
            {
                using (mh = new MouseHook())
                {
                    Application.Run(new Form1());
                }
            }
        }
    }
}
