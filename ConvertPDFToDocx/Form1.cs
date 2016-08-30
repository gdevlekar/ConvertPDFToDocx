using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {

        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName,
            string lpWindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        Task backgroundTask;
        public Form1()
        {
            InitializeComponent();


             
            //IntPtr calculatorHandle = FindWindow("CalcFrame", @"C:\Program Files(x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe\E-Evoice8648BHTY185485256 - 0001.pdf");

            //Process.Start(@"E:\software projects\personal expriments\RJT01398\RJT01398\E-Evoice8648BHTY185485256 - 0001.pdf");
            //SendKeys.Send("%{f}");
        }

         



    }
}
