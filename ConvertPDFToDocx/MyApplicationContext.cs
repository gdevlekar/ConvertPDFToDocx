using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DSOFile;

namespace WindowsFormsApplication2
{
    class MyApplicationContext : System.Windows.Forms.ApplicationContext
    {

        [STAThread]
        static void Main(string[] args)
        {
            // Create the MyApplicationContext, that derives from ApplicationContext,
            // that manages when the application should exit.
            MyApplicationContext context = new MyApplicationContext();

            // Run the application with the specific context. It will exit when



            // the task completes and calls Exit().
            Application.Run(context);
        }

        Task backgroundTask;

        // This is the constructor of the ApplicationContext, we do not want to 
        // block here.
        private MyApplicationContext()
        {
            //// ConvertPDFToDocx
            //backgroundTask = Task.Factory.StartNew(ConvertPDFToDocx);
            //backgroundTask.ContinueWith(TaskComplete);

            //// ConvertDocxToDoc
            backgroundTask = Task.Factory.StartNew(ConvertDocxToDoc);
            backgroundTask.ContinueWith(TaskComplete);

            //// removeMetaData
            //backgroundTask = Task.Factory.StartNew(removeMetaData);
            //backgroundTask.ContinueWith(TaskComplete);


        }

        // This will allow the Application.Run(context) in the main function to 
        // unblock.
        private void TaskComplete(Task src)
        {
            this.ExitThread();
        }

         
        //Perform your actual work here.
        private void ConvertPDFToDocx()
        {
            //Stuff

            string[] files = Directory.GetFiles(@"E:\software projects\personal expriments\RJT01398\RJT01398\", "*.pdf");
            int index = 1;
            foreach (string item in files)
            {

                //if (index>00)

                if (!File.Exists(item.Replace(".pdf", ".docx")))
                {
                    //   Process.Start(@"E:\software projects\personal expriments\RJT01398\RJT01398\E-Evoice8648BHTY185485256 - 0001.pdf");
                    Process proc = Process.Start(item);

                    Thread.Sleep(1000);

                    SendKeys.SendWait("%{f}");

                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);

                    SendKeys.SendWait("{DOWN}");

                    Thread.Sleep(500);
                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);
                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);
                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);

                    SendKeys.SendWait("{ENTER}");                                   //on export button
                    Thread.Sleep(500);


                    SendKeys.SendWait("{ENTER}");           //ms word
                    Thread.Sleep(500);

                    SendKeys.SendWait("{ENTER}");           //word doc
                    Thread.Sleep(500);

                    SendKeys.SendWait("{ENTER}");           //save
                    Thread.Sleep(500);

                    Thread.Sleep(7000);
                    SendKeys.SendWait("%{F4}");             //close
                    Thread.Sleep(100);


                }
                index++;
            }



        }

        private void ConvertDocxToDoc()
        {
            //Stuff
            string filesDestination = @"E:\software projects\personal expriments\Work done\block 4\";
            string[] files = Directory.GetFiles(filesDestination, "*.docx").Where(file => file.EndsWith(".docx", StringComparison.CurrentCultureIgnoreCase))
             .ToArray();
            int index = 1;
            foreach (string item in files)
            {

                //if (index>00)

                if (!File.Exists(item.Replace(".docx", ".doc")))
                {
                    //   Process.Start(@"E:\software projects\personal expriments\RJT01398\RJT01398\E-Evoice8648BHTY185485256 - 0001.pdf");
                    Process proc = Process.Start(item);

                    Thread.Sleep(1000);

                    SendKeys.SendWait("%{f}");
                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);

                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);

                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);

                    SendKeys.SendWait("{RIGHT}");
                    Thread.Sleep(500);

                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);

                    SendKeys.SendWait("{DOWN}");
                    Thread.Sleep(500);

                    SendKeys.SendWait("{ENTER}");           //word 97-2003 
                    Thread.Sleep(1500);

                    SendKeys.SendWait("{ENTER}");                                   //save as button
                    Thread.Sleep(800);

                    SendKeys.SendWait("%{F4}");             //close
                    Thread.Sleep(100);

                }
                index++;
            }



        }


        private void removeMetaData()
        {

            string filesDestination = @"E:\software projects\personal expriments\RJT01398\step 2\";
            string[] files = Directory.GetFiles(filesDestination, "*.doc").Where(file => file.EndsWith(".doc", StringComparison.CurrentCultureIgnoreCase))
             .ToArray();
            foreach (string item in files)
            {
                FileInfo fi = new FileInfo(item);


                //creates new class of oledocumentproperties
                var doc = new OleDocumentPropertiesClass();

                //open your selWected file
                doc.Open(fi.FullName, false, dsoFileOpenOptions.dsoOptionDefault);

                //you can set properties with summaryproperties.nameOfProperty = value; for example
                doc.SummaryProperties.Company = "";
                doc.SummaryProperties.Author = "";
                doc.SummaryProperties.Title = "";
                doc.SummaryProperties.Keywords = "";

                //after making changes, you need to use this line to save them
                doc.Save();

                //-------------------------------------


                //var file = ShellFile.FromFilePath(fi.FullName);

                //string[] oldAuthors = file.Properties.System.Author.Value;
                //string oldTitle = file.Properties.System.Title.Value;

                //ShellProperties shellproperties = file.Properties;


                //ShellPropertyWriter propertyWriter = file.Properties.GetPropertyWriter();
                //propertyWriter.WriteProperty(SystemProperties.System.Author, new string[] { "Author" });
                ////propertyWriter.WriteProperty(SystemProperties.System., new string[] { "Author" });
                //propertyWriter.Close();

            }

        }
    }
}
