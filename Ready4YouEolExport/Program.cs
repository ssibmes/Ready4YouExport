using System;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace Ready4YouEolExport
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new CsvImport());
            }
            catch (WebException e)
            {
                MessageBox.Show("Something went wrong.\nPlease contact the Admin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    Errorlog(((HttpWebResponse)e.Response).StatusCode + "\n" + ((HttpWebResponse)e.Response).StatusDescription);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong.\nPlease contact the Admin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Errorlog(ex.Message + "\n" + ex.StackTrace);
            }
        }
        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show("Something went wrong.\nPlease contact the Admin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Errorlog((e.Exception as Exception).Message + "\n" + (e.Exception as Exception).StackTrace);
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show("Something went wrong.\nPlease contact the Admin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Errorlog((e.ExceptionObject as Exception).Message + "\n" + (e.ExceptionObject as Exception).StackTrace);
        }

        static void Errorlog(string message)
        {
            using (var writer = System.IO.File.AppendText("ErrorLog.txt"))
            {
                writer.WriteLine(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss:ffffff") + ": " + message);
                writer.Flush();
                writer.Close();
            }
        }

    }
}
