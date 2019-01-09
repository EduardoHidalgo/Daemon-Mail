using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Net;
using System.Diagnostics;

namespace ExecuteBackground
{
    class Program
    {
        public static void BackgroundProcess()
        {
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo("BackgroundEmail.exe");
            p.StartInfo.WorkingDirectory = @"C:\Users\Cancun\Desktop\BackgroundEmail\BackgroundEmail\bin\Debug";
            p.StartInfo.CreateNoWindow = false;
            //p.StartInfo.CreateNoWindow = true;
            p.StartInfo.UseShellExecute = false;
            p.Start();
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Presiona ENTER para iniciar el proceso en segundo plano");
            Console.ReadKey();
            BackgroundProcess();
        }
    }
}
