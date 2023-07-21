using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using ReportHistoryCashflow.Data;
using System;
using System.IO;
using System.ServiceProcess;

namespace ReportHistoryCashflow
{
    class Program
    {
        static void Main(string[] args)
        {
           
            using (var service = new FileWriteService())
            {
                
                service.OnDebug();
                if (Environment.UserInteractive)
                {
                    Console.WriteLine("Press enter to stop the service...");
                    Console.ReadLine();
                    Environment.Exit(0);
                }
                else
                {
                    service.Working();
                }
            }
        }

     
    }
}

