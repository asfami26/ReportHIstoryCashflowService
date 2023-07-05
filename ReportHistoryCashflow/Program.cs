using Dapper;
using Microsoft.IdentityModel.Protocols;
using System;
using System.Data;
using System.ServiceProcess;
using Microsoft.Extensions.Configuration;
using ClosedXML.Excel;



namespace ReportHistoryCashflow
{
    class Program
    {
        static void Main(String[] args)
        {

            using (var service = new FileWriteService())
            {
                //ServiceBase.Run(service);
               
                service.OnDebug();
            }
        }
    }
}

