using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ReportHistoryCashflow.Data;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ReportHistoryCashflow
{
    class Program
    {
        static void Main(string[] args)
        {
        #if DEBUG
                    // Jalankan layanan dalam mode debug
                    var service = new FileWriteService();
                    service.StartAsync(CancellationToken.None).Wait();
        #else
                    // Jalankan GenericHost dalam mode production
                    var host = CreateHostBuilder(args).Build();
                    host.Run();
        #endif

        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    // Daftarkan hosted service (FileWriteService) di sini
                    services.AddHostedService<FileWriteService>();
                });
    }
}
