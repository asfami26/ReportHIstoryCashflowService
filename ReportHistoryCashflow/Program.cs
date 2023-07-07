
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

