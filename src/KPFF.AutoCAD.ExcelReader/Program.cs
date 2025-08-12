using KPFF.AutoCAD.ExcelReader.Services;

namespace KPFF.AutoCAD.ExcelReader;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("KPFF AutoCAD Excel Reader - Starting...");
        
        var processor = new ExcelProcessor();
        var server = new NamedPipeServer(processor);

        // Handle graceful shutdown
        Console.CancelKeyPress += (sender, e) => {
            e.Cancel = true;
            server.Stop();
        };

        try
        {
            await server.StartAsync();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
            Environment.Exit(1);
        }
        
        Console.WriteLine("Excel Reader stopped.");
    }
}