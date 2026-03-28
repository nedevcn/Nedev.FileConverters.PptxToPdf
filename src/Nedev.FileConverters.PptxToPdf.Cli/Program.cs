using Nedev.FileConverters;
using Nedev.FileConverters.PptxToPdf;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            ShowHelp();
            return;
        }

        string pptxFilePath = args[0];
        string pdfFilePath = args[1];
        bool useCoreConverter = false;
        bool useParallel = false;

        // Parse additional options
        for (int i = 2; i < args.Length; i++)
        {
            if (args[i] == "--parallel" || args[i] == "-p")
            {
                useParallel = true;
            }
            else if (args[i] == "--core" || args[i] == "-c")
            {
                useCoreConverter = true;
            }
            else if (args[i] == "--help" || args[i] == "-h")
            {
                ShowHelp();
                return;
            }
        }

        try
        {
            Console.WriteLine($"Converting {pptxFilePath} to {pdfFilePath}...");
            
            if (useCoreConverter)
            {
                // Use the Core package entry point after explicit registration.
                using var inputStream = File.OpenRead(pptxFilePath);
                PptxToPdfCoreRegistration.EnsureRegistered();
                using var convertedStream = Converter.Convert(inputStream, "pptx", "pdf");
                using var outputStream = File.Create(pdfFilePath);
                convertedStream.CopyTo(outputStream);
                Console.WriteLine("Conversion completed successfully using the Core entry point!");
            }
            else
            {
                // Use the direct converter with parallel processing option
                var converter = new PptxToPdfConverter();
                converter.Convert(pptxFilePath, pdfFilePath, useParallel);
                Console.WriteLine("Conversion completed successfully!");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Environment.Exit(1);
        }
    }

    static void ShowHelp()
    {
        Console.WriteLine("Nedev.FileConverters.PptxToPdf.Cli - Converts PPTX files to PDF");
        Console.WriteLine("Usage:");
        Console.WriteLine("  Nedev.FileConverters.PptxToPdf.Cli <input.pptx> <output.pdf> [options]");
        Console.WriteLine("Options:");
        Console.WriteLine("  --parallel, -p    Use parallel processing for faster conversion (direct mode only)");
        Console.WriteLine("  --core, -c        Register with Core and use the Converter entry point");
        Console.WriteLine("  --help, -h        Show this help message");
    }
}
