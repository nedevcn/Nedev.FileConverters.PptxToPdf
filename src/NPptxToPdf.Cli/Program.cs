using NPptxToPdf;

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
        bool useParallel = false;

        // Parse additional options
        for (int i = 2; i < args.Length; i++)
        {
            if (args[i] == "--parallel" || args[i] == "-p")
            {
                useParallel = true;
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
            
            var converter = new PptxToPdfConverter();
            converter.Convert(pptxFilePath, pdfFilePath, useParallel);
            
            Console.WriteLine("Conversion completed successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Environment.Exit(1);
        }
    }

    static void ShowHelp()
    {
        Console.WriteLine("NPptxToPdf.Cli - Converts PPTX files to PDF");
        Console.WriteLine("Usage:");
        Console.WriteLine("  NPptxToPdf.Cli <input.pptx> <output.pdf> [options]");
        Console.WriteLine("Options:");
        Console.WriteLine("  --parallel, -p    Use parallel processing for faster conversion");
        Console.WriteLine("  --help, -h        Show this help message");
    }
}
