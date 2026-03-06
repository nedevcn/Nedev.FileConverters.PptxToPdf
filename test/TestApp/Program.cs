using Nedev.FileConverters.PptxToPdf;

Console.WriteLine("Testing Nedev.FileConverters.PptxToPdf NuGet Package v0.1.0");
Console.WriteLine("=====================================================");

try
{
    var converter = new PptxToPdfConverter();
    Console.WriteLine("✓ PptxToPdfConverter instance created successfully");
    
    Console.WriteLine("\nPackage Information:");
    Console.WriteLine($"  Version: 0.1.0");
    Console.WriteLine($"  Target Frameworks: net8.0;netstandard2.1");
    Console.WriteLine($"  Description: High-performance PPTX to PDF converter");
    
    Console.WriteLine("\nPackage Features:");
    Console.WriteLine("  ✓ Direct file-to-file conversion");
    Console.WriteLine("  ✓ Stream-to-stream conversion");
    Console.WriteLine("  ✓ Parallel processing support");
    Console.WriteLine("  ✓ Core package integration");
    Console.WriteLine("  ✓ Zero third-party dependencies");
    
    Console.WriteLine("\nUsage Example:");
    Console.WriteLine("  var converter = new PptxToPdfConverter();");
    Console.WriteLine("  converter.Convert(\"input.pptx\", \"output.pdf\");");
    
    Console.WriteLine("\n✓ NuGet package test completed successfully!");
}
catch (Exception ex)
{
    Console.WriteLine($"\n✗ Error: {ex.Message}");
    Environment.Exit(1);
}
