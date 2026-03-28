using Nedev.FileConverters.PptxToPdf.Pdf;
using Nedev.FileConverters.PptxToPdf.Pptx;

namespace Nedev.FileConverters.PptxToPdf;

public class PptxToPdfConverter
{
    public event EventHandler<ConversionProgressEventArgs>? ProgressChanged;
    public event EventHandler<ConversionErrorEventArgs>? ErrorOccurred;

    public void Convert(string pptxFilePath, string pdfFilePath, bool useParallelProcessing = false)
    {
        if (string.IsNullOrEmpty(pptxFilePath))
            throw new ArgumentNullException(nameof(pptxFilePath), "PPTX file path cannot be null or empty");

        if (string.IsNullOrEmpty(pdfFilePath))
            throw new ArgumentNullException(nameof(pdfFilePath), "PDF file path cannot be null or empty");

        if (!File.Exists(pptxFilePath))
            throw new FileNotFoundException($"PPTX file not found: {pptxFilePath}");

        try
        {
            using var pptx = PptxDocument.Open(pptxFilePath);
            using var pdf = new PdfDocument(pdfFilePath);
            ConvertInternal(pptx, pdf, useParallelProcessing);
        }
        catch (Exception ex)
        {
            OnErrorOccurred(ex, "Error during conversion");
            throw;
        }
    }

    public void Convert(Stream pptxStream, Stream pdfStream, bool useParallelProcessing = false)
    {
        if (pptxStream == null)
            throw new ArgumentNullException(nameof(pptxStream), "PPTX stream cannot be null");

        if (pdfStream == null)
            throw new ArgumentNullException(nameof(pdfStream), "PDF stream cannot be null");

        try
        {
            using var pptx = PptxDocument.Open(pptxStream);
            using var pdf = new PdfDocument(pdfStream);
            ConvertInternal(pptx, pdf, useParallelProcessing);
        }
        catch (Exception ex)
        {
            OnErrorOccurred(ex, "Error during conversion");
            throw;
        }
    }

    private void ConvertInternal(PptxDocument pptx, PdfDocument pdf, bool useParallelProcessing)
    {
        OnProgressChanged(0, "Starting conversion...");

        pdf.Initialize();

        var presentation = pptx.Presentation
            ?? throw new InvalidOperationException("Presentation metadata could not be loaded from the PPTX file.");

        int slideCount = pptx.Slides.Count;
        if (slideCount == 0)
        {
            pdf.Save();
            OnProgressChanged(100, "Conversion completed successfully (no slides found)");
            return;
        }

        var slideWidth = presentation.SlideWidth / 914400.0 * 72;
        var slideHeight = presentation.SlideHeight / 914400.0 * 72;
        var renderer = new PdfRenderer(pdf);

        if (useParallelProcessing && slideCount > 1)
        {
            var pages = new PdfPage[slideCount];
            for (int i = 0; i < slideCount; i++)
            {
                pages[i] = pdf.AddPage(slideWidth, slideHeight);
            }

            var progressLock = new object();
            var slideTasks = new List<Task>(slideCount);
            int processedSlides = 0;

            for (int i = 0; i < slideCount; i++)
            {
                var slideIndex = i;
                slideTasks.Add(Task.Run(() =>
                {
                    try
                    {
                        lock (pdf)
                        {
                            renderer.RenderSlide(pages[slideIndex], pptx.Slides[slideIndex], pptx);
                        }

                        lock (progressLock)
                        {
                            processedSlides++;
                            OnProgressChanged(processedSlides * 100 / slideCount, $"Processing slide {processedSlides}/{slideCount}");
                        }
                    }
                    catch (Exception ex)
                    {
                        OnErrorOccurred(ex, $"Error processing slide {slideIndex + 1}");
                    }
                }));
            }

            Task.WaitAll(slideTasks.ToArray());
        }
        else
        {
            for (int i = 0; i < slideCount; i++)
            {
                try
                {
                    var page = pdf.AddPage(slideWidth, slideHeight);
                    renderer.RenderSlide(page, pptx.Slides[i], pptx);
                    OnProgressChanged((i + 1) * 100 / slideCount, $"Processing slide {i + 1}/{slideCount}");
                }
                catch (Exception ex)
                {
                    OnErrorOccurred(ex, $"Error processing slide {i + 1}");
                }
            }
        }

        pdf.Save();
        OnProgressChanged(100, "Conversion completed successfully");
    }

    protected virtual void OnProgressChanged(int progress, string message)
    {
        ProgressChanged?.Invoke(this, new ConversionProgressEventArgs(progress, message));
    }

    protected virtual void OnErrorOccurred(Exception exception, string message)
    {
        ErrorOccurred?.Invoke(this, new ConversionErrorEventArgs(exception, message));
    }
}

public class ConversionProgressEventArgs : EventArgs
{
    public int Progress { get; }
    public string Message { get; }

    public ConversionProgressEventArgs(int progress, string message)
    {
        Progress = progress;
        Message = message;
    }
}

public class ConversionErrorEventArgs : EventArgs
{
    public Exception Exception { get; }
    public string Message { get; }

    public ConversionErrorEventArgs(Exception exception, string message)
    {
        Exception = exception;
        Message = message;
    }
}
