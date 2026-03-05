using Nedev.PptxToPdf.Pdf;
using Nedev.PptxToPdf.Pptx;

namespace Nedev.PptxToPdf;

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
            OnProgressChanged(0, "Starting conversion...");

            using var pptx = PptxDocument.Open(pptxFilePath);
            using var pdf = new PdfDocument(pdfFilePath);

            pdf.Initialize();

            int slideCount = pptx.Slides.Count;

            if (useParallelProcessing && slideCount > 1)
            {
                // Use parallel processing for multiple slides
                var renderer = new PdfRenderer(pdf);
                var slideTasks = new List<Task>();
                var progressLock = new object();
                int processedSlides = 0;

                for (int i = 0; i < slideCount; i++)
                {
                    var slideIndex = i;
                    slideTasks.Add(Task.Run(() =>
                    {
                        var slide = pptx.Slides[slideIndex];
                        if (pptx.Presentation == null) return;

                        try
                        {
                            var width = pptx.Presentation.SlideWidth / 914400.0 * 72;
                            var height = pptx.Presentation.SlideHeight / 914400.0 * 72;

                            lock (pdf) // Ensure thread-safe access to PDF document
                            {
                                var page = pdf.AddPage(width, height);
                                renderer.RenderSlide(page, slide, pptx);
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
                // Use sequential processing
                var renderer = new PdfRenderer(pdf);

                for (int i = 0; i < slideCount; i++)
                {
                    var slide = pptx.Slides[i];
                    if (pptx.Presentation == null) continue;

                    try
                    {
                        OnProgressChanged((i + 1) * 100 / slideCount, $"Processing slide {i + 1}/{slideCount}");

                        var width = pptx.Presentation.SlideWidth / 914400.0 * 72;
                        var height = pptx.Presentation.SlideHeight / 914400.0 * 72;

                        var page = pdf.AddPage(width, height);
                        renderer.RenderSlide(page, slide, pptx);
                    }
                    catch (Exception ex)
                    {
                        OnErrorOccurred(ex, $"Error processing slide {i + 1}");
                        // Continue processing other slides
                    }
                }
            }

            pdf.Save();
            OnProgressChanged(100, "Conversion completed successfully");
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
            OnProgressChanged(0, "Starting conversion...");

            using var pptx = PptxDocument.Open(pptxStream);
            using var pdf = new PdfDocument(pdfStream);

            pdf.Initialize();

            int slideCount = pptx.Slides.Count;

            if (useParallelProcessing && slideCount > 1)
            {
                // Use parallel processing for multiple slides
                var renderer = new PdfRenderer(pdf);
                var slideTasks = new List<Task>();
                var progressLock = new object();
                int processedSlides = 0;

                for (int i = 0; i < slideCount; i++)
                {
                    var slideIndex = i;
                    slideTasks.Add(Task.Run(() =>
                    {
                        var slide = pptx.Slides[slideIndex];
                        if (pptx.Presentation == null) return;

                        try
                        {
                            var width = pptx.Presentation.SlideWidth / 914400.0 * 72;
                            var height = pptx.Presentation.SlideHeight / 914400.0 * 72;

                            lock (pdf) // Ensure thread-safe access to PDF document
                            {
                                var page = pdf.AddPage(width, height);
                                renderer.RenderSlide(page, slide, pptx);
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
                // Use sequential processing
                var renderer = new PdfRenderer(pdf);

                for (int i = 0; i < slideCount; i++)
                {
                    var slide = pptx.Slides[i];
                    if (pptx.Presentation == null) continue;

                    try
                    {
                        OnProgressChanged((i + 1) * 100 / slideCount, $"Processing slide {i + 1}/{slideCount}");

                        var width = pptx.Presentation.SlideWidth / 914400.0 * 72;
                        var height = pptx.Presentation.SlideHeight / 914400.0 * 72;

                        var page = pdf.AddPage(width, height);
                        renderer.RenderSlide(page, slide, pptx);
                    }
                    catch (Exception ex)
                    {
                        OnErrorOccurred(ex, $"Error processing slide {i + 1}");
                        // Continue processing other slides
                    }
                }
            }

            pdf.Save();
            OnProgressChanged(100, "Conversion completed successfully");
        }
        catch (Exception ex)
        {
            OnErrorOccurred(ex, "Error during conversion");
            throw;
        }
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
