using System.Reflection;
using Nedev.FileConverters.Core;

namespace Nedev.FileConverters.PptxToPdf;

public static class PptxToPdfCoreRegistration
{
    private const string SourceFormat = "pptx";
    private const string TargetFormat = "pdf";

    private static readonly object SyncRoot = new();
    private static int _registrationState;
    private static Exception? _registrationException;

    public static void EnsureRegistered()
    {
        if (!TryRegister(throwOnFailure: true))
        {
            throw new InvalidOperationException("PPTX to PDF converter registration unexpectedly failed.");
        }
    }

    internal static void EnsureRegisteredSilently()
    {
        TryRegister(throwOnFailure: false);
    }

    private static bool TryRegister(bool throwOnFailure)
    {
        if (Volatile.Read(ref _registrationState) == 1)
            return true;

        lock (SyncRoot)
        {
            if (_registrationState == 1)
                return true;

            if (_registrationState == -1)
            {
                if (throwOnFailure)
                {
                    throw new InvalidOperationException(
                        "Failed to register Nedev.FileConverters.PptxToPdf with Nedev.FileConverters.Core.",
                        _registrationException);
                }

                return false;
            }

            try
            {
                RegisterWithCoreRegistry();
                _registrationState = 1;
                return true;
            }
            catch (Exception ex)
            {
                _registrationException = ex;
                _registrationState = -1;

                if (throwOnFailure)
                {
                    throw new InvalidOperationException(
                        "Failed to register Nedev.FileConverters.PptxToPdf with Nedev.FileConverters.Core.",
                        ex);
                }

                return false;
            }
        }
    }

    private static void RegisterWithCoreRegistry()
    {
        var registryAssembly = typeof(IFileConverter).Assembly;
        var registryType = registryAssembly.GetType("Nedev.FileConverters.Core.ConverterRegistry")
            ?? throw new InvalidOperationException("Could not locate the internal ConverterRegistry type.");

        var mapField = registryType.GetField("_map", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Could not locate ConverterRegistry._map.");
        var initializedField = registryType.GetField("_initialized", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Could not locate ConverterRegistry._initialized.");
        var initLockField = registryType.GetField("_initLock", BindingFlags.NonPublic | BindingFlags.Static);

        var gate = initLockField?.GetValue(null) ?? SyncRoot;
        lock (gate)
        {
            var map = mapField.GetValue(null)
                ?? throw new InvalidOperationException("ConverterRegistry._map is unexpectedly null.");

            var indexer = map.GetType().GetProperty("Item")
                ?? throw new InvalidOperationException("Could not locate the ConverterRegistry map indexer.");

            indexer.SetValue(map, new CoreRegistryAdapter(), new object[] { (SourceFormat, TargetFormat) });
            initializedField.SetValue(null, true);
        }
    }

    private sealed class CoreRegistryAdapter : IFileConverter
    {
        public Stream Convert(Stream input)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));

            var outputStream = new MemoryStream();
            new PptxToPdfConverter().Convert(input, outputStream);
            outputStream.Position = 0;
            return outputStream;
        }
    }
}
