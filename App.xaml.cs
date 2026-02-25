using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using LeakTrendViewer.Services;

namespace LeakTrendViewer;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private const string AutoExportFlag = "--auto-export";
    private const string StartArg = "--start";
    private const string EndArg = "--end";
    private const string OutputArg = "--output";
    private const string DataDirArg = "--datadir";
    private const string DateFormat = "yyyy-MM-dd HH:mm";
    private const string DefaultDataDir = @"C:\B_MilliData";

    protected override void OnStartup(StartupEventArgs e)
    {
        if (e.Args.Contains(AutoExportFlag, StringComparer.OrdinalIgnoreCase))
        {
            RunAutoExport(e.Args);
            Shutdown();
            return;
        }

        base.OnStartup(e);
    }

    private static void RunAutoExport(IReadOnlyList<string> args)
    {
        // Run on a plain thread to avoid WPF SynchronizationContext deadlock
        // when calling async methods with .GetAwaiter().GetResult()
        Exception? threadError = null;
        var thread = new Thread(() =>
        {
            try
            {
                RunAutoExportCore(args);
            }
            catch (Exception ex)
            {
                threadError = ex;
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (threadError != null)
        {
            Console.Error.WriteLine($"[AutoExport] Failed: {threadError}");
            MessageBox.Show(
                threadError.Message,
                "LeakTrendViewer Auto Export Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static void RunAutoExportCore(IReadOnlyList<string> args)
    {
        Console.WriteLine("[AutoExport] Starting");

        var startText = GetOptionValue(args, StartArg);
        var endText = GetOptionValue(args, EndArg);
        var outputPath = GetOptionValue(args, OutputArg);
        var dataDir = GetOptionValue(args, DataDirArg) ?? DefaultDataDir;

        if (string.IsNullOrWhiteSpace(startText) || string.IsNullOrWhiteSpace(endText))
        {
            throw new ArgumentException($"Both {StartArg} and {EndArg} are required. Expected format: {DateFormat}");
        }

        if (!DateTime.TryParseExact(startText, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var start))
        {
            throw new ArgumentException($"Invalid {StartArg} format. Expected format: {DateFormat}");
        }

        if (!DateTime.TryParseExact(endText, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var end))
        {
            throw new ArgumentException($"Invalid {EndArg} format. Expected format: {DateFormat}");
        }

        if (end < start)
        {
            throw new ArgumentException("End time must be greater than or equal to start time.");
        }

        if (!Directory.Exists(dataDir))
        {
            throw new DirectoryNotFoundException($"Data directory not found: {dataDir}");
        }

        var finalOutputPath = string.IsNullOrWhiteSpace(outputPath)
            ? Path.Combine(dataDir, $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx")
            : outputPath;

        var outputDirectory = Path.GetDirectoryName(finalOutputPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        // Resolve daily MDB files: yyyyMMdd_000000.mdb pattern
        var fileResolver = new MdbFileResolver();
        var resolved = fileResolver.ResolveFiles(dataDir, start, end);

        Console.WriteLine($"[AutoExport] Date range: {start:yyyy-MM-dd HH:mm} ~ {end:yyyy-MM-dd HH:mm}");
        Console.WriteLine($"[AutoExport] Found {resolved.ExistingFiles.Count} MDB file(s):");
        foreach (var f in resolved.ExistingFiles)
        {
            Console.WriteLine($"  - {Path.GetFileName(f)}");
        }

        if (resolved.MissingFiles.Count > 0)
        {
            Console.WriteLine($"[AutoExport] Missing {resolved.MissingFiles.Count} file(s):");
            foreach (var m in resolved.MissingFiles)
            {
                Console.WriteLine($"  - {m}");
            }
        }

        if (resolved.ExistingFiles.Count == 0)
        {
            throw new FileNotFoundException(
                $"No MDB files found in '{dataDir}' for date range {start:yyyy-MM-dd} ~ {end:yyyy-MM-dd}");
        }

        // Load data from all resolved daily files
        var dataLoader = new MdbDataLoader();
        var loadResult = dataLoader.LoadAsync(resolved.ExistingFiles, start, end).GetAwaiter().GetResult();

        var rowNo = 1;
        foreach (var record in loadResult.Records)
        {
            record.RowNo = rowNo++;
        }

        foreach (var warning in loadResult.Warnings)
        {
            Console.WriteLine($"[AutoExport][Warning] {warning}");
        }

        Console.WriteLine($"[AutoExport] Loaded {loadResult.Records.Count} records from {resolved.ExistingFiles.Count} file(s)");

        if (loadResult.Records.Count == 0)
        {
            throw new InvalidOperationException(
                $"No data found in the specified time range: {start:yyyy-MM-dd HH:mm} ~ {end:yyyy-MM-dd HH:mm}");
        }

        var visibleColumns = new List<string> { "Data Time", "LeakRate2" };
        var exporter = new ExcelExporter();
        exporter.Export(finalOutputPath, loadResult.Records, visibleColumns);

        Console.WriteLine($"[AutoExport] Completed: {finalOutputPath}");

        Process.Start(new ProcessStartInfo
        {
            FileName = finalOutputPath,
            UseShellExecute = true
        });
    }

    private static string? GetOptionValue(IReadOnlyList<string> args, string optionName)
    {
        for (var i = 0; i < args.Count; i++)
        {
            if (!string.Equals(args[i], optionName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (i + 1 >= args.Count)
            {
                return null;
            }

            return args[i + 1];
        }

        return null;
    }
}

