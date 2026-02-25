using System.Globalization;
using System.IO;

namespace LeakTrendViewer.Services;

public class  MdbFileResolver
{
    private static readonly string[] NamePatterns =
    [
        "{0:yyyyMMdd}_000000.mdb",
        "{0:yyyyMMdd}.mdb",
        "{0:yyyy-MM-dd}.mdb"
    ];

    public FileResolveResult ResolveFiles(IEnumerable<string> baseDirectories, DateTime start, DateTime end)
    {
        var allCandidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var missing = new List<string>();
        var roots = baseDirectories.Where(Directory.Exists).ToList();

        var day = start.Date;
        while (day <= end.Date)
        {
            var matched = false;
            foreach (var root in roots)
            {
                foreach (var pattern in NamePatterns)
                {
                    var fileName = string.Format(CultureInfo.InvariantCulture, pattern, day);
                    var candidate = Path.Combine(root, fileName);
                    allCandidates.Add(candidate);
                    if (File.Exists(candidate))
                    {
                        existing.Add(candidate);
                        matched = true;
                        break;
                    }

                    IEnumerable<string> discovered = Array.Empty<string>();
                    try
                    {
                        discovered = Directory.EnumerateFiles(root, fileName, SearchOption.AllDirectories);
                    }
                    catch (UnauthorizedAccessException)
                    {
                    }
                    catch (DirectoryNotFoundException)
                    {
                    }

                    foreach (var found in discovered)
                    {
                        existing.Add(found);
                        matched = true;
                    }

                    if (matched)
                    {
                        break;
                    }
                }

                if (matched)
                {
                    break;
                }
            }

            if (!matched)
            {
                missing.Add(string.Format(CultureInfo.InvariantCulture, NamePatterns[0], day));
            }

            day = day.AddDays(1);
        }

        return new FileResolveResult(existing.OrderBy(x => x).ToList(), missing, allCandidates.ToList());
    }

    public FileResolveResult ResolveFiles(string baseDirectory, DateTime start, DateTime end)
        => ResolveFiles([baseDirectory], start, end);
}

public sealed record FileResolveResult(
    IReadOnlyList<string> ExistingFiles,
    IReadOnlyList<string> MissingFiles,
    IReadOnlyList<string> AllCandidates);
