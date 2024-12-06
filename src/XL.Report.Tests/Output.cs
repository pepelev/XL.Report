using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml.Linq;

namespace XL.Report.Tests;

public sealed class Output : IDisposable
{
    private readonly string testName;

    private Output(string testName, Stream stream)
    {
        this.testName = testName;
        Stream = stream;
    }

    public static Output Prepare(string testName)
    {
        Directory.CreateDirectory(GetExtractionDirectoryPath(testName));
        var stream = new WriteOnlyStream(
            new FileStream(GetStreamPath(testName), FileMode.Create)
        );

        return new Output(testName, stream);
    }

    private const string ResultsDirectory = "results";

    public Stream Stream { get; }

    private static string GetStreamPath(string testName) => $"{ResultsDirectory}/{testName}.xlsx";
    private static string GetExtractionDirectoryPath(string testName) => $"{ResultsDirectory}/{testName}";
    private string StreamPath => GetStreamPath(testName);
    private string ExtractionDirectoryPath => GetExtractionDirectoryPath(testName);

    public async Task VerifyAsync([CallerFilePath] string sourceFile = "")
    {
        await Stream.DisposeAsync();
        Extract();
        await VerifyDirectory(
            path: ExtractionDirectoryPath,
            // ReSharper disable ExplicitCallerInfoArgument
            sourceFile: sourceFile
            // ReSharper restore ExplicitCallerInfoArgument
        ).UseDirectory("verified").ToTask();
    }

    private void Extract()
    {
        if (Directory.Exists(ExtractionDirectoryPath))
        {
            Directory.Delete(ExtractionDirectoryPath, recursive: true);
        }

        using var file = new FileStream(StreamPath, FileMode.Open);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
        {
            var extractedEntryPath = Path.Combine(ExtractionDirectoryPath, entry.FullName);
            var entryDirectory = Path.GetDirectoryName(extractedEntryPath) ?? throw new Exception("No directory");
            Directory.CreateDirectory(entryDirectory);
            using var target = new FileStream(extractedEntryPath, FileMode.CreateNew);
            using var entryContent = entry.Open();
            if (Path.GetExtension(entry.FullName) is ".xml" or ".rels")
            {
                using var reader = new StreamReader(entryContent, Encoding.UTF8);
                var xml = reader.ReadToEnd();
                var document = XDocument.Parse(xml);
                var prettified = Encoding.UTF8.GetBytes(document.ToString());
                target.Write(prettified);
            }
            else
            {
                entryContent.CopyTo(target);
            }
        }
    }

    public void Dispose()
    {
        Stream.Dispose();
    }
}