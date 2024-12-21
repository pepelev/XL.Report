#region Legal
// Copyright 2024 Pepelev Alexey
// 
// This file is part of XL.Report.
// 
// XL.Report is free software: you can redistribute it and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation, either
// version 3 of the License, or (at your option) any later version.
// 
// XL.Report is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License along with XL.Report.
// If not, see <https://www.gnu.org/licenses/>.
#endregion

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

    public static bool AutoVerify => false;
    private static string GetStreamPath(string testName) => $"{ResultsDirectory}/{testName}.xlsx";
    private static string GetExtractionDirectoryPath(string testName) => $"{ResultsDirectory}/{testName}";
    private string StreamPath => GetStreamPath(testName);
    private string ExtractionDirectoryPath => GetExtractionDirectoryPath(testName);

    public async Task VerifyAsync([CallerFilePath] string sourceFile = "")
    {
        await Stream.DisposeAsync();
        Extract();
        var verify = VerifyDirectory(
                path: ExtractionDirectoryPath,
                // ReSharper disable ExplicitCallerInfoArgument
                sourceFile: sourceFile
                // ReSharper restore ExplicitCallerInfoArgument
            )
            .UseDirectory("verified");

        if (AutoVerify)
        {
            verify = verify.AutoVerify();
        }

        await verify.ToTask();
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