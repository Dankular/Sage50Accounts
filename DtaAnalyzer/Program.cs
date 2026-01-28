using System.Text;
using System.Text.RegularExpressions;

/// <summary>
/// Sage 50 DTA File Analyzer CLI
/// Usage:
///   DtaAnalyzer <accdata_path>           - Full analysis of all DTA files
///   DtaAnalyzer <accdata_path> scan      - Quick scan for version markers
///   DtaAnalyzer <file.dta> dump          - Hex dump of specific file
///   DtaAnalyzer <file.dta> guids         - Extract all GUIDs from file
///   DtaAnalyzer <accdata_path> fingerprint - Generate version fingerprint
/// </summary>
class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            ShowHelp();
            return;
        }

        var path = args[0];
        var command = args.Length > 1 ? args[1].ToLowerInvariant() : "analyze";

        if (File.Exists(path))
        {
            // Single file operations
            switch (command)
            {
                case "dump":
                    HexDump(path, args.Length > 2 ? int.Parse(args[2]) : 256);
                    break;
                case "guids":
                    ExtractGuids(path);
                    break;
                case "strings":
                    ExtractStrings(path);
                    break;
                case "structure":
                    AnalyzeStructure(path);
                    break;
                default:
                    AnalyzeStructure(path);
                    break;
            }
        }
        else if (Directory.Exists(path))
        {
            // Directory operations
            switch (command)
            {
                case "scan":
                    QuickScan(path);
                    break;
                case "fingerprint":
                    GenerateFingerprint(path);
                    break;
                case "guids":
                    ExtractAllGuids(path);
                    break;
                default:
                    FullAnalysis(path);
                    break;
            }
        }
        else
        {
            Console.WriteLine($"Error: Path not found: {path}");
        }
    }

    static void ShowHelp()
    {
        Console.WriteLine("Sage 50 DTA File Analyzer");
        Console.WriteLine("========================");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  DtaAnalyzer <accdata_path>              Full analysis of ACCDATA");
        Console.WriteLine("  DtaAnalyzer <accdata_path> scan         Quick scan for version markers");
        Console.WriteLine("  DtaAnalyzer <accdata_path> fingerprint  Generate version fingerprint");
        Console.WriteLine("  DtaAnalyzer <accdata_path> guids        Extract all GUIDs");
        Console.WriteLine("  DtaAnalyzer <file.dta> dump [bytes]     Hex dump (default 256 bytes)");
        Console.WriteLine("  DtaAnalyzer <file.dta> guids            Extract GUIDs from file");
        Console.WriteLine("  DtaAnalyzer <file.dta> strings          Extract ASCII strings");
        Console.WriteLine("  DtaAnalyzer <file.dta> structure        Analyze file structure");
    }

    static void HexDump(string filePath, int bytes)
    {
        Console.WriteLine($"Hex dump of {Path.GetFileName(filePath)} ({bytes} bytes):");
        Console.WriteLine(new string('-', 76));

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var buffer = new byte[bytes];
        var read = fs.Read(buffer, 0, bytes);

        for (int row = 0; row < read; row += 16)
        {
            Console.Write($"{row:X6}: ");

            // Hex
            for (int col = 0; col < 16; col++)
            {
                if (row + col < read)
                    Console.Write($"{buffer[row + col]:X2} ");
                else
                    Console.Write("   ");
            }

            Console.Write(" |");

            // ASCII
            for (int col = 0; col < 16 && row + col < read; col++)
            {
                char c = (char)buffer[row + col];
                Console.Write(c >= 0x20 && c < 0x7F ? c : '.');
            }

            Console.WriteLine("|");
        }

        Console.WriteLine();
        Console.WriteLine($"File size: {fs.Length} bytes");
    }

    static void ExtractGuids(string filePath)
    {
        Console.WriteLine($"GUIDs in {Path.GetFileName(filePath)}:");
        Console.WriteLine(new string('-', 50));

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var buffer = new byte[Math.Min(65536, (int)fs.Length)];
        fs.Read(buffer, 0, buffer.Length);

        var content = Encoding.ASCII.GetString(buffer);
        var guidRegex = new Regex(@"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");

        var guids = new HashSet<string>();
        foreach (Match match in guidRegex.Matches(content))
        {
            guids.Add(match.Value.ToUpperInvariant());
        }

        foreach (var guid in guids.OrderBy(g => g))
        {
            Console.WriteLine($"  {guid}");
        }

        Console.WriteLine($"\nTotal unique GUIDs: {guids.Count}");
    }

    static void ExtractStrings(string filePath)
    {
        Console.WriteLine($"ASCII strings in {Path.GetFileName(filePath)}:");
        Console.WriteLine(new string('-', 50));

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var buffer = new byte[Math.Min(65536, (int)fs.Length)];
        fs.Read(buffer, 0, buffer.Length);

        var strings = new List<(int offset, string value)>();
        var current = new StringBuilder();
        int startOffset = 0;

        for (int i = 0; i < buffer.Length; i++)
        {
            if (buffer[i] >= 0x20 && buffer[i] < 0x7F)
            {
                if (current.Length == 0) startOffset = i;
                current.Append((char)buffer[i]);
            }
            else
            {
                if (current.Length >= 4)
                {
                    strings.Add((startOffset, current.ToString().Trim()));
                }
                current.Clear();
            }
        }

        foreach (var (offset, value) in strings.Where(s => s.value.Length >= 4).Take(50))
        {
            Console.WriteLine($"  0x{offset:X4}: \"{value}\"");
        }

        Console.WriteLine($"\nTotal strings (4+ chars): {strings.Count(s => s.value.Length >= 4)}");
    }

    static void AnalyzeStructure(string filePath)
    {
        Console.WriteLine($"Structure analysis of {Path.GetFileName(filePath)}:");
        Console.WriteLine(new string('-', 60));

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        Console.WriteLine($"File size: {fs.Length} bytes");

        var buffer = new byte[Math.Min(512, (int)fs.Length)];
        fs.Read(buffer, 0, buffer.Length);

        Console.WriteLine("\nHeader fields:");
        if (buffer.Length >= 4)
        {
            var magic = BitConverter.ToUInt32(buffer, 0);
            Console.WriteLine($"  Bytes 0-3:   0x{magic:X8} ({magic})");

            // Check if it's ASCII
            if (buffer[0] >= 0x20 && buffer[0] < 0x7F)
            {
                var ascii = Encoding.ASCII.GetString(buffer, 0, 4);
                Console.WriteLine($"               ASCII: \"{ascii}\"");
            }
        }

        if (buffer.Length >= 8)
        {
            Console.WriteLine($"  Bytes 4-7:   0x{BitConverter.ToUInt32(buffer, 4):X8}");
        }

        // Look for record structure
        Console.WriteLine("\nRecord analysis:");
        int nullCount = 0;
        int recordStart = -1;
        var recordSizes = new List<int>();

        for (int i = 0; i < buffer.Length; i++)
        {
            if (buffer[i] == 0)
            {
                nullCount++;
            }
            else
            {
                if (nullCount >= 4)
                {
                    if (recordStart >= 0)
                    {
                        recordSizes.Add(i - nullCount - recordStart);
                    }
                    recordStart = i;
                }
                nullCount = 0;
            }
        }

        if (recordSizes.Count > 1)
        {
            var mostCommon = recordSizes.GroupBy(x => x).OrderByDescending(g => g.Count()).First();
            Console.WriteLine($"  Likely record size: {mostCommon.Key} bytes (appears {mostCommon.Count()} times)");
        }

        // GUIDs
        var content = Encoding.ASCII.GetString(buffer);
        var guidMatch = Regex.Match(content, @"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");
        if (guidMatch.Success)
        {
            Console.WriteLine($"\nFirst GUID found: {guidMatch.Value}");
        }
    }

    static void QuickScan(string accDataPath)
    {
        Console.WriteLine($"Quick scan of {accDataPath}");
        Console.WriteLine(new string('=', 60));

        var files = Directory.GetFiles(accDataPath, "*.DTA");
        Console.WriteLine($"Found {files.Length} DTA files\n");

        // Scan key files for version markers
        var keyFiles = new[] { "ACCSTAT.DTA", "HEADER.DTA", "NOMINAL.DTA", "TABLEMETADATA.DTA" };

        foreach (var fileName in keyFiles)
        {
            var filePath = Path.Combine(accDataPath, fileName);
            if (!File.Exists(filePath)) continue;

            Console.WriteLine($"[{fileName}]");
            try
            {
                using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var buffer = new byte[Math.Min(512, (int)fs.Length)];
                fs.Read(buffer, 0, buffer.Length);

                // Check for GUID
                var content = Encoding.ASCII.GetString(buffer);
                var guidMatch = Regex.Match(content, @"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");
                if (guidMatch.Success)
                {
                    Console.WriteLine($"  GUID: {guidMatch.Value}");
                }

                // Check magic bytes
                if (buffer.Length >= 4)
                {
                    var magic = BitConverter.ToUInt32(buffer, 0);
                    Console.WriteLine($"  Magic: 0x{magic:X8}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error: {ex.Message}");
            }
            Console.WriteLine();
        }
    }

    static void GenerateFingerprint(string accDataPath)
    {
        Console.WriteLine($"Generating version fingerprint for {accDataPath}");
        Console.WriteLine(new string('=', 60));

        var fingerprint = new Dictionary<string, string>();

        // ACCSTAT.DTA GUID
        var accstat = Path.Combine(accDataPath, "ACCSTAT.DTA");
        if (File.Exists(accstat))
        {
            using var fs = new FileStream(accstat, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var buffer = new byte[256];
            fs.Read(buffer, 0, buffer.Length);
            var content = Encoding.ASCII.GetString(buffer);
            var guidMatch = Regex.Match(content, @"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");
            if (guidMatch.Success)
            {
                fingerprint["ACCSTAT_GUID"] = guidMatch.Value.ToUpperInvariant();
            }
        }

        // File count and total size
        var dtaFiles = Directory.GetFiles(accDataPath, "*.DTA");
        fingerprint["DTA_FILE_COUNT"] = dtaFiles.Length.ToString();
        fingerprint["TOTAL_SIZE"] = dtaFiles.Sum(f => new FileInfo(f).Length).ToString();

        // TABLEMETADATA.DTA first GUID
        var tableMeta = Path.Combine(accDataPath, "TABLEMETADATA.DTA");
        if (File.Exists(tableMeta))
        {
            using var fs = new FileStream(tableMeta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var buffer = new byte[512];
            fs.Read(buffer, 0, buffer.Length);
            var content = Encoding.ASCII.GetString(buffer);
            var guidMatch = Regex.Match(content, @"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");
            if (guidMatch.Success)
            {
                fingerprint["TABLEMETADATA_GUID"] = guidMatch.Value.ToUpperInvariant();
            }
        }

        Console.WriteLine("\nFingerprint:");
        foreach (var kvp in fingerprint)
        {
            Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
        }

        Console.WriteLine("\nTo add version mapping, update KnownSchemaGuids in SdkManager:");
        if (fingerprint.TryGetValue("ACCSTAT_GUID", out var accstatGuid))
        {
            Console.WriteLine($"  [\"{accstatGuid}\"] = \"XX.X\", // TODO: Set actual version");
        }
    }

    static void ExtractAllGuids(string accDataPath)
    {
        Console.WriteLine($"Extracting all GUIDs from {accDataPath}");
        Console.WriteLine(new string('=', 60));

        var allGuids = new Dictionary<string, List<string>>();
        var dtaFiles = Directory.GetFiles(accDataPath, "*.DTA");

        foreach (var file in dtaFiles)
        {
            try
            {
                using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var buffer = new byte[Math.Min(8192, (int)fs.Length)];
                fs.Read(buffer, 0, buffer.Length);

                var content = Encoding.ASCII.GetString(buffer);
                var guidRegex = new Regex(@"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");

                foreach (Match match in guidRegex.Matches(content))
                {
                    var guid = match.Value.ToUpperInvariant();
                    if (!allGuids.ContainsKey(guid))
                        allGuids[guid] = new List<string>();
                    var fileName = Path.GetFileName(file);
                    if (!allGuids[guid].Contains(fileName))
                        allGuids[guid].Add(fileName);
                }
            }
            catch { }
        }

        Console.WriteLine($"\nFound {allGuids.Count} unique GUIDs:\n");
        foreach (var kvp in allGuids.OrderByDescending(x => x.Value.Count).ThenBy(x => x.Key))
        {
            Console.WriteLine($"{kvp.Key}");
            Console.WriteLine($"  Found in: {string.Join(", ", kvp.Value.Take(5))}{(kvp.Value.Count > 5 ? "..." : "")}");
        }
    }

    static void FullAnalysis(string accDataPath)
    {
        Console.WriteLine($"Full analysis of {accDataPath}");
        Console.WriteLine(new string('=', 60));

        QuickScan(accDataPath);
        Console.WriteLine();
        GenerateFingerprint(accDataPath);
    }
}
