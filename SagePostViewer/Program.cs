using PeNet;
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Runtime.InteropServices;

class Program
{
    static void Main(string[] args)
    {
        string dllPath = args.Length > 0
            ? args[0]
            : Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "sg50SdoEngine.dll");

        if (!File.Exists(dllPath))
        {
            Console.WriteLine($"File not found: {dllPath}");
            return;
        }

        Console.WriteLine($"Analyzing: {Path.GetFullPath(dllPath)}");
        Console.WriteLine(new string('=', 70));

        AnalyzePE(dllPath);
        AnalyzeNetMetadata(dllPath);
        FindSageComObjects();
    }

    static void AnalyzePE(string dllPath)
    {
        try
        {
            var peFile = new PeFile(dllPath);

            Console.WriteLine("\n[BASIC INFO]");
            Console.WriteLine($"  Is 64-bit: {peFile.Is64Bit}");
            Console.WriteLine($"  Is DLL: {peFile.IsDll}");
            Console.WriteLine($"  Is .NET: {peFile.IsDotNet}");

            // Sage/Invoice related exports
            if (peFile.ExportedFunctions != null && peFile.ExportedFunctions.Length > 0)
            {
                var sageExports = peFile.ExportedFunctions
                    .Where(e => !string.IsNullOrEmpty(e.Name))
                    .Where(e =>
                        e.Name.Contains("Invoice", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.Contains("SalesOrder", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.Contains("PurchaseOrder", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.Contains("Customer", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.Contains("Supplier", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.Contains("Transaction", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.Contains("Post", StringComparison.OrdinalIgnoreCase) ||
                        e.Name.Contains("Nominal", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(e => e.Name)
                    .ToList();

                if (sageExports.Any())
                {
                    Console.WriteLine($"\n[BUSINESS-RELATED EXPORTS] ({sageExports.Count} found)");
                    foreach (var exp in sageExports.Take(60))
                    {
                        // Demangle C++ names for readability
                        var name = exp.Name;
                        if (name.Contains("CBInvoice"))
                            Console.WriteLine($"    {name}");
                        else if (name.Contains("CBSalesOrder"))
                            Console.WriteLine($"    {name}");
                        else if (name.Contains("CBPurchaseOrder"))
                            Console.WriteLine($"    {name}");
                    }
                }
            }

            // Version Info
            if (peFile.Resources?.VsVersionInfo != null)
            {
                Console.WriteLine("\n[VERSION INFO]");
                var vi = peFile.Resources.VsVersionInfo;
                if (vi.StringFileInfo?.StringTable != null)
                {
                    foreach (var table in vi.StringFileInfo.StringTable)
                    {
                        Console.WriteLine($"  Product: {table.ProductName} v{table.ProductVersion}");
                        Console.WriteLine($"  Company: {table.CompanyName}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"PE Analysis error: {ex.Message}");
        }
    }

    static void AnalyzeNetMetadata(string dllPath)
    {
        Console.WriteLine("\n[.NET METADATA EXPLORATION]");

        try
        {
            using var fs = new FileStream(dllPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var peReader = new PEReader(fs);

            if (!peReader.HasMetadata)
            {
                Console.WriteLine("  No .NET metadata found (native DLL)");
                return;
            }

            var mdReader = peReader.GetMetadataReader();

            // Assembly info
            if (mdReader.IsAssembly)
            {
                var asmDef = mdReader.GetAssemblyDefinition();
                Console.WriteLine($"  Assembly: {mdReader.GetString(asmDef.Name)} v{asmDef.Version}");
            }

            // Type definitions
            var types = new List<(string Namespace, string Name, bool IsPublic, bool IsInterface, bool IsClass)>();

            foreach (var typeHandle in mdReader.TypeDefinitions)
            {
                var typeDef = mdReader.GetTypeDefinition(typeHandle);
                var ns = mdReader.GetString(typeDef.Namespace);
                var name = mdReader.GetString(typeDef.Name);
                var attrs = typeDef.Attributes;

                bool isPublic = (attrs & TypeAttributes.VisibilityMask) == TypeAttributes.Public;
                bool isInterface = (attrs & TypeAttributes.ClassSemanticsMask) == TypeAttributes.Interface;
                bool isClass = (attrs & TypeAttributes.ClassSemanticsMask) == TypeAttributes.Class;

                if (!string.IsNullOrEmpty(name) && !name.StartsWith("<"))
                {
                    types.Add((ns, name, isPublic, isInterface, isClass));
                }
            }

            // Show public types
            var publicTypes = types.Where(t => t.IsPublic).OrderBy(t => t.Namespace).ThenBy(t => t.Name).ToList();

            if (publicTypes.Any())
            {
                Console.WriteLine($"\n  [PUBLIC TYPES] ({publicTypes.Count} total)");

                var byNamespace = publicTypes.GroupBy(t => t.Namespace).OrderBy(g => g.Key);
                foreach (var nsGroup in byNamespace.Take(20))
                {
                    Console.WriteLine($"\n    Namespace: {(string.IsNullOrEmpty(nsGroup.Key) ? "(global)" : nsGroup.Key)}");
                    foreach (var t in nsGroup.Take(30))
                    {
                        var typeKind = t.IsInterface ? "interface" : "class";
                        Console.WriteLine($"      [{typeKind}] {t.Name}");
                    }
                    if (nsGroup.Count() > 30)
                        Console.WriteLine($"      ... and {nsGroup.Count() - 30} more");
                }
            }

            // Look for Sage-specific types
            var sageTypes = types.Where(t =>
                t.Name.Contains("Invoice") ||
                t.Name.Contains("Sales") ||
                t.Name.Contains("Purchase") ||
                t.Name.Contains("Customer") ||
                t.Name.Contains("Supplier") ||
                t.Name.Contains("SDO") ||
                t.Name.Contains("Sdo") ||
                t.Name.Contains("Engine") ||
                t.Name.Contains("Company") ||
                t.Name.Contains("Nominal") ||
                t.Name.Contains("Transaction") ||
                t.Name.Contains("Post"))
                .OrderBy(t => t.Name)
                .ToList();

            if (sageTypes.Any())
            {
                Console.WriteLine($"\n  [SAGE BUSINESS TYPES] ({sageTypes.Count} found)");
                foreach (var t in sageTypes.Take(50))
                {
                    var visibility = t.IsPublic ? "public" : "internal";
                    var typeKind = t.IsInterface ? "interface" : "class";
                    Console.WriteLine($"    [{visibility} {typeKind}] {t.Namespace}.{t.Name}");
                }
            }

            // Type references (what external types does it use?)
            Console.WriteLine($"\n  [TYPE REFERENCES]");
            var typeRefs = new HashSet<string>();
            foreach (var refHandle in mdReader.TypeReferences)
            {
                var typeRef = mdReader.GetTypeReference(refHandle);
                var ns = mdReader.GetString(typeRef.Namespace);
                var name = mdReader.GetString(typeRef.Name);
                if (!string.IsNullOrEmpty(name) && (ns.Contains("Sage") || name.Contains("SDO") || name.Contains("Sdo")))
                {
                    typeRefs.Add($"{ns}.{name}");
                }
            }

            foreach (var tr in typeRefs.OrderBy(x => x).Take(30))
            {
                Console.WriteLine($"    {tr}");
            }

            // Assembly references
            Console.WriteLine($"\n  [ASSEMBLY REFERENCES]");
            foreach (var refHandle in mdReader.AssemblyReferences)
            {
                var asmRef = mdReader.GetAssemblyReference(refHandle);
                var name = mdReader.GetString(asmRef.Name);
                Console.WriteLine($"    {name} v{asmRef.Version}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  .NET Metadata error: {ex.Message}");
        }
    }

    static void FindSageComObjects()
    {
        Console.WriteLine("\n[SEARCHING FOR SAGE COM OBJECTS IN REGISTRY]");

        try
        {
            var sageProgIds = new List<string>();

            // Search HKEY_CLASSES_ROOT for Sage-related ProgIDs
            using var classesRoot = Microsoft.Win32.Registry.ClassesRoot;

            foreach (var subKeyName in classesRoot.GetSubKeyNames())
            {
                if (subKeyName.Contains("Sage", StringComparison.OrdinalIgnoreCase) ||
                    subKeyName.Contains("SDO", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        using var subKey = classesRoot.OpenSubKey(subKeyName);
                        if (subKey?.OpenSubKey("CLSID") != null)
                        {
                            sageProgIds.Add(subKeyName);
                        }
                    }
                    catch { }
                }
            }

            if (sageProgIds.Any())
            {
                Console.WriteLine($"  Found {sageProgIds.Count} Sage-related COM ProgIDs:\n");

                // Group by prefix
                var grouped = sageProgIds
                    .OrderBy(x => x)
                    .GroupBy(x => x.Split('.').FirstOrDefault() ?? x);

                foreach (var group in grouped)
                {
                    Console.WriteLine($"  [{group.Key}]");
                    foreach (var progId in group.Take(25))
                    {
                        Console.WriteLine($"    {progId}");
                    }
                    if (group.Count() > 25)
                        Console.WriteLine($"    ... and {group.Count() - 25} more");
                }

                // Look specifically for SDOEngine
                var sdoEngine = sageProgIds.FirstOrDefault(p =>
                    p.Contains("SDOEngine", StringComparison.OrdinalIgnoreCase) ||
                    p.Contains("SdoEngine", StringComparison.OrdinalIgnoreCase));

                if (sdoEngine != null)
                {
                    Console.WriteLine($"\n  ** Found SDO Engine: {sdoEngine} **");
                }
            }
            else
            {
                Console.WriteLine("  No Sage COM objects found in registry.");
                Console.WriteLine("  (Sage 50 may not be installed or COM registration missing)");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Registry search error: {ex.Message}");
        }
    }
}
