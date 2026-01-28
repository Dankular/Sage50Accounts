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
            : Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "SageConnector", "SageSDO.Interop.dll");

        if (!File.Exists(dllPath))
        {
            Console.WriteLine($"File not found: {dllPath}");
            return;
        }

        Console.WriteLine($"Analyzing: {Path.GetFullPath(dllPath)}");
        Console.WriteLine(new string('=', 70));

        AnalyzePE(dllPath);
        AnalyzeInteropTypes(dllPath);
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
        }
        catch (Exception ex)
        {
            Console.WriteLine($"PE Analysis error: {ex.Message}");
        }
    }

    static void AnalyzeInteropTypes(string dllPath)
    {
        Console.WriteLine("\n[ANALYZING INTEROP TYPES - TransactionPost Focus]");

        try
        {
            using var fs = new FileStream(dllPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var peReader = new PEReader(fs);

            if (!peReader.HasMetadata)
            {
                Console.WriteLine("  No .NET metadata found");
                return;
            }

            var mdReader = peReader.GetMetadataReader();

            // Find all types related to Transaction, Post, Purchase, Supplier
            var relevantTypes = new List<(TypeDefinitionHandle Handle, string Name, string Namespace, bool IsInterface)>();

            foreach (var typeHandle in mdReader.TypeDefinitions)
            {
                var typeDef = mdReader.GetTypeDefinition(typeHandle);
                var name = mdReader.GetString(typeDef.Name);
                var ns = mdReader.GetString(typeDef.Namespace);
                var attrs = typeDef.Attributes;
                bool isInterface = (attrs & TypeAttributes.ClassSemanticsMask) == TypeAttributes.Interface;

                if (name.Contains("Transaction", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("Post", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("Purchase", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("Supplier", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("Creditor", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("Split", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("Audit", StringComparison.OrdinalIgnoreCase))
                {
                    relevantTypes.Add((typeHandle, name, ns, isInterface));
                }
            }

            Console.WriteLine($"\n  Found {relevantTypes.Count} relevant types:\n");

            foreach (var (handle, name, ns, isInterface) in relevantTypes.OrderBy(t => t.Name))
            {
                var kind = isInterface ? "interface" : "class";
                Console.WriteLine($"  [{kind}] {ns}.{name}");

                // Get detailed info for this type
                var typeDef = mdReader.GetTypeDefinition(handle);

                // Show methods
                var methods = typeDef.GetMethods();
                var methodList = new List<string>();
                foreach (var methodHandle in methods)
                {
                    var method = mdReader.GetMethodDefinition(methodHandle);
                    var methodName = mdReader.GetString(method.Name);
                    if (!methodName.StartsWith("get_") && !methodName.StartsWith("set_") &&
                        !methodName.StartsWith(".") && methodName != "QueryInterface" &&
                        methodName != "QueryInterface" && methodName != "QueryInterface" &&
                        methodName != "QueryInterface" && methodName != "AddRef" && methodName != "Release" &&
                        methodName != "QueryInterface" && methodName != "Queryinterface" &&
                        methodName != "QueryInterface" && methodName != "Queryinterface")
                    {
                        // Get parameters
                        var sig = method.DecodeSignature(new SignatureDecoder(mdReader), null);
                        var paramNames = new List<string>();
                        foreach (var paramHandle in method.GetParameters())
                        {
                            var param = mdReader.GetParameter(paramHandle);
                            if (param.SequenceNumber > 0) // Skip return value
                            {
                                var paramName = mdReader.GetString(param.Name);
                                paramNames.Add(paramName);
                            }
                        }
                        var paramStr = paramNames.Any() ? $"({string.Join(", ", paramNames)})" : "()";
                        methodList.Add($"{methodName}{paramStr}");
                    }
                }

                if (methodList.Any())
                {
                    Console.WriteLine($"      Methods:");
                    foreach (var m in methodList.Take(20))
                    {
                        Console.WriteLine($"        - {m}");
                    }
                    if (methodList.Count > 20)
                        Console.WriteLine($"        ... and {methodList.Count - 20} more");
                }

                // Show properties
                var props = typeDef.GetProperties();
                var propList = new List<string>();
                foreach (var propHandle in props)
                {
                    var prop = mdReader.GetPropertyDefinition(propHandle);
                    var propName = mdReader.GetString(prop.Name);
                    propList.Add(propName);
                }

                if (propList.Any())
                {
                    Console.WriteLine($"      Properties:");
                    foreach (var p in propList.Take(30))
                    {
                        Console.WriteLine($"        - {p}");
                    }
                    if (propList.Count > 30)
                        Console.WriteLine($"        ... and {propList.Count - 30} more");
                }

                Console.WriteLine();
            }

            // Look for enums that might define TYPE values
            Console.WriteLine("\n[LOOKING FOR TYPE/TRANSACTION ENUMS]");
            foreach (var typeHandle in mdReader.TypeDefinitions)
            {
                var typeDef = mdReader.GetTypeDefinition(typeHandle);
                var name = mdReader.GetString(typeDef.Name);
                var baseType = typeDef.BaseType;

                // Check if it's an enum (extends System.Enum)
                if (!baseType.IsNil)
                {
                    string baseTypeName = "";
                    if (baseType.Kind == HandleKind.TypeReference)
                    {
                        var typeRef = mdReader.GetTypeReference((TypeReferenceHandle)baseType);
                        baseTypeName = mdReader.GetString(typeRef.Name);
                    }

                    if (baseTypeName == "Enum" &&
                        (name.Contains("Type", StringComparison.OrdinalIgnoreCase) ||
                         name.Contains("Trans", StringComparison.OrdinalIgnoreCase)))
                    {
                        Console.WriteLine($"\n  [enum] {name}");

                        // Get enum values
                        foreach (var fieldHandle in typeDef.GetFields())
                        {
                            var field = mdReader.GetFieldDefinition(fieldHandle);
                            var fieldName = mdReader.GetString(field.Name);
                            if (fieldName != "value__")
                            {
                                // Try to get the constant value
                                var constHandle = field.GetDefaultValue();
                                if (!constHandle.IsNil)
                                {
                                    var constant = mdReader.GetConstant(constHandle);
                                    var value = mdReader.GetBlobReader(constant.Value).ReadInt32();
                                    Console.WriteLine($"      {fieldName} = {value}");
                                }
                                else
                                {
                                    Console.WriteLine($"      {fieldName}");
                                }
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            Console.WriteLine($"  Stack: {ex.StackTrace}");
        }
    }

    static void FindSageComObjects()
    {
        Console.WriteLine("\n[SEARCHING FOR SAGE COM OBJECTS IN REGISTRY]");

        try
        {
            var sageProgIds = new List<string>();
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

            // Focus on transaction-related
            var transactionRelated = sageProgIds.Where(p =>
                p.Contains("Transaction", StringComparison.OrdinalIgnoreCase) ||
                p.Contains("Post", StringComparison.OrdinalIgnoreCase) ||
                p.Contains("Purchase", StringComparison.OrdinalIgnoreCase) ||
                p.Contains("Supplier", StringComparison.OrdinalIgnoreCase)).ToList();

            if (transactionRelated.Any())
            {
                Console.WriteLine($"\n  Transaction-related COM objects:");
                foreach (var progId in transactionRelated)
                {
                    Console.WriteLine($"    {progId}");
                }
            }

            Console.WriteLine($"\n  Total Sage COM objects: {sageProgIds.Count}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Registry search error: {ex.Message}");
        }
    }
}

// Simple signature decoder for method signatures
class SignatureDecoder : ISignatureTypeProvider<string, object?>
{
    private readonly MetadataReader _reader;

    public SignatureDecoder(MetadataReader reader) => _reader = reader;

    public string GetPrimitiveType(PrimitiveTypeCode typeCode) => typeCode.ToString();
    public string GetTypeFromDefinition(MetadataReader reader, TypeDefinitionHandle handle, byte rawTypeKind) =>
        reader.GetString(reader.GetTypeDefinition(handle).Name);
    public string GetTypeFromReference(MetadataReader reader, TypeReferenceHandle handle, byte rawTypeKind) =>
        reader.GetString(reader.GetTypeReference(handle).Name);
    public string GetTypeFromSpecification(MetadataReader reader, object? genericContext, TypeSpecificationHandle handle, byte rawTypeKind) => "TypeSpec";
    public string GetSZArrayType(string elementType) => $"{elementType}[]";
    public string GetPointerType(string elementType) => $"{elementType}*";
    public string GetByReferenceType(string elementType) => $"ref {elementType}";
    public string GetGenericInstantiation(string genericType, System.Collections.Immutable.ImmutableArray<string> typeArguments) =>
        $"{genericType}<{string.Join(", ", typeArguments)}>";
    public string GetGenericTypeParameter(object? genericContext, int index) => $"T{index}";
    public string GetGenericMethodParameter(object? genericContext, int index) => $"M{index}";
    public string GetFunctionPointerType(MethodSignature<string> signature) => "FnPtr";
    public string GetModifiedType(string modifier, string unmodifiedType, bool isRequired) => unmodifiedType;
    public string GetPinnedType(string elementType) => elementType;
    public string GetArrayType(string elementType, ArrayShape shape) => $"{elementType}[{shape.Rank}]";
}
