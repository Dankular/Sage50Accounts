using System.Reflection;

namespace SageConnector;

public static class TypeExplorer
{
    public static void ExploreInteropAssembly()
    {
        Console.WriteLine("\n[EXPLORING SAGE SDO INTEROP ASSEMBLY]");
        Console.WriteLine(new string('=', 60));

        try
        {
            var asm = Assembly.LoadFrom("SageSDO.Interop.dll");

            var types = asm.GetExportedTypes()
                .OrderBy(t => t.Name)
                .ToList();

            Console.WriteLine($"\nFound {types.Count} public types\n");

            // Find key types
            var keyTypes = types.Where(t =>
                t.Name.Contains("Engine", StringComparison.OrdinalIgnoreCase) ||
                t.Name.Contains("Workspace", StringComparison.OrdinalIgnoreCase) ||
                t.Name.Contains("Company", StringComparison.OrdinalIgnoreCase) ||
                t.Name.Contains("Invoice", StringComparison.OrdinalIgnoreCase) ||
                t.Name.Contains("Customer", StringComparison.OrdinalIgnoreCase) ||
                t.Name.Contains("Sales", StringComparison.OrdinalIgnoreCase) ||
                t.Name.Contains("Record", StringComparison.OrdinalIgnoreCase) ||
                t.Name.Contains("Data", StringComparison.OrdinalIgnoreCase))
                .Take(50)
                .ToList();

            Console.WriteLine("[KEY TYPES]");
            foreach (var type in keyTypes)
            {
                var kind = type.IsInterface ? "interface" : type.IsClass ? "class" : "other";
                Console.WriteLine($"\n  [{kind}] {type.Name}");

                // Show properties
                var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                if (props.Any())
                {
                    Console.WriteLine("    Properties:");
                    foreach (var prop in props.Take(15))
                    {
                        Console.WriteLine($"      {prop.PropertyType.Name} {prop.Name}");
                    }
                    if (props.Length > 15)
                        Console.WriteLine($"      ... and {props.Length - 15} more");
                }

                // Show methods
                var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .Where(m => !m.Name.StartsWith("get_") && !m.Name.StartsWith("set_") &&
                               m.DeclaringType != typeof(object))
                    .ToArray();

                if (methods.Any())
                {
                    Console.WriteLine("    Methods:");
                    foreach (var method in methods.Take(15))
                    {
                        var parms = string.Join(", ", method.GetParameters()
                            .Select(p => $"{p.ParameterType.Name} {p.Name}"));
                        Console.WriteLine($"      {method.ReturnType.Name} {method.Name}({parms})");
                    }
                    if (methods.Length > 15)
                        Console.WriteLine($"      ... and {methods.Length - 15} more");
                }
            }

            // Show all type names for reference
            Console.WriteLine("\n\n[ALL TYPE NAMES]");
            foreach (var type in types)
            {
                var kind = type.IsInterface ? "I" : type.IsClass ? "C" : "?";
                Console.WriteLine($"  [{kind}] {type.Name}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
