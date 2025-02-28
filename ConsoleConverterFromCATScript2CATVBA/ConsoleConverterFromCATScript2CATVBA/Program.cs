using System;
using System.IO;
using System.Text.RegularExpressions;
class Program
{
    static void Main(string[] args)
    {
        // Check if any arguments are provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: ConvertCatScriptToVBA <input files>");
            Console.WriteLine("Example: ConvertCatScriptToVBA *.CATScript");
            return;
        }

        // Process each input file
        foreach (string inputFile in args)
        {
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"File not found: {inputFile}");
                continue;
            }

            // Change the file extension to .CATVBA while keeping the original filename
            string outputFile = Path.ChangeExtension(inputFile, ".CATVBA");

            try
            {
                // Read the content of the .CATScript file
                string scriptCode = File.ReadAllText(inputFile);

                // Convert to .CATVBA format
                string convertedCode = ConvertCatScriptToVBA(scriptCode);

                // Write the converted code to a new .CATVBA file
                File.WriteAllText(outputFile, convertedCode);

                Console.WriteLine($"Converted: {inputFile} → {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {inputFile}: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Converts .CATScript code to .CATVBA format
    /// </summary>
    static string ConvertCatScriptToVBA(string scriptCode)
    {
        // Define conversion rules using regex
        var replacements = new (string Pattern, string Replacement)[]
        {
            (@"Dim (\w+)", "Dim $1 As Object"),       // Add explicit type declarations
            (@"Set (\w+) = (\w+)\(", "Set $1 = $2("), // Ensure proper 'Set' syntax
            (@"(\w+) = CATIA\.(\w+)", "Set $1 = CATIA.$2"), // Fix CATIA object assignments
            (@"MsgBox (.+)", "MsgBox $1, vbInformation"),  // Standardize MsgBox format
            (@"Sub (\w+)", "Public Sub $1"),         // Make Sub procedures public
            (@"Function (\w+)", "Public Function $1")// Make Functions public
        };

        // Apply conversion rules sequentially
        foreach (var (pattern, replacement) in replacements)
        {
            scriptCode = Regex.Replace(scriptCode, pattern, replacement);
        }

        return scriptCode;
    }
}
