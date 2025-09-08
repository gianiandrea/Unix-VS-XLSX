#!/bin/bash

# XLSX Reader Script for MSYS2 on Windows
# Written by Andrea Giani - v 0.7g
# Usage: ./xlsx_reader.sh [options] <file.xlsx>

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# Default values
EXPORT_CSV=false
EXPORT_TXT=false
OUTPUT_FILE=""
SHEET_NAME="sheet1"
XLSX_FILE=""

# Function to show help
show_help() {
    echo -e "${CYAN}XLSX Reader - Read Excel files without external libraries${NC}"
    echo ""
    echo -e "${YELLOW}Usage:${NC}"
    echo "  $0 [options] <file.xlsx>"
    echo ""
    echo -e "${YELLOW}Options:${NC}"
    echo "  -c, --csv         Export to CSV format"
    echo "  -t, --txt         Export to TXT format"
    echo "  -o, --output      Output file name (optional)"
    echo "  -s, --sheet       Sheet name/number (default: sheet1)"
    echo "  -h, --help        Show this help"
    echo ""
    echo -e "${YELLOW}Examples:${NC}"
    echo "  $0 data.xlsx"
    echo "  $0 -c data.xlsx"
    echo "  $0 -c -o output.csv data.xlsx"
    echo "  $0 -t -s sheet2 data.xlsx"
}

# Function to print colored messages
print_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

print_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -c|--csv)
            EXPORT_CSV=true
            shift
            ;;
        -t|--txt)
            EXPORT_TXT=true
            shift
            ;;
        -o|--output)
            OUTPUT_FILE="$2"
            shift 2
            ;;
        -s|--sheet)
            SHEET_NAME="$2"
            shift 2
            ;;
        -h|--help)
            show_help
            exit 0
            ;;
        *)
            if [[ -z "$XLSX_FILE" ]]; then
                XLSX_FILE="$1"
            else
                print_error "Unknown option: $1"
                exit 1
            fi
            shift
            ;;
    esac
done

# Check if file is provided
if [[ -z "$XLSX_FILE" ]]; then
    print_error "No XLSX file specified"
    show_help
    exit 1
fi

# Check if file exists
if [[ ! -f "$XLSX_FILE" ]]; then
    print_error "File not found: $XLSX_FILE"
    exit 1
fi

# Check if file has .xlsx extension
if [[ ! "$XLSX_FILE" =~ \.xlsx$ ]]; then
    print_warning "File doesn't have .xlsx extension, proceeding anyway..."
fi

# Get file size
FILE_SIZE=$(stat -c%s "$XLSX_FILE" 2>/dev/null || stat -f%z "$XLSX_FILE" 2>/dev/null || echo "unknown")

print_info "Processing file: $XLSX_FILE"
print_info "File size: $FILE_SIZE bytes"
print_info "Sheet: $SHEET_NAME"

# Create temporary directory for C# files - use a Windows-friendly path
if [ -d "/c/temp" ]; then
    TEMP_DIR=$(mktemp -d -p /c/temp)
else
    TEMP_DIR=$(mktemp -d)
fi

# Clean up any problematic characters in the path
TEMP_DIR=$(echo "$TEMP_DIR" | tr -d '\n' | tr -d '\r')
CS_FILE="$TEMP_DIR/XlsxReader.cs"
EXE_FILE="$TEMP_DIR/XlsxReader.exe"

# Cleanup function
cleanup() {
    rm -rf "$TEMP_DIR"
}
trap cleanup EXIT

# Generate C# code - FIXED VERSION
cat > "$CS_FILE" << 'EOF'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

class XlsxReader
{
    private Dictionary<int, string> sharedStrings = new Dictionary<int, string>();
    private List<List<string>> rowsData = new List<List<string>>();
    private string filePath;
    private string sheetName;
    private Stopwatch stopwatch = new Stopwatch();
    
    static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: XlsxReader.exe <file.xlsx> [sheetName] [exportFormat] [outputFile]");
            return;
        }
        
        string file = args[0];
        string sheet = args.Length > 1 ? args[1] : "sheet1";
        string format = args.Length > 2 ? args[2] : "console";
        string output = args.Length > 3 ? args[3] : "";
        
        var reader = new XlsxReader();
        reader.filePath = file;
        reader.sheetName = sheet;
        reader.ProcessFile(format, output);
    }
    
    public void ProcessFile(string exportFormat, string outputFile)
    {
        try
        {
            stopwatch.Start();
            
            Console.WriteLine("=== XLSX Reader Report ===");
            Console.WriteLine("File: " + Path.GetFileName(filePath));
            Console.WriteLine("Full path: " + Path.GetFullPath(filePath));
            Console.WriteLine("File size: " + new FileInfo(filePath).Length.ToString("N0") + " bytes");
            Console.WriteLine("Sheet: " + sheetName);
            Console.WriteLine("Started at: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            Console.WriteLine();
            
            ReadXlsxFile();
            
            stopwatch.Stop();
            
            // Statistics
            int totalRows = rowsData.Count;
            int totalCells = 0;
            foreach (var row in rowsData) totalCells += row.Count;
            int maxColumns = 0;
            foreach (var row in rowsData) 
                if (row.Count > maxColumns) maxColumns = row.Count;
            
            Console.WriteLine("=== Statistics ===");
            Console.WriteLine("Reading time: " + stopwatch.ElapsedMilliseconds + " ms");
            Console.WriteLine("Total rows: " + totalRows.ToString("N0"));
            Console.WriteLine("Maximum columns: " + maxColumns);
            Console.WriteLine("Total cells: " + totalCells.ToString("N0"));
            Console.WriteLine("Shared strings count: " + sharedStrings.Count.ToString("N0"));
            Console.WriteLine();
            
            // Export or display data
            switch (exportFormat.ToLower())
            {
                case "csv":
                    ExportToCsv(outputFile);
                    break;
                case "txt":
                    ExportToTxt(outputFile);
                    break;
                default:
                    DisplayData();
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: " + ex.Message);
            Environment.Exit(1);
        }
    }
    
    private void ReadXlsxFile()
    {
        Console.WriteLine("Opening XLSX file...");
        
        using (ZipArchive archive = ZipFile.OpenRead(filePath))
        {
            // Load shared strings
            Console.WriteLine("Loading shared strings...");
            LoadSharedStrings(archive);
            
            // Load worksheet data
            Console.WriteLine("Loading worksheet: " + sheetName + "...");
            LoadWorksheetData(archive);
        }
    }
    
    private void LoadSharedStrings(ZipArchive archive)
    {
        var sharedEntry = archive.GetEntry("xl/sharedStrings.xml");
        if (sharedEntry != null)
        {
            using (var reader = new StreamReader(sharedEntry.Open()))
            {
                var doc = XDocument.Load(reader);
                var strings = doc.Descendants().Where(e => e.Name.LocalName == "t");
                int index = 0;
                foreach (var s in strings)
                {
                    sharedStrings[index++] = s.Value;
                }
            }
        }
    }
    
    private void LoadWorksheetData(ZipArchive archive)
    {
        // Try different sheet paths
        string[] possiblePaths = {
            "xl/worksheets/" + sheetName + ".xml",
            "xl/worksheets/sheet" + sheetName + ".xml",
            "xl/worksheets/sheet1.xml"
        };
        
        ZipArchiveEntry sheetEntry = null;
        foreach (var path in possiblePaths)
        {
            sheetEntry = archive.GetEntry(path);
            if (sheetEntry != null) break;
        }
        
        if (sheetEntry == null)
        {
            Console.WriteLine("Available sheets:");
            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.StartsWith("xl/worksheets/") && entry.FullName.EndsWith(".xml"))
                {
                    Console.WriteLine("  - " + entry.FullName);
                }
            }
            throw new Exception("Sheet '" + sheetName + "' not found");
        }
        
        using (var reader = new StreamReader(sheetEntry.Open()))
        {
            var doc = XDocument.Load(reader);
            var rows = doc.Descendants().Where(e => e.Name.LocalName == "row");
            
            foreach (var row in rows)
            {
                List<string> rowValues = new List<string>();
                var cells = row.Elements().Where(e => e.Name.LocalName == "c");
                
                foreach (var cell in cells)
                {
                    string cellType = cell.Attribute("t") != null ? cell.Attribute("t").Value : null;
                    var valueElement = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
                    
                    if (valueElement != null)
                    {
                        string rawValue = valueElement.Value;
                        if (cellType == "s")
                        {
                            int sIndex;
                            if (int.TryParse(rawValue, out sIndex) && sharedStrings.ContainsKey(sIndex))
                            {
                                rowValues.Add(sharedStrings[sIndex]);
                            }
                            else
                            {
                                rowValues.Add(rawValue);
                            }
                        }
                        else
                        {
                            rowValues.Add(rawValue);
                        }
                    }
                    else
                    {
                        rowValues.Add(""); // empty cell
                    }
                }
                
                if (rowValues.Count > 0)
                {
                    rowsData.Add(rowValues);
                }
            }
        }
    }
    
    private void DisplayData()
    {
        Console.WriteLine("=== Data Preview (First 10 rows) ===");
        int displayRows = Math.Min(10, rowsData.Count);
        
        for (int i = 0; i < displayRows; i++)
        {
            var row = rowsData[i];
            List<string> displayCols = new List<string>();
            for (int j = 0; j < Math.Min(10, row.Count); j++)
            {
                displayCols.Add(row[j]);
            }
            Console.WriteLine("Row " + (i + 1).ToString("D3") + ": " + string.Join(" | ", displayCols));
            if (row.Count > 10)
            {
                Console.WriteLine("      ... and " + (row.Count - 10) + " more columns");
            }
        }
        
        if (rowsData.Count > displayRows)
        {
            Console.WriteLine("... and " + (rowsData.Count - displayRows) + " more rows");
        }
    }
    
    private void ExportToCsv(string outputFile)
    {
        if (string.IsNullOrEmpty(outputFile))
        {
            outputFile = Path.ChangeExtension(filePath, ".csv");
        }
        
        Console.WriteLine("Exporting to CSV: " + outputFile);
        
        using (var writer = new StreamWriter(outputFile, false, Encoding.UTF8))
        {
            foreach (var row in rowsData)
            {
                List<string> csvRow = new List<string>();
                foreach (var cell in row)
                {
                    if (cell.Contains(",") || cell.Contains("\"") || cell.Contains("\n"))
                    {
                        csvRow.Add("\"" + cell.Replace("\"", "\"\"") + "\"");
                    }
                    else
                    {
                        csvRow.Add(cell);
                    }
                }
                writer.WriteLine(string.Join(",", csvRow));
            }
        }
        
        Console.WriteLine("CSV export completed: " + new FileInfo(outputFile).Length.ToString("N0") + " bytes");
    }
    
    private void ExportToTxt(string outputFile)
    {
        if (string.IsNullOrEmpty(outputFile))
        {
            outputFile = Path.ChangeExtension(filePath, ".txt");
        }
        
        Console.WriteLine("Exporting to TXT: " + outputFile);
        
        using (var writer = new StreamWriter(outputFile, false, Encoding.UTF8))
        {
            foreach (var row in rowsData)
            {
                writer.WriteLine(string.Join("\t", row));
            }
        }
        
        Console.WriteLine("TXT export completed: " + new FileInfo(outputFile).Length.ToString("N0") + " bytes");
    }
}
EOF

# Check if .NET is available
if ! command -v dotnet &> /dev/null; then
    print_warning ".NET not found in PATH. Trying to locate..."
    
    # Common Windows paths for .NET
    DOTNET_PATHS=(
        "/c/Program Files/dotnet"
        "/c/Program Files (x86)/dotnet"
    )
    
    # Check PROGRAMFILES environment variables
    if [ ! -z "$PROGRAMFILES" ]; then
        DOTNET_PATHS+=("$PROGRAMFILES/dotnet")
    fi
    if [ ! -z "$PROGRAMW6432" ]; then
        DOTNET_PATHS+=("$PROGRAMW6432/dotnet")
    fi
    
    DOTNET_FOUND=0
    for DOTNET_PATH in "${DOTNET_PATHS[@]}"; do
        if [ -d "$DOTNET_PATH" ] && [ -f "$DOTNET_PATH/dotnet.exe" ]; then
            print_success "Found .NET at: $DOTNET_PATH"
            export PATH="$DOTNET_PATH:$PATH"
            DOTNET_FOUND=1
            break
        fi
    done
    
    if [ $DOTNET_FOUND -eq 0 ]; then
        print_error ".NET SDK not found. Please install it from: https://dotnet.microsoft.com/download"
        exit 1
    fi
else
    print_success ".NET found in PATH"
fi

# Check .NET version
print_info ".NET version:"
dotnet --version || print_warning "Cannot get .NET version"

print_info "Compiling C# program..."
print_info "Temporary directory: $TEMP_DIR"

# Try compilation methods in order of preference
COMPILED=0

# Method 1: Try csc.exe directly (most reliable for simple programs)
if [ $COMPILED -eq 0 ]; then
    if command -v csc.exe &> /dev/null; then
        print_info "Trying csc.exe..."
        if csc.exe /out:"$EXE_FILE" "$CS_FILE" 2>/dev/null; then
            COMPILED=1
            print_success "Compiled with csc.exe"
        else
            print_warning "Compilation with csc.exe failed"
        fi
    fi
fi

# Method 2: Try csc without .exe
if [ $COMPILED -eq 0 ]; then
    if command -v csc &> /dev/null; then
        print_info "Trying csc..."
        if csc /out:"$EXE_FILE" "$CS_FILE" 2>/dev/null; then
            COMPILED=1
            print_success "Compiled with csc"
        else
            print_warning "Compilation with csc failed"
        fi
    fi
fi

# Method 3: Try dotnet approach
if [ $COMPILED -eq 0 ]; then
    print_info "Trying dotnet compilation..."
    
    # Create a simple .csproj file in temp directory
    cat > "$TEMP_DIR/XlsxReader.csproj" << 'PROJEOF'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net9.0</TargetFramework>
    <ImplicitUsings>disable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <GenerateAssemblyInfo>false</GenerateAssemblyInfo>
    <StartupObject>XlsxReader</StartupObject>
  </PropertyGroup>
  <ItemGroup>
    <Compile Include="XlsxReader.cs" />
  </ItemGroup>
</Project>
PROJEOF
    
    # Try to build with verbose output for debugging
    print_info "Building dotnet project..."
    
    # Build with restore first to ensure all dependencies are available
    if dotnet restore "$TEMP_DIR/XlsxReader.csproj"; then
        if dotnet build "$TEMP_DIR/XlsxReader.csproj" -o "$TEMP_DIR" --verbosity minimal; then
            if [ -f "$EXE_FILE" ]; then
                COMPILED=1
                print_success "Compiled with dotnet"
            else
                # Check for other possible output file names
                if [ -f "$TEMP_DIR/XlsxReader" ]; then
                    EXE_FILE="$TEMP_DIR/XlsxReader"
                    COMPILED=1
                    print_success "Compiled with dotnet (alternative executable name)"
                else
                    print_warning "dotnet build succeeded but no executable found"
                    print_info "Files in build directory:"
                    ls -la "$TEMP_DIR/"
                fi
            fi
        else
            print_error "dotnet build failed"
        fi
    else
        print_error "dotnet restore failed"
    fi
fi

# Check if compilation was successful
if [ $COMPILED -eq 0 ]; then
    print_error "Failed to compile C# program"
    print_info "Make sure you have .NET SDK or C# compiler installed"
    exit 1
fi

# Prepare arguments for C# program
EXPORT_FORMAT="console"
if [[ "$EXPORT_CSV" == true ]]; then
    EXPORT_FORMAT="csv"
elif [[ "$EXPORT_TXT" == true ]]; then
    EXPORT_FORMAT="txt"
fi

# Convert XLSX file path to Windows format for .NET if needed
if [[ "$XLSX_FILE" == /c/* ]]; then
    XLSX_FILE_WIN=$(echo "$XLSX_FILE" | sed 's|/c/|C:\\|' | sed 's|/|\\|g')
else
    XLSX_FILE_WIN="$XLSX_FILE"
fi

# Run the C# program
print_info "Executing XLSX reader..."
echo ""

if [[ -n "$OUTPUT_FILE" ]]; then
    # Convert output file path to Windows format if needed
    if [[ "$OUTPUT_FILE" == /c/* ]]; then
        OUTPUT_FILE_WIN=$(echo "$OUTPUT_FILE" | sed 's|/c/|C:\\|' | sed 's|/|\\|g')
    else
        OUTPUT_FILE_WIN="$OUTPUT_FILE"
    fi
    "$EXE_FILE" "$XLSX_FILE_WIN" "$SHEET_NAME" "$EXPORT_FORMAT" "$OUTPUT_FILE_WIN"
else
    "$EXE_FILE" "$XLSX_FILE_WIN" "$SHEET_NAME" "$EXPORT_FORMAT"
fi

print_success "Processing completed successfully"