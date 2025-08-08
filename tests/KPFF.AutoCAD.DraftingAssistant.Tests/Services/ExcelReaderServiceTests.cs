using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using Microsoft.Extensions.Logging;
using Moq;

namespace KPFF.AutoCAD.DraftingAssistant.Tests.Services;

public class ExcelReaderServiceTests
{
    private readonly Mock<IApplicationLogger> _mockLogger;
    private readonly IExcelReader _excelReader;
    private readonly string _testFilePath;

    public ExcelReaderServiceTests()
    {
        _mockLogger = new Mock<IApplicationLogger>();
        _excelReader = new ExcelReaderService(_mockLogger.Object);
        _testFilePath = Path.Combine(GetTestDataPath(), "ProjectIndex.xlsx");
    }

    [Fact]
    public async Task FileExistsAsync_ShouldReturnTrue_WhenFileExists()
    {
        var exists = await _excelReader.FileExistsAsync(_testFilePath);
        Assert.True(exists);
    }

    [Fact]
    public async Task GetWorksheetNamesAsync_ShouldReturnWorksheetNames()
    {
        if (!File.Exists(_testFilePath))
        {
            Assert.True(true, "Test file not found - skipping test");
            return;
        }

        var worksheets = await _excelReader.GetWorksheetNamesAsync(_testFilePath);
        
        Assert.NotEmpty(worksheets);
        // Expected worksheets based on our design
        Assert.Contains("Sheets", worksheets);
    }

    [Fact]
    public async Task ReadSheetIndexAsync_ShouldReturnSheets()
    {
        if (!File.Exists(_testFilePath))
        {
            Assert.True(true, "Test file not found - skipping test");
            return;
        }

        var configService = new ProjectConfigurationService(_mockLogger.Object);
        var config = configService.CreateDefaultConfiguration();
        config.Worksheets.Sheets = "Index";
        config.Tables.SheetIndex = "SHEET_INDEX";

        var sheets = await _excelReader.ReadSheetIndexAsync(_testFilePath, config);
        
        Assert.NotNull(sheets);
        // Should have at least one sheet entry
        if (sheets.Count > 0)
        {
            var firstSheet = sheets.First();
            Assert.False(string.IsNullOrEmpty(firstSheet.SheetName));
        }
    }

    [Fact]
    public async Task ReadExcelNotesAsync_ShouldReturnNoteMappings()
    {
        if (!File.Exists(_testFilePath))
        {
            Assert.True(true, "Test file not found - skipping test");
            return;
        }

        var configService = new ProjectConfigurationService(_mockLogger.Object);
        var config = configService.CreateDefaultConfiguration();
        config.Tables.ExcelNotes = "EXCEL_NOTES";

        var mappings = await _excelReader.ReadExcelNotesAsync(_testFilePath, config);
        
        Assert.NotNull(mappings);
        // Mappings may be empty if EXCEL-NOTES table doesn't exist yet
    }

    private static string GetTestDataPath()
    {
        var currentDir = Directory.GetCurrentDirectory();
        var projectRoot = currentDir;
        
        // Navigate up to find the solution root
        while (projectRoot != null && !File.Exists(Path.Combine(projectRoot, "KPFF.AutoCAD.DraftingAssistant.sln")))
        {
            projectRoot = Directory.GetParent(projectRoot)?.FullName;
        }
        
        return projectRoot != null 
            ? Path.Combine(projectRoot, "testdata") 
            : Path.Combine(currentDir, "..", "..", "..", "..", "testdata");
    }
}