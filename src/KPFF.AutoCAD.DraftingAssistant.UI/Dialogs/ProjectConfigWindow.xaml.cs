using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;

public partial class ProjectConfigWindow : Window
{
    private ProjectConfiguration _config;
    private readonly string _configFilePath;
    private List<string> _multileaderStyles;
    private List<NoteBlockConfiguration> _noteBlocks;

    public ProjectConfiguration? UpdatedConfiguration { get; private set; }
    public bool ConfigurationSaved { get; private set; }

    public ProjectConfigWindow(ProjectConfiguration config, string configFilePath)
    {
        InitializeComponent();
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _configFilePath = configFilePath ?? throw new ArgumentNullException(nameof(configFilePath));
        
        _multileaderStyles = new List<string>();
        _noteBlocks = new List<NoteBlockConfiguration>();
        
        LoadConfigurationIntoUI();
    }

    private void LoadConfigurationIntoUI()
    {
        // General tab - Project Information
        ProjectNameTextBox.Text = _config.ProjectName;
        ClientNameTextBox.Text = _config.ClientName;
        ExcelFilePathTextBox.Text = _config.ProjectIndexFilePath;
        DwgFilePathTextBox.Text = _config.ProjectDWGFilePath;

        // General tab - Excel Index Tables (informational) - now displayed as static content
        
        // General tab - Series length configuration
        LoadSeriesLengthSetting();

        // Construction notes tab - Multileader styles
        _multileaderStyles = _config.ConstructionNotes.MultileaderStyleNames.ToList();
        UpdateStylesDisplay();
        
        // Construction notes tab - Note blocks
        _noteBlocks = _config.ConstructionNotes.NoteBlocks.ToList();
        UpdateBlocksDisplay();

        // Construction notes tab - Block information (informational display)
        MaxNotesDisplayTextBlock.Text = _config.ConstructionNotes.MaxNotesPerSheet.ToString();
        NumberAttributeDisplayTextBlock.Text = _config.ConstructionNotes.Attributes.NumberAttribute;
        NoteAttributeDisplayTextBlock.Text = _config.ConstructionNotes.Attributes.NoteAttribute;
        VisibilityPropertyDisplayTextBlock.Text = _config.ConstructionNotes.VisibilityPropertyName;

        // Construction notes tab - File paths
        NoteBlockFilePathTextBox.Text = _config.ConstructionNotes.NoteBlockFilePath;

        // Title blocks tab - File paths
        TitleBlockFilePathTextBox.Text = _config.TitleBlocks.TitleBlockFilePath;
        
        // Plotting tab - Settings
        PlotOutputDirectoryTextBox.Text = _config.Plotting.OutputDirectory;
    }

    private void UpdateStylesDisplay()
    {
        if (_multileaderStyles.Count == 0)
        {
            StylesDisplayTextBlock.Text = "No styles configured";
        }
        else
        {
            StylesDisplayTextBlock.Text = string.Join("\n", _multileaderStyles.Select(s => $"• {s}"));
        }
    }

    private void UpdateBlocksDisplay()
    {
        if (_noteBlocks.Count == 0)
        {
            BlocksDisplayTextBlock.Text = "No blocks configured";
        }
        else
        {
            BlocksDisplayTextBlock.Text = string.Join("\n", _noteBlocks.Select(nb => $"• {nb.BlockName} → {nb.AttributeName}"));
        }
    }

    private void SaveConfigurationFromUI()
    {
        // General tab - Project Information
        _config.ProjectName = ProjectNameTextBox.Text.Trim();
        _config.ClientName = ClientNameTextBox.Text.Trim();
        _config.ProjectIndexFilePath = ExcelFilePathTextBox.Text.Trim();
        _config.ProjectDWGFilePath = DwgFilePathTextBox.Text.Trim();
        
        // General tab - Series length configuration
        if (SeriesDetectionComboBox.SelectedItem is ComboBoxItem selectedItem)
        {
            _config.SheetNaming.SeriesLength = int.Parse(selectedItem.Tag.ToString()!);
        }

        // Construction notes tab - Multileader styles
        _config.ConstructionNotes.MultileaderStyleNames = _multileaderStyles;
        
        // Construction notes tab - Note blocks
        _config.ConstructionNotes.NoteBlocks = _noteBlocks;

        // Construction notes tab - File paths
        _config.ConstructionNotes.NoteBlockFilePath = NoteBlockFilePathTextBox.Text.Trim();

        // Title blocks tab - File paths
        _config.TitleBlocks.TitleBlockFilePath = TitleBlockFilePathTextBox.Text.Trim();
        
        // Plotting tab - Settings
        _config.Plotting.OutputDirectory = PlotOutputDirectoryTextBox.Text.Trim();
        
        // The block attributes and other settings remain as configured (non-editable in UI)
    }

    private void BrowseExcelButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select Excel Index File",
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            InitialDirectory = Path.GetDirectoryName(ExcelFilePathTextBox.Text)
        };

        if (dialog.ShowDialog() == true)
        {
            ExcelFilePathTextBox.Text = dialog.FileName;
        }
    }

    private void BrowseDwgButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select DWG Files Directory",
            ValidateNames = false,
            CheckFileExists = false,
            CheckPathExists = true,
            FileName = "Select Folder",
            InitialDirectory = Directory.Exists(DwgFilePathTextBox.Text) ? DwgFilePathTextBox.Text : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        if (dialog.ShowDialog() == true && !string.IsNullOrEmpty(dialog.FileName))
        {
            var selectedDirectory = Path.GetDirectoryName(dialog.FileName);
            if (!string.IsNullOrEmpty(selectedDirectory))
            {
                DwgFilePathTextBox.Text = selectedDirectory;
            }
        }
    }

    private void EditStylesButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new MultileaderStylesDialog(_multileaderStyles.ToList())
        {
            Owner = this
        };

        if (dialog.ShowDialog() == true)
        {
            _multileaderStyles = dialog.MultileaderStyles;
            UpdateStylesDisplay();
        }
    }

    private void EditBlocksButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new NoteBlocksDialog(_noteBlocks.ToList())
        {
            Owner = this
        };

        if (dialog.ShowDialog() == true)
        {
            _noteBlocks = dialog.NoteBlocks;
            UpdateBlocksDisplay();
        }
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // Validate inputs
            var validationErrors = ValidateConfiguration();
            if (validationErrors.Count > 0)
            {
                var errorMessage = "Please fix the following errors:\n\n" + string.Join("\n", validationErrors);
                MessageBox.Show(errorMessage, "Configuration Errors", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Save configuration from UI to object
            SaveConfigurationFromUI();

            // Serialize to JSON with indentation
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var jsonString = JsonSerializer.Serialize(_config, options);
            File.WriteAllText(_configFilePath, jsonString);

            UpdatedConfiguration = _config;
            ConfigurationSaved = true;
            DialogResult = true;
            Close();

            MessageBox.Show("Configuration saved successfully!", 
                           "Save Complete", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error saving configuration: {ex.Message}", "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private List<string> ValidateConfiguration()
    {
        var errors = new List<string>();

        // Validate required fields
        if (string.IsNullOrWhiteSpace(ProjectNameTextBox.Text))
            errors.Add("Project name is required");

        if (string.IsNullOrWhiteSpace(ClientNameTextBox.Text))
            errors.Add("Client name is required");

        // Validate file paths
        if (string.IsNullOrWhiteSpace(ExcelFilePathTextBox.Text))
        {
            errors.Add("Excel index file path is required");
        }
        else if (!File.Exists(ExcelFilePathTextBox.Text))
        {
            errors.Add("Excel index file does not exist");
        }

        if (string.IsNullOrWhiteSpace(DwgFilePathTextBox.Text))
        {
            errors.Add("DWG files path is required");
        }
        else if (!Directory.Exists(DwgFilePathTextBox.Text))
        {
            errors.Add("DWG files directory does not exist");
        }

        // Validate multileader styles
        if (_multileaderStyles.Count == 0)
            errors.Add("At least one multileader style is required");

        // Validate construction note block file path
        if (string.IsNullOrWhiteSpace(NoteBlockFilePathTextBox.Text))
        {
            errors.Add("Construction note block file path is required");
        }
        else if (!File.Exists(NoteBlockFilePathTextBox.Text))
        {
            errors.Add("Construction note block file does not exist");
        }
        else if (!Path.GetFileNameWithoutExtension(NoteBlockFilePathTextBox.Text).Equals("NTXX", StringComparison.OrdinalIgnoreCase))
        {
            errors.Add("Construction note block file must be named 'NTXX.dwg'");
        }

        // Validate title block file path
        if (string.IsNullOrWhiteSpace(TitleBlockFilePathTextBox.Text))
        {
            errors.Add("Title block file path is required");
        }
        else if (!File.Exists(TitleBlockFilePathTextBox.Text))
        {
            errors.Add("Title block file does not exist");
        }

        return errors;
    }

    private void BrowseNoteBlockButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select NTXX Construction Note Block File",
            Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*",
            InitialDirectory = Path.GetDirectoryName(NoteBlockFilePathTextBox.Text)
        };

        if (dialog.ShowDialog() == true)
        {
            var fileName = Path.GetFileNameWithoutExtension(dialog.FileName);
            if (fileName.Equals("NTXX", StringComparison.OrdinalIgnoreCase))
            {
                NoteBlockFilePathTextBox.Text = dialog.FileName;
            }
            else
            {
                MessageBox.Show("The selected file must be named 'NTXX.dwg'.", "Invalid File Name", 
                               MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }

    private void BrowseTitleBlockButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select Title Block DWG File",
            Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*",
            InitialDirectory = Path.GetDirectoryName(TitleBlockFilePathTextBox.Text)
        };

        if (dialog.ShowDialog() == true)
        {
            TitleBlockFilePathTextBox.Text = dialog.FileName;
        }
    }

    private void BrowseOutputDirectoryButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select Plot Output Directory",
            ValidateNames = false,
            CheckFileExists = false,
            CheckPathExists = true,
            FileName = "Select Folder",
            InitialDirectory = Directory.Exists(PlotOutputDirectoryTextBox.Text.Trim()) 
                ? PlotOutputDirectoryTextBox.Text.Trim() 
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        if (dialog.ShowDialog() == true && !string.IsNullOrEmpty(dialog.FileName))
        {
            var selectedDirectory = Path.GetDirectoryName(dialog.FileName);
            if (!string.IsNullOrEmpty(selectedDirectory))
            {
                PlotOutputDirectoryTextBox.Text = selectedDirectory;
            }
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void LoadSeriesLengthSetting()
    {
        // Set the combobox selection based on the series length
        SeriesDetectionComboBox.SelectedIndex = Math.Max(0, Math.Min(_config.SheetNaming.SeriesLength, 8));
        UpdateSeriesPreview();
    }

    private void SeriesDetectionComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateSeriesPreview();
    }

    private void UpdateSeriesPreview()
    {
        if (SeriesDetectionComboBox.SelectedItem is ComboBoxItem selectedItem)
        {
            var seriesLength = int.Parse(selectedItem.Tag.ToString()!);
            var examples = new[]
            {
                "ABC-123",
                "PV-104A", 
                "L85-UCP300",
                "A1"
            };

            var previews = new List<string>();
            foreach (var example in examples)
            {
                var parts = SimulateSeriesExtraction(example, seriesLength);
                if (parts.Length == 2)
                {
                    previews.Add($"{example} → Series: {parts[0]}, Number: {parts[1]}");
                }
            }

            SeriesPreviewTextBlock.Text = string.Join("\n", previews);
        }
    }

    private string[] SimulateSeriesExtraction(string sheetName, int seriesLength)
    {
        if (seriesLength == 0)
        {
            // Simulate auto-detect logic (simplified)
            if (sheetName.Contains('-'))
            {
                var parts = sheetName.Split('-');
                if (parts.Length >= 2)
                {
                    var series = parts[0];
                    var number = string.Join("", parts.Skip(1));
                    return new[] { series, number };
                }
            }
            else
            {
                // Simple regex-like extraction for non-hyphenated
                var letterPart = "";
                var numberPart = "";
                var foundNumber = false;
                
                for (int i = 0; i < sheetName.Length; i++)
                {
                    if (char.IsDigit(sheetName[i]) && !foundNumber)
                    {
                        foundNumber = true;
                        letterPart = sheetName.Substring(0, i);
                        numberPart = sheetName.Substring(i);
                        break;
                    }
                }
                
                if (foundNumber)
                {
                    return new[] { letterPart, numberPart };
                }
                else
                {
                    return new[] { sheetName, "" };
                }
            }
        }
        else
        {
            // Manual series length - preserve hyphens
            if (sheetName.Length <= seriesLength)
            {
                return new[] { sheetName, "" };
            }
            
            var series = sheetName.Substring(0, seriesLength);
            var number = sheetName.Substring(seriesLength);
            
            return new[] { series, number };
        }

        return Array.Empty<string>();
    }
}