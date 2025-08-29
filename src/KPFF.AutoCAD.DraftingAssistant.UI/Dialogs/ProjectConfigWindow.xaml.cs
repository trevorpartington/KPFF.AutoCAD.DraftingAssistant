using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows;
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

        // General tab - Excel Index Tables (informational) - hardcoded display as requested
        SheetIndexTableTextBlock.Text = _config.Tables.SheetIndex;
        ExcelNotesTableTextBlock.Text = _config.Tables.ExcelNotes;
        // NotesPatternTextBlock.Text is already set in XAML to "ABC_NOTES, AB_NOTES, A_NOTES"

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

        // Title blocks tab - File paths and information
        TitleBlockFilePathTextBox.Text = _config.TitleBlocks.TitleBlockFilePath;
        TitleBlockPatternDisplayTextBlock.Text = _config.TitleBlocks.TitleBlockPattern;
        MaxTitleBlockAttributesDisplayTextBlock.Text = _config.TitleBlocks.MaxAttributesPerTitleBlock.ToString();
        TitleBlockVisibilityDisplayTextBlock.Text = _config.TitleBlocks.VisibilityPropertyName;
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

        // Construction notes tab - Multileader styles
        _config.ConstructionNotes.MultileaderStyleNames = _multileaderStyles;
        
        // Construction notes tab - Note blocks
        _config.ConstructionNotes.NoteBlocks = _noteBlocks;

        // Construction notes tab - File paths
        _config.ConstructionNotes.NoteBlockFilePath = NoteBlockFilePathTextBox.Text.Trim();

        // Title blocks tab - File paths
        _config.TitleBlocks.TitleBlockFilePath = TitleBlockFilePathTextBox.Text.Trim();
        
        // The block attributes and other settings remain as configured (non-editable in UI)
    }

    private void BrowseExcelButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
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
        var dialog = new OpenFileDialog
        {
            Title = "Select DWG Files Directory (Select any file in the folder)",
            Filter = "All Files (*.*)|*.*",
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

            // Create backup of original file
            var backupPath = $"{_configFilePath}.backup.{DateTime.Now:yyyyMMdd-HHmmss}";
            File.Copy(_configFilePath, backupPath);

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

            MessageBox.Show($"Configuration saved successfully!\n\nBackup created: {Path.GetFileName(backupPath)}", 
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
        var dialog = new OpenFileDialog
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
        var dialog = new OpenFileDialog
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

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}