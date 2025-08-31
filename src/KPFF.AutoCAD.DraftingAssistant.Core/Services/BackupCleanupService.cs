using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Service for managing cleanup of backup files created during drawing updates
/// </summary>
public class BackupCleanupService
{
    private readonly ILogger _logger;
    private const string BackupExtension = ".bak.beforeupdate";

    public BackupCleanupService(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Cleans up ALL backup files in the specified directory immediately
    /// </summary>
    /// <param name="directoryPath">Directory to scan for backup files</param>
    /// <returns>Number of files cleaned up</returns>
    public int CleanupAllBackupFiles(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
        {
            _logger.LogWarning($"Directory does not exist: {directoryPath}");
            return 0;
        }

        try
        {
            var backupFiles = Directory.GetFiles(directoryPath, $"*{BackupExtension}");
            if (backupFiles.Length == 0)
            {
                _logger.LogDebug($"No backup files found in: {directoryPath}");
                return 0;
            }

            _logger.LogInformation($"Found {backupFiles.Length} backup files in: {directoryPath}");

            int cleanedCount = 0;

            foreach (string backupFile in backupFiles)
            {
                try
                {
                    File.Delete(backupFile);
                    cleanedCount++;
                    _logger.LogDebug($"Deleted backup file: {Path.GetFileName(backupFile)}");
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Failed to delete backup file {Path.GetFileName(backupFile)}: {ex.Message}");
                }
            }

            if (cleanedCount > 0)
            {
                _logger.LogInformation($"Cleaned up {cleanedCount} backup files from: {directoryPath}");
            }

            return cleanedCount;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error during backup cleanup in {directoryPath}: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Gets count of backup files in the specified directory
    /// </summary>
    /// <param name="directoryPath">Directory to scan</param>
    /// <returns>Count of backup files</returns>
    public int GetBackupFileCount(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
        {
            return 0;
        }

        try
        {
            return Directory.GetFiles(directoryPath, $"*{BackupExtension}").Length;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error counting backup files in {directoryPath}: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Gets information about backup files in the directory
    /// </summary>
    /// <param name="directoryPath">Directory to scan</param>
    /// <returns>List of backup file information</returns>
    public List<BackupFileInfo> GetBackupFileInfo(string directoryPath)
    {
        var result = new List<BackupFileInfo>();

        if (!Directory.Exists(directoryPath))
        {
            return result;
        }

        try
        {
            var backupFiles = Directory.GetFiles(directoryPath, $"*{BackupExtension}");
            
            foreach (string backupFile in backupFiles)
            {
                try
                {
                    var fileInfo = new FileInfo(backupFile);
                    result.Add(new BackupFileInfo
                    {
                        FileName = Path.GetFileName(backupFile),
                        FullPath = backupFile,
                        CreatedDate = fileInfo.CreationTime,
                        Size = fileInfo.Length
                    });
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error reading backup file info for {backupFile}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting backup file info from {directoryPath}: {ex.Message}");
        }

        return result.OrderBy(f => f.CreatedDate).ToList();
    }

    /// <summary>
    /// Deletes a specific backup file immediately
    /// </summary>
    /// <param name="backupFilePath">Full path to the backup file</param>
    /// <returns>True if successfully deleted</returns>
    public bool DeleteSpecificBackupFile(string backupFilePath)
    {
        if (!File.Exists(backupFilePath))
        {
            _logger.LogWarning($"Backup file does not exist: {backupFilePath}");
            return false;
        }

        if (!backupFilePath.EndsWith(BackupExtension))
        {
            _logger.LogWarning($"File is not a backup file: {backupFilePath}");
            return false;
        }

        try
        {
            File.Delete(backupFilePath);
            _logger.LogInformation($"Deleted backup file: {Path.GetFileName(backupFilePath)}");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Failed to delete backup file {Path.GetFileName(backupFilePath)}: {ex.Message}");
            return false;
        }
    }
}

/// <summary>
/// Information about a backup file
/// </summary>
public class BackupFileInfo
{
    public string FileName { get; set; } = string.Empty;
    public string FullPath { get; set; } = string.Empty;
    public DateTime CreatedDate { get; set; }
    public long Size { get; set; }
}