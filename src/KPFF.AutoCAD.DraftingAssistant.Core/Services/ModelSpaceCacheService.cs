using Autodesk.AutoCAD.DatabaseServices;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Models;
using System.Collections.Concurrent;
using System.IO;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// High-performance caching service for model space entities to eliminate duplicate scanning.
/// Provides thread-safe caching with automatic invalidation based on file modification timestamps.
/// </summary>
public class ModelSpaceCacheService
{
    private readonly ILogger _logger;
    private readonly ConcurrentDictionary<string, ModelSpaceCacheEntry> _cache = new();
    private readonly object _cacheLock = new object();

    public ModelSpaceCacheService(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Represents cached model space data for a specific drawing
    /// </summary>
    public class ModelSpaceCacheEntry
    {
        public string DrawingPath { get; set; } = string.Empty;
        public DateTime FileLastModified { get; set; }
        public DateTime CacheCreated { get; set; }
        public List<MultileaderAnalyzer.MultileaderInfo> Multileaders { get; set; } = new();
        public List<BlockAnalyzer.BlockInfo> Blocks { get; set; } = new();
        public bool IsValid { get; set; } = true;

        public bool ShouldInvalidate(DateTime currentFileModified)
        {
            return !IsValid || currentFileModified > FileLastModified;
        }

        public override string ToString()
        {
            return $"Cache[{Path.GetFileName(DrawingPath)}: {Multileaders.Count} multileaders, {Blocks.Count} blocks, Modified: {FileLastModified}, Created: {CacheCreated}]";
        }
    }

    /// <summary>
    /// Gets cached model space data or scans and caches if not available/invalid
    /// </summary>
    /// <param name="drawingPath">Full path to the drawing file</param>
    /// <param name="database">Database instance (for active/inactive drawings) or null (for closed drawings)</param>
    /// <param name="transaction">Transaction instance (for active/inactive drawings) or null (for closed drawings)</param>
    /// <param name="multileaderStyleNames">Multileader styles to filter</param>
    /// <param name="blockConfigurations">Block configurations to detect</param>
    /// <returns>Cached model space data</returns>
    public ModelSpaceCacheEntry GetOrScanModelSpace(
        string drawingPath,
        Database? database,
        Transaction? transaction,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations)
    {
        if (string.IsNullOrEmpty(drawingPath))
        {
            throw new ArgumentException("Drawing path cannot be null or empty", nameof(drawingPath));
        }

        // Normalize the path for consistent cache keys
        var normalizedPath = Path.GetFullPath(drawingPath);
        var cacheKey = normalizedPath.ToLowerInvariant();

        lock (_cacheLock)
        {
            // Get file modification time for cache validation
            DateTime fileModified = DateTime.MinValue;
            try
            {
                if (File.Exists(normalizedPath))
                {
                    fileModified = File.GetLastWriteTime(normalizedPath);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Could not get file modification time for {Path.GetFileName(normalizedPath)}: {ex.Message}");
                fileModified = DateTime.Now; // Force cache miss if file access fails
            }

            // Check if we have valid cached data
            if (_cache.TryGetValue(cacheKey, out var existingEntry))
            {
                if (!existingEntry.ShouldInvalidate(fileModified))
                {
                    _logger.LogDebug($"Cache HIT: {existingEntry}");
                    return existingEntry;
                }
                else
                {
                    _logger.LogDebug($"Cache INVALIDATED: {existingEntry} (file modified: {fileModified})");
                    _cache.TryRemove(cacheKey, out _);
                }
            }
            else
            {
                _logger.LogDebug($"Cache MISS: {Path.GetFileName(normalizedPath)}");
            }

            // Create new cache entry by scanning model space
            var newEntry = ScanModelSpaceForCache(
                normalizedPath,
                database,
                transaction,
                multileaderStyleNames,
                blockConfigurations,
                fileModified);

            if (newEntry != null)
            {
                _cache[cacheKey] = newEntry;
                _logger.LogInformation($"Cache STORED: {newEntry}");
                return newEntry;
            }
            else
            {
                _logger.LogWarning($"Failed to scan model space for caching: {Path.GetFileName(normalizedPath)}");
                
                // Return empty entry to prevent repeated failures
                var emptyEntry = new ModelSpaceCacheEntry
                {
                    DrawingPath = normalizedPath,
                    FileLastModified = fileModified,
                    CacheCreated = DateTime.Now,
                    Multileaders = new List<MultileaderAnalyzer.MultileaderInfo>(),
                    Blocks = new List<BlockAnalyzer.BlockInfo>(),
                    IsValid = false
                };
                
                _cache[cacheKey] = emptyEntry;
                return emptyEntry;
            }
        }
    }

    /// <summary>
    /// Scans model space and creates a new cache entry
    /// </summary>
    private ModelSpaceCacheEntry? ScanModelSpaceForCache(
        string drawingPath,
        Database? database,
        Transaction? transaction,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations,
        DateTime fileModified)
    {
        var startTime = DateTime.Now;
        _logger.LogDebug($"Starting model space scan for caching: {Path.GetFileName(drawingPath)}");

        try
        {
            // Handle different drawing states
            if (database != null && transaction != null)
            {
                // Active or Inactive drawing - use provided database and transaction
                return ScanOpenDrawingForCache(
                    drawingPath,
                    database,
                    transaction,
                    multileaderStyleNames,
                    blockConfigurations,
                    fileModified);
            }
            else
            {
                // Closed drawing - create external database
                return ScanClosedDrawingForCache(
                    drawingPath,
                    multileaderStyleNames,
                    blockConfigurations,
                    fileModified);
            }
        }
        catch (Exception ex)
        {
            var elapsed = DateTime.Now - startTime;
            _logger.LogError($"Error scanning model space for cache ({elapsed.TotalMilliseconds:F0}ms): {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// Scans an open drawing (active or inactive) for cache
    /// </summary>
    private ModelSpaceCacheEntry ScanOpenDrawingForCache(
        string drawingPath,
        Database database,
        Transaction transaction,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations,
        DateTime fileModified)
    {
        var startTime = DateTime.Now;
        
        var multileaderAnalyzer = new MultileaderAnalyzer(_logger);
        var blockAnalyzer = new BlockAnalyzer(_logger);

        var multileaders = multileaderAnalyzer.FindMultileadersInModelSpace(database, transaction, multileaderStyleNames);
        var blocks = blockAnalyzer.FindBlocksInModelSpace(database, transaction, blockConfigurations);

        var elapsed = DateTime.Now - startTime;
        _logger.LogInformation($"Model space scan completed ({elapsed.TotalMilliseconds:F0}ms): {multileaders.Count} multileaders, {blocks.Count} blocks");

        return new ModelSpaceCacheEntry
        {
            DrawingPath = drawingPath,
            FileLastModified = fileModified,
            CacheCreated = DateTime.Now,
            Multileaders = multileaders,
            Blocks = blocks,
            IsValid = true
        };
    }

    /// <summary>
    /// Scans a closed drawing for cache using external database
    /// </summary>
    private ModelSpaceCacheEntry ScanClosedDrawingForCache(
        string drawingPath,
        List<string>? multileaderStyleNames,
        List<NoteBlockConfiguration>? blockConfigurations,
        DateTime fileModified)
    {
        var startTime = DateTime.Now;
        
        using (var db = new Database(false, true)) // buildDefaultDrawing=false, noDocument=true
        {
            _logger.LogDebug($"Reading closed DWG file for cache: {Path.GetFileName(drawingPath)}");
            db.ReadDwgFile(drawingPath, FileOpenMode.OpenForReadAndAllShare, true, null);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var multileaderAnalyzer = new MultileaderAnalyzer(_logger);
                var blockAnalyzer = new BlockAnalyzer(_logger);

                var multileaders = multileaderAnalyzer.FindMultileadersInModelSpace(db, tr, multileaderStyleNames);
                var blocks = blockAnalyzer.FindBlocksInModelSpace(db, tr, blockConfigurations);

                tr.Commit();

                var elapsed = DateTime.Now - startTime;
                _logger.LogInformation($"Closed drawing scan completed ({elapsed.TotalMilliseconds:F0}ms): {multileaders.Count} multileaders, {blocks.Count} blocks");

                return new ModelSpaceCacheEntry
                {
                    DrawingPath = drawingPath,
                    FileLastModified = fileModified,
                    CacheCreated = DateTime.Now,
                    Multileaders = multileaders,
                    Blocks = blocks,
                    IsValid = true
                };
            }
        }
    }

    /// <summary>
    /// Invalidates cache for a specific drawing
    /// </summary>
    public void InvalidateDrawing(string drawingPath)
    {
        if (string.IsNullOrEmpty(drawingPath)) return;

        var normalizedPath = Path.GetFullPath(drawingPath);
        var cacheKey = normalizedPath.ToLowerInvariant();

        lock (_cacheLock)
        {
            if (_cache.TryRemove(cacheKey, out var entry))
            {
                _logger.LogDebug($"Cache entry invalidated: {entry}");
            }
        }
    }

    /// <summary>
    /// Clears all cached entries
    /// </summary>
    public void ClearCache()
    {
        lock (_cacheLock)
        {
            var count = _cache.Count;
            _cache.Clear();
            _logger.LogInformation($"Cleared {count} cache entries");
        }
    }

    /// <summary>
    /// Gets cache statistics for monitoring
    /// </summary>
    public CacheStatistics GetCacheStatistics()
    {
        lock (_cacheLock)
        {
            var stats = new CacheStatistics
            {
                TotalEntries = _cache.Count,
                ValidEntries = _cache.Values.Count(e => e.IsValid),
                TotalMultileaders = _cache.Values.Sum(e => e.Multileaders.Count),
                TotalBlocks = _cache.Values.Sum(e => e.Blocks.Count),
                OldestCacheTime = _cache.Values.Any() ? _cache.Values.Min(e => e.CacheCreated) : DateTime.MinValue,
                NewestCacheTime = _cache.Values.Any() ? _cache.Values.Max(e => e.CacheCreated) : DateTime.MinValue
            };

            return stats;
        }
    }
}

/// <summary>
/// Statistics about the model space cache
/// </summary>
public class CacheStatistics
{
    public int TotalEntries { get; set; }
    public int ValidEntries { get; set; }
    public int TotalMultileaders { get; set; }
    public int TotalBlocks { get; set; }
    public DateTime OldestCacheTime { get; set; }
    public DateTime NewestCacheTime { get; set; }

    public override string ToString()
    {
        return $"Cache Stats: {ValidEntries}/{TotalEntries} valid entries, {TotalMultileaders} multileaders, {TotalBlocks} blocks";
    }
}