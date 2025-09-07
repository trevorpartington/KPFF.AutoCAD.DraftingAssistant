using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Utilities;
using System.Collections.Concurrent;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// High-performance caching service for viewport boundary calculations.
/// Eliminates expensive viewport polygon calculations through intelligent caching.
/// </summary>
public class ViewportBoundaryCacheService
{
    private readonly ILogger _logger;
    private readonly ConcurrentDictionary<string, ViewportBoundaryCacheEntry> _cache = new();
    private readonly object _cacheLock = new object();

    public ViewportBoundaryCacheService(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Represents cached viewport boundary data
    /// </summary>
    public class ViewportBoundaryCacheEntry
    {
        public string CacheKey { get; set; } = string.Empty;
        public Point3dCollection Boundary { get; set; } = new();
        public DateTime CacheCreated { get; set; }
        public ViewportProperties Properties { get; set; } = new();
        public bool IsValid { get; set; } = true;

        public bool ShouldInvalidate(ViewportProperties currentProperties)
        {
            return !IsValid || !Properties.Equals(currentProperties);
        }

        public override string ToString()
        {
            return $"ViewportCache[Key={CacheKey}, Points={Boundary.Count}, Scale={Properties.CustomScale:F3}, Created={CacheCreated:HH:mm:ss.fff}]";
        }
    }

    /// <summary>
    /// Viewport properties used for cache validation
    /// </summary>
    public class ViewportProperties
    {
        public Point3d ViewCenter { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double CustomScale { get; set; }
        public bool NonRectClipOn { get; set; }
        public ObjectId ClipEntityId { get; set; }

        public override bool Equals(object? obj)
        {
            if (obj is not ViewportProperties other) return false;
            
            return ViewCenter.IsEqualTo(other.ViewCenter, Tolerance.Global) &&
                   Math.Abs(Width - other.Width) < Tolerance.Global.EqualPoint &&
                   Math.Abs(Height - other.Height) < Tolerance.Global.EqualPoint &&
                   Math.Abs(CustomScale - other.CustomScale) < 1e-6 &&
                   NonRectClipOn == other.NonRectClipOn &&
                   ClipEntityId == other.ClipEntityId;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(
                ViewCenter.GetHashCode(),
                Width.GetHashCode(),
                Height.GetHashCode(),
                CustomScale.GetHashCode(),
                NonRectClipOn.GetHashCode(),
                ClipEntityId.GetHashCode()
            );
        }

        public static ViewportProperties FromViewport(Viewport viewport)
        {
            return new ViewportProperties
            {
                ViewCenter = new Point3d(viewport.ViewCenter.X, viewport.ViewCenter.Y, 0),
                Width = viewport.Width,
                Height = viewport.Height,
                CustomScale = viewport.CustomScale,
                NonRectClipOn = viewport.NonRectClipOn,
                ClipEntityId = viewport.NonRectClipEntityId
            };
        }

        public override string ToString()
        {
            return $"VP[{Width:F1}×{Height:F1}, Scale={CustomScale:F3}, Clip={NonRectClipOn}]";
        }
    }

    /// <summary>
    /// Gets cached viewport boundary or calculates and caches if not available/invalid
    /// </summary>
    /// <param name="viewport">The viewport to get boundary for</param>
    /// <param name="transaction">Transaction for database access</param>
    /// <param name="layoutName">Layout name for cache key differentiation</param>
    /// <returns>Cached viewport boundary points</returns>
    public Point3dCollection GetOrCalculateViewportBoundary(
        Viewport viewport,
        Transaction transaction,
        string layoutName)
    {
        if (viewport == null)
        {
            _logger.LogWarning("Cannot cache boundary for null viewport");
            return new Point3dCollection();
        }

        // Generate cache key combining layout, viewport properties, and object ID
        var properties = ViewportProperties.FromViewport(viewport);
        var cacheKey = GenerateCacheKey(layoutName, viewport.ObjectId, properties);

        lock (_cacheLock)
        {
            // Check if we have valid cached data
            if (_cache.TryGetValue(cacheKey, out var existingEntry))
            {
                if (!existingEntry.ShouldInvalidate(properties))
                {
                    _logger.LogDebug($"Viewport boundary cache HIT: {existingEntry}");
                    return ClonePointCollection(existingEntry.Boundary);
                }
                else
                {
                    _logger.LogDebug($"Viewport boundary cache INVALIDATED: {existingEntry}");
                    _cache.TryRemove(cacheKey, out _);
                }
            }
            else
            {
                _logger.LogDebug($"Viewport boundary cache MISS: {properties}");
            }

            // Calculate new boundary and cache it
            var startTime = DateTime.Now;
            var boundary = ViewportBoundaryCalculator.GetViewportFootprint(viewport, transaction);
            var elapsed = DateTime.Now - startTime;

            if (boundary != null && boundary.Count > 0)
            {
                var newEntry = new ViewportBoundaryCacheEntry
                {
                    CacheKey = cacheKey,
                    Boundary = ClonePointCollection(boundary),
                    CacheCreated = DateTime.Now,
                    Properties = properties,
                    IsValid = true
                };

                _cache[cacheKey] = newEntry;
                _logger.LogInformation($"Viewport boundary cache STORED ({elapsed.TotalMilliseconds:F1}ms): {newEntry}");
                return ClonePointCollection(boundary);
            }
            else
            {
                _logger.LogWarning($"Failed to calculate viewport boundary for caching");
                return new Point3dCollection();
            }
        }
    }

    /// <summary>
    /// Generates a unique cache key for a viewport
    /// </summary>
    private string GenerateCacheKey(string layoutName, ObjectId viewportId, ViewportProperties properties)
    {
        // Use a combination of layout name, viewport ID, and a hash of properties
        // This ensures uniqueness even if viewports are similar but in different layouts
        var propertiesHash = properties.GetHashCode().ToString("X8");
        return $"{layoutName}:{viewportId.Handle.Value}:{propertiesHash}";
    }

    /// <summary>
    /// Creates a deep copy of a Point3dCollection to prevent cache corruption
    /// </summary>
    private static Point3dCollection ClonePointCollection(Point3dCollection original)
    {
        var clone = new Point3dCollection();
        foreach (Point3d point in original)
        {
            clone.Add(point);
        }
        return clone;
    }

    /// <summary>
    /// Invalidates all cached boundaries for a specific layout
    /// </summary>
    public void InvalidateLayout(string layoutName)
    {
        if (string.IsNullOrEmpty(layoutName)) return;

        lock (_cacheLock)
        {
            var keysToRemove = _cache.Keys
                .Where(key => key.StartsWith($"{layoutName}:", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var removedCount = 0;
            foreach (var key in keysToRemove)
            {
                if (_cache.TryRemove(key, out var entry))
                {
                    removedCount++;
                    _logger.LogDebug($"Viewport boundary cache entry invalidated: {entry}");
                }
            }

            if (removedCount > 0)
            {
                _logger.LogDebug($"Invalidated {removedCount} viewport boundary cache entries for layout '{layoutName}'");
            }
        }
    }

    /// <summary>
    /// Invalidates a specific viewport boundary cache
    /// </summary>
    public void InvalidateViewport(string layoutName, ObjectId viewportId)
    {
        if (string.IsNullOrEmpty(layoutName)) return;

        lock (_cacheLock)
        {
            var keysToRemove = _cache.Keys
                .Where(key => key.StartsWith($"{layoutName}:{viewportId.Handle.Value}:", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var removedCount = 0;
            foreach (var key in keysToRemove)
            {
                if (_cache.TryRemove(key, out var entry))
                {
                    removedCount++;
                    _logger.LogDebug($"Viewport boundary cache entry invalidated: {entry}");
                }
            }

            if (removedCount > 0)
            {
                _logger.LogDebug($"Invalidated {removedCount} viewport boundary cache entries for viewport {viewportId.Handle.Value}");
            }
        }
    }

    /// <summary>
    /// Clears all cached viewport boundaries
    /// </summary>
    public void ClearCache()
    {
        lock (_cacheLock)
        {
            var count = _cache.Count;
            _cache.Clear();
            _logger.LogInformation($"Cleared {count} viewport boundary cache entries");
        }
    }

    /// <summary>
    /// Gets cache statistics for monitoring
    /// </summary>
    public ViewportCacheStatistics GetCacheStatistics()
    {
        lock (_cacheLock)
        {
            var stats = new ViewportCacheStatistics
            {
                TotalEntries = _cache.Count,
                ValidEntries = _cache.Values.Count(e => e.IsValid),
                AveragePointsPerBoundary = _cache.Values.Any() ? _cache.Values.Average(e => e.Boundary.Count) : 0,
                OldestCacheTime = _cache.Values.Any() ? _cache.Values.Min(e => e.CacheCreated) : DateTime.MinValue,
                NewestCacheTime = _cache.Values.Any() ? _cache.Values.Max(e => e.CacheCreated) : DateTime.MinValue,
                LayoutBreakdown = _cache.Keys
                    .GroupBy(key => key.Split(':')[0])
                    .ToDictionary(g => g.Key, g => g.Count())
            };

            return stats;
        }
    }
}

/// <summary>
/// Statistics about the viewport boundary cache
/// </summary>
public class ViewportCacheStatistics
{
    public int TotalEntries { get; set; }
    public int ValidEntries { get; set; }
    public double AveragePointsPerBoundary { get; set; }
    public DateTime OldestCacheTime { get; set; }
    public DateTime NewestCacheTime { get; set; }
    public Dictionary<string, int> LayoutBreakdown { get; set; } = new();

    public override string ToString()
    {
        var layoutInfo = LayoutBreakdown.Any() 
            ? $", Layouts: {string.Join(", ", LayoutBreakdown.Select(kvp => $"{kvp.Key}({kvp.Value})"))})"
            : "";
        
        return $"Viewport Cache Stats: {ValidEntries}/{TotalEntries} valid entries, {AveragePointsPerBoundary:F1} avg points{layoutInfo}";
    }
}