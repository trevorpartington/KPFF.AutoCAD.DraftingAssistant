using Microsoft.Win32;
using System;
using System.IO;
using System.Reflection;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// Handles registry-based auto-loading configuration for AutoCAD plugin
/// </summary>
public static class RegistryInstaller
{
    private const string AutoCadAppsKey = @"SOFTWARE\Autodesk\AutoCAD\R25.0\ACAD-5001:409\Applications";
    private const string AppName = "KPFFDraftingAssistant";
    
    /// <summary>
    /// Registers the plugin for auto-loading with AutoCAD
    /// </summary>
    /// <param name="logger">Logger for diagnostic messages</param>
    /// <returns>True if registration successful, false otherwise</returns>
    public static bool RegisterForAutoLoad(ILogger? logger = null)
    {
        try
        {
            var assemblyPath = Assembly.GetExecutingAssembly().Location;
            if (string.IsNullOrEmpty(assemblyPath) || !File.Exists(assemblyPath))
            {
                logger?.LogError($"Cannot find assembly at: {assemblyPath}");
                return false;
            }

            // Open or create the AutoCAD applications registry key
            using var appsKey = Registry.CurrentUser.CreateSubKey(AutoCadAppsKey);
            if (appsKey == null)
            {
                logger?.LogError("Failed to open AutoCAD applications registry key");
                return false;
            }

            // Create our plugin subkey
            using var pluginKey = appsKey.CreateSubKey(AppName);
            if (pluginKey == null)
            {
                logger?.LogError($"Failed to create plugin registry key: {AppName}");
                return false;
            }

            // Set the required registry values
            pluginKey.SetValue("DESCRIPTION", "KPFF Drafting Assistant Plugin");
            pluginKey.SetValue("LOADCTRLS", 14); // Load on startup + demand load
            pluginKey.SetValue("LOADER", assemblyPath);
            pluginKey.SetValue("MANAGED", 1); // Managed .NET assembly

            logger?.LogInformation($"Successfully registered plugin for auto-loading: {assemblyPath}");
            return true;
        }
        catch (Exception ex)
        {
            logger?.LogError($"Failed to register plugin for auto-loading: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Unregisters the plugin from auto-loading
    /// </summary>
    /// <param name="logger">Logger for diagnostic messages</param>
    /// <returns>True if unregistration successful, false otherwise</returns>
    public static bool UnregisterFromAutoLoad(ILogger? logger = null)
    {
        try
        {
            using var appsKey = Registry.CurrentUser.OpenSubKey(AutoCadAppsKey, true);
            if (appsKey == null)
            {
                logger?.LogWarning("AutoCAD applications registry key not found");
                return true; // Already not registered
            }

            appsKey.DeleteSubKeyTree(AppName, false);
            logger?.LogInformation("Successfully unregistered plugin from auto-loading");
            return true;
        }
        catch (Exception ex)
        {
            logger?.LogError($"Failed to unregister plugin from auto-loading: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Checks if the plugin is currently registered for auto-loading
    /// </summary>
    /// <param name="logger">Logger for diagnostic messages</param>
    /// <returns>True if registered, false otherwise</returns>
    public static bool IsRegisteredForAutoLoad(ILogger? logger = null)
    {
        try
        {
            using var appsKey = Registry.CurrentUser.OpenSubKey(AutoCadAppsKey);
            if (appsKey == null)
            {
                return false;
            }

            using var pluginKey = appsKey.OpenSubKey(AppName);
            return pluginKey != null;
        }
        catch (Exception ex)
        {
            logger?.LogError($"Failed to check plugin registration status: {ex.Message}");
            return false;
        }
    }
}