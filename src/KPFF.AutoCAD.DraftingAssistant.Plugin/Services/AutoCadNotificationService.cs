using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;
using KPFF.AutoCAD.DraftingAssistant.Core.Services;
using KPFF.AutoCAD.DraftingAssistant.UI.Dialogs;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using System.Windows;
using System.Windows.Threading;
using WPFApp = System.Windows.Application;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin.Services;

/// <summary>
/// AutoCAD-specific notification service implementation with proper confirmation dialog
/// </summary>
public class AutoCadNotificationService : INotificationService
{
    private readonly ILogger? _logger;

    public AutoCadNotificationService(ILogger? logger = null)
    {
        _logger = logger;
    }

    public void ShowInformation(string title, string message)
    {
        try
        {
            AutoCADApp.ShowAlertDialog($"{title}\n\n{message}");
            _logger?.LogDebug($"Showed information dialog: {title}");
        }
        catch (Exception ex)
        {
            _logger?.LogError("Failed to show information dialog", ex);
            // Fallback to system message box
            FallbackToSystemMessageBox(title, message, MessageBoxImage.Information);
        }
    }

    public void ShowWarning(string title, string message)
    {
        try
        {
            AutoCADApp.ShowAlertDialog($"Warning - {title}\n\n{message}");
            _logger?.LogDebug($"Showed warning dialog: {title}");
        }
        catch (Exception ex)
        {
            _logger?.LogError("Failed to show warning dialog", ex);
            // Fallback to system message box
            FallbackToSystemMessageBox(title, message, MessageBoxImage.Warning);
        }
    }

    public void ShowError(string title, string message)
    {
        try
        {
            AutoCADApp.ShowAlertDialog($"Error - {title}\n\n{message}");
            _logger?.LogDebug($"Showed error dialog: {title}");
        }
        catch (Exception ex)
        {
            _logger?.LogError("Failed to show error dialog", ex);
            // Fallback to system message box
            FallbackToSystemMessageBox(title, message, MessageBoxImage.Error);
        }
    }

    public bool ShowConfirmation(string title, string message)
    {
        try
        {
            // Use WPF confirmation dialog for better user experience
            return ExecuteOnUIThread(() => 
            {
                try
                {
                    // Try to get the main AutoCAD window as owner
                    var owner = GetAutoCadMainWindow();
                    var result = ConfirmationDialog.ShowConfirmation(owner, title, message);
                    _logger?.LogDebug($"Showed confirmation dialog: {title}, Result: {result}");
                    return result;
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"Custom confirmation dialog failed, falling back to MessageBox: {ex.Message}");
                    // Fallback to standard MessageBox
                    var result = MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
                    return result == MessageBoxResult.Yes;
                }
            });
        }
        catch (Exception ex)
        {
            _logger?.LogError("Failed to show confirmation dialog", ex);
            // Final fallback - assume "No" for safety
            return false;
        }
    }

    private void FallbackToSystemMessageBox(string title, string message, MessageBoxImage icon)
    {
        try
        {
            ExecuteOnUIThread(() => 
            {
                MessageBox.Show(message, title, MessageBoxButton.OK, icon);
            });
        }
        catch (Exception ex)
        {
            _logger?.LogError("Even system MessageBox failed", ex);
        }
    }

    private T ExecuteOnUIThread<T>(Func<T> action)
    {
        if (WPFApp.Current?.MainWindow?.Dispatcher.CheckAccess() ?? Dispatcher.CurrentDispatcher.CheckAccess())
        {
            return action();
        }
        else
        {
            return (WPFApp.Current?.MainWindow?.Dispatcher ?? Dispatcher.CurrentDispatcher)
                .Invoke(action);
        }
    }

    private void ExecuteOnUIThread(Action action)
    {
        ExecuteOnUIThread(() => 
        {
            action();
            return 0; // Dummy return value
        });
    }

    private static Window? GetAutoCadMainWindow()
    {
        try
        {
            // Try to find the main AutoCAD window
            var autoCadProcess = System.Diagnostics.Process.GetCurrentProcess();
            var mainWindowHandle = autoCadProcess.MainWindowHandle;
            
            if (mainWindowHandle != IntPtr.Zero)
            {
                // Convert Win32 window to WPF Window if possible
                return WPFApp.Current?.MainWindow;
            }
            
            return null;
        }
        catch
        {
            return null;
        }
    }
}