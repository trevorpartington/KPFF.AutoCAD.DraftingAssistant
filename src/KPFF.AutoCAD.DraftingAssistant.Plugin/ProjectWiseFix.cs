using Autodesk.AutoCAD.ApplicationServices;
using System;
using System.Threading.Tasks;

namespace KPFF.AutoCAD.DraftingAssistant.Plugin;

/// <summary>
/// Simple ProjectWise compatibility fix
/// Automatically triggers ESC to force ProjectWise ODMA loading before any drawing operations
/// </summary>
public static class ProjectWiseFix
{
    private static bool _hasTriggered = false;

    /// <summary>
    /// Triggers ProjectWise initialization by sending ESC/Cancel commands
    /// Should be called during plugin startup, after Application.Idle
    /// </summary>
    public static async Task TriggerProjectWiseInitialization()
    {
        if (_hasTriggered) 
        {
            System.Diagnostics.Debug.WriteLine("ProjectWiseFix: Already triggered - skipping");
            return;
        }
        _hasTriggered = true;

        try
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            System.Diagnostics.Debug.WriteLine($"ProjectWiseFix: [{timestamp}] Starting ProjectWise initialization trigger...");
            
            // Reduced wait time since we're already called from Application.Idle
            await Task.Delay(100);
            
            // Send ESC to trigger ProjectWise ODMA load
            // This is what happens when user presses ESC after "Other is completed"
            var docManager = Application.DocumentManager;
            if (docManager?.MdiActiveDocument != null)
            {
                timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                System.Diagnostics.Debug.WriteLine($"ProjectWiseFix: [{timestamp}] Sending ESC to trigger ProjectWise ODMA loading...");
                
                // Send ESC character (\x1B) - this is the correct way to simulate ESC key press
                docManager.MdiActiveDocument.SendStringToExecute("\x1B", false, false, false);
                
                // Send a second ESC to ensure it takes effect
                await Task.Delay(50); // Small delay between ESC presses
                docManager.MdiActiveDocument.SendStringToExecute("\x1B", false, false, false);
                
                timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                System.Diagnostics.Debug.WriteLine($"ProjectWiseFix: [{timestamp}] ESC characters sent successfully");
                
                // Give ProjectWise time to load properly
                await Task.Delay(1500);
                
                timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                System.Diagnostics.Debug.WriteLine($"ProjectWiseFix: [{timestamp}] ProjectWise initialization complete");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("ProjectWiseFix: No active document available");
            }
        }
        catch (System.Exception ex)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            System.Diagnostics.Debug.WriteLine($"ProjectWiseFix: [{timestamp}] Error during ProjectWise initialization: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"ProjectWiseFix: [{timestamp}] Stack trace: {ex.StackTrace}");
            // Non-critical error - let plugin continue
        }
    }
}