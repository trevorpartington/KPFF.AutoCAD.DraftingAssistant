using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Standardized exception handling utility
/// </summary>
public static class ExceptionHandler
{
    /// <summary>
    /// Handles exceptions with standardized logging and optional user notification
    /// </summary>
    /// <param name="exception">The exception to handle</param>
    /// <param name="logger">Logger for recording the exception</param>
    /// <param name="notificationService">Optional notification service for user messages</param>
    /// <param name="context">Context description for where the exception occurred</param>
    /// <param name="showUserMessage">Whether to show a user-friendly message</param>
    /// <param name="rethrow">Whether to rethrow the exception after handling</param>
    /// <returns>True if handled successfully, false if critical error</returns>
    public static bool HandleException(
        Exception exception,
        ILogger logger,
        INotificationService? notificationService = null,
        string context = "",
        bool showUserMessage = false,
        bool rethrow = false)
    {
        try
        {
            // Log the exception with full details
            var contextMessage = string.IsNullOrEmpty(context) 
                ? "An error occurred"
                : $"Error in {context}";

            logger.LogError(contextMessage, exception);

            // Show user-friendly message if requested
            if (showUserMessage && notificationService != null)
            {
                var userMessage = GetUserFriendlyMessage(exception);
                notificationService.ShowError("Application Error", userMessage);
            }

            // Rethrow if requested
            if (rethrow)
            {
                throw exception;
            }

            return true;
        }
        catch (Exception handlingException)
        {
            // If exception handling itself fails, try basic logging
            try
            {
                if (logger is IApplicationLogger appLogger)
                {
                    appLogger.LogCritical("Critical error in exception handling", handlingException);
                }
                else
                {
                    logger.LogError("Critical error in exception handling", handlingException);
                }
            }
            catch
            {
                // Last resort - system debug output
                System.Diagnostics.Debug.WriteLine($"CRITICAL: Exception handling failed: {handlingException}");
            }

            return false;
        }
    }

    /// <summary>
    /// Safely executes an action with standardized exception handling
    /// </summary>
    public static bool TryExecute(
        Action action,
        ILogger logger,
        INotificationService? notificationService = null,
        string context = "",
        bool showUserMessage = false)
    {
        try
        {
            action();
            return true;
        }
        catch (Exception ex)
        {
            return HandleException(ex, logger, notificationService, context, showUserMessage);
        }
    }

    /// <summary>
    /// Safely executes a function with standardized exception handling
    /// </summary>
    public static (bool Success, T? Result) TryExecute<T>(
        Func<T> function,
        ILogger logger,
        INotificationService? notificationService = null,
        string context = "",
        bool showUserMessage = false)
    {
        try
        {
            var result = function();
            return (true, result);
        }
        catch (Exception ex)
        {
            var handled = HandleException(ex, logger, notificationService, context, showUserMessage);
            return (handled, default(T));
        }
    }

    /// <summary>
    /// Converts technical exceptions to user-friendly messages
    /// </summary>
    private static string GetUserFriendlyMessage(Exception exception)
    {
        return exception switch
        {
            ArgumentException => "Invalid input provided. Please check your data and try again.",
            InvalidOperationException => "This operation cannot be performed at this time. Please try again later.",
            UnauthorizedAccessException => "Access denied. You may not have permission to perform this action.",
            TimeoutException => "The operation timed out. Please try again.",
            System.IO.FileNotFoundException => "Required file not found. Please check the installation.",
            System.IO.DirectoryNotFoundException => "Required directory not found. Please check the installation.",
            OutOfMemoryException => "Insufficient memory to complete the operation. Please close other applications and try again.",
            NotSupportedException => "This operation is not supported in the current environment.",
            _ => $"An unexpected error occurred: {exception.Message}"
        };
    }

    /// <summary>
    /// Determines if an exception is critical and should terminate the application
    /// </summary>
    public static bool IsCriticalException(Exception exception)
    {
        return exception is OutOfMemoryException or
               StackOverflowException or
               AccessViolationException or
               AppDomainUnloadedException or
               BadImageFormatException or
               InvalidProgramException;
    }
}