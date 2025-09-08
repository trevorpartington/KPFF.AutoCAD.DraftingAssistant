using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KPFF.AutoCAD.DraftingAssistant.UI.Utilities
{
    public static class MessageFormatHelper
    {
        /// <summary>
        /// Creates a standardized progress message for processing operations
        /// </summary>
        /// <param name="operation">The operation name (e.g., "Auto Notes", "Excel Notes", "Title Block Attributes")</param>
        /// <param name="currentSheet">Current sheet number (1-based)</param>
        /// <param name="totalSheets">Total number of sheets</param>
        /// <returns>Formatted progress message</returns>
        public static string CreateProgressMessage(string operation, int currentSheet, int totalSheets)
        {
            return $"{operation}: Processing sheet {currentSheet} of {totalSheets}";
        }

        /// <summary>
        /// Creates a standardized progress message for plotting operations
        /// </summary>
        /// <param name="totalSheets">Total number of sheets being plotted</param>
        /// <returns>Formatted plotting message</returns>
        public static string CreatePlottingMessage(int totalSheets)
        {
            return $"Plotting {totalSheets} sheets";
        }

        /// <summary>
        /// Creates a standardized completion message for successful operations
        /// </summary>
        /// <param name="operation">The operation name</param>
        /// <param name="successfulSheets">List of successfully processed sheets</param>
        /// <param name="failedSheets">Dictionary of failed sheets and their error messages</param>
        /// <returns>Formatted completion message</returns>
        public static string CreateCompletionMessage(
            string operation, 
            List<string> successfulSheets, 
            Dictionary<string, string>? failedSheets = null)
        {
            var message = new StringBuilder();
            
            // Header with appropriate emoji
            var hasErrors = failedSheets != null && failedSheets.Count > 0;
            var emoji = hasErrors ? "⚠️" : "✅";
            message.AppendLine($"{operation}: Complete {emoji}");

            // Successful sheets section
            if (successfulSheets.Count > 0)
            {
                message.AppendLine();
                var successLabel = operation == "Plotting" ? "Successfully plotted:" : "Successfully updated:";
                message.AppendLine(successLabel);
                foreach (var sheet in successfulSheets)
                {
                    message.AppendLine($"✓ {sheet}");
                }
            }

            // Errors section
            if (hasErrors)
            {
                message.AppendLine();
                message.AppendLine("Errors:");
                foreach (var (sheet, error) in failedSheets!)
                {
                    message.AppendLine($"✗ {sheet}: {error}");
                }
            }
            else if (successfulSheets.Count > 0)
            {
                message.AppendLine();
                message.AppendLine("Errors: None");
            }

            return message.ToString().TrimEnd();
        }

        /// <summary>
        /// Creates a completion message for plotting operations with specific formatting
        /// </summary>
        /// <param name="totalSheets">Total sheets processed</param>
        /// <param name="successfulSheets">List of successfully plotted sheets</param>
        /// <param name="failedSheets">Dictionary of failed sheets and their error messages</param>
        /// <returns>Formatted plotting completion message</returns>
        public static string CreatePlottingCompletionMessage(
            int totalSheets,
            List<string> successfulSheets, 
            Dictionary<string, string>? failedSheets = null)
        {
            // For plotting, just use the standard completion format without the Results line
            return CreateCompletionMessage("Plotting", successfulSheets, failedSheets);
        }
    }
}