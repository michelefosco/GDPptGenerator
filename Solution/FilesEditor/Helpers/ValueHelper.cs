
using System.Text.RegularExpressions;

namespace FilesEditor.Helpers
{
    public class ValuesHelper
    {
        internal static bool AreAllNulls(
            object obj01, object obj02, object obj03 = null, object obj04 = null, object obj05 = null,
            object obj06 = null, object obj07 = null, object obj08 = null, object obj09 = null, object obj10 = null,
            object obj11 = null, object obj12 = null, object obj13 = null, object obj14 = null, object obj15 = null,
            object obj16 = null, object obj17 = null, object obj18 = null, object obj19 = null, object obj20 = null)
        {
            return obj01 == null && obj02 == null && obj03 == null && obj04 == null && obj05 == null
                && obj06 == null && obj07 == null && obj08 == null && obj09 == null && obj10 == null
                && obj11 == null && obj12 == null && obj13 == null && obj14 == null && obj15 == null
                && obj16 == null && obj17 == null && obj18 == null && obj19 == null && obj20 == null;
        }

        /// <summary>
        /// Restituisce true se 'input' corrisponde al pattern con carattere jolly '*'.
        /// </summary>
        internal static bool StringMatch(string input, string pattern)
        {
            if (input == null || pattern == null)
                return false;

            // Escapa i caratteri speciali regex, tranne l'asterisco
            string regexPattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";
            return Regex.IsMatch(input, regexPattern, RegexOptions.IgnoreCase);
        }

    }
}