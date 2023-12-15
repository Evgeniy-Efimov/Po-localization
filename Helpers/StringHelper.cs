using System.IO;
using System.Linq;

namespace LocalizePo.Helpers
{
    public static class StringHelper
    {
        public static string Substring(this string str, string startStr, string endStr)
        {
            try
            {
                var result = str.Substring(string.IsNullOrEmpty(startStr) ? 0 : str.IndexOf(startStr) + startStr.Length);
                return result.Substring(0, result.IndexOf(endStr));
            }
            catch { return string.Empty; }
        }

        public static string NormalizeForComparing(this string str)
        {
            return string.Concat((str ?? string.Empty).Where(ch => !char.IsWhiteSpace(ch))).ToUpper();
        }

        public static string FormatFileName(this string str)
        {
            return string.Concat(string.Concat((str ?? string.Empty)
                .Select(ch => char.IsWhiteSpace(ch) ? "_" : ch.ToString()))
                .Split(Path.GetInvalidFileNameChars()));
        }
    }
}
