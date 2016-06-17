namespace EmployeeInfoGrabber
{
    /// <summary>
    /// Privodes additional functionality for <code>string </code> type.
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Replaces specified occurances from the string.
        /// </summary>
        /// <param name="str">Original string.</param>
        /// <param name="replacebleStr">Specify any amount of strings for removing.</param>
        /// <param name="replaceWith">The string to replace with.</param>
        /// <returns>New string with replaced occurences.</returns>
        public static string ReplaceMany(this string str, string[] replacebleStr, string replaceWith)
        {
            foreach (var rmStr in replacebleStr)
            {
                str = str.Replace(rmStr, replaceWith);
            }

            return str;
        }
    }
}