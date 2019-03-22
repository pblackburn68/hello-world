namespace FFS.MMSqlServer.Extensions
{
    public static class StringExtensions
    {
        // SQL supported quotename characters
        public enum enQuoteChar
        {
            Bracket = '[',
            SingleQuote = '\'',
            DoubleQuote = '\"',
            Parenthesis = '(',
            GreaterLess = '<',
            Brace = '{',
            BackTick = '`'
        }

        /// <summary>
        /// Append quote characters around a string similar to the SQL function QUOTENAME()
        /// </summary>
        /// <param name="str">String to quote</param>
        /// <param name="quote">Optional quote character</param>
        /// <returns>Quoted string</returns>
        public static string QuoteName(this string str, enQuoteChar quote = enQuoteChar.Bracket )
        {
            char schar, eChar;
            schar = eChar = (char)quote;

            switch (quote)
            {
                case enQuoteChar.Bracket:
                    eChar = ']';
                    break;
                case enQuoteChar.Parenthesis:
                    eChar = ')';
                    break;
                case enQuoteChar.GreaterLess:
                    eChar = '>';
                    break;
                case enQuoteChar.Brace:
                    eChar = '}';
                    break;
            }

            if ( !str.StartsWith(schar.ToString()) )
                str = schar + str;
            if ( !str.EndsWith(eChar.ToString() ))
                str += eChar;
            return str;
        }

        /// <summary>
        /// Remove [] quote characters from a string similar to the SQL function PARSENAME()
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ParseName(this string str)
        {
            if (str.StartsWith("["))
                str = str.Remove(0, 1);
            if (str.EndsWith("]"))
                str = str.Remove(str.Length - 1);
            return str;
        }
    }
}
