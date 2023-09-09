using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel4Engineers
{
    public static class StringHelpers
    {
        public static bool? ParseTextToBool(this string text)
        {
            switch (text.ToLower())
            {
                case "true":
                case "yes":
                case "1":
                    return true;
                case "false":
                case "no":
                case "0":
                    return false;
                default:
                    return null;

            }
        }
    }
}
