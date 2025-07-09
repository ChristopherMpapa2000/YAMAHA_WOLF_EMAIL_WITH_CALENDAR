using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WOLF_EMAIL_WITH_CALENDAR
{
    public static class StringUtils
    {
        public static string StringEmpty(string str)
        {
            str = str.Trim();
            if (str.Length == 0)
            {
                str = null;
            }
            return str;
        }
    }
}