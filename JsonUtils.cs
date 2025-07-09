using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WOLF_EMAIL_WITH_CALENDAR
{
    public static class JsonUtils
    {
        public static JObject createJsonObject(string jsonStr)
        {
            JObject json = null;
            try
            {
                json = (JObject)JProperty.Parse(jsonStr);
            }
            catch(Exception ex)
            {
                json = new JObject();
            }
            return json;
        }
        public static JArray createJsonArray(string jsonStr)
        {
            JArray json = null;
            try
            {
                json = JArray.Parse(jsonStr);
            }
            catch (Exception ex)
            {
                
            }
            return json;
        }
    }
}