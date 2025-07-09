using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
namespace WOLF_START_SOM
{
    public static class ControlUtils
    {
      
        public static string getDataFromShortText(JObject data)
        {
            string result = null;
            try
            {
                JObject Data = (JObject)data["data"];
                result = (string)Data["value"];
                
            }
            catch(Exception)
            {
                result = null;
            }
            return result;
        }
        public static string getDataFromParagraph(JObject data)
        {
            string result = null;
            try
            {
                result = (string)data["value"];
            }
            catch (Exception)
            {
                result = null;
            }
            return result;
        }
        public static string getDataFromNumber(JObject data)
        {
            string result = null;
            try
            {
                result = (string)data["value"];
            }
            catch (Exception)
            {
                result = null;
            }
            return result;
        }
        public static string getDataFromCurrency(JObject template, JObject data)
        {
            string result = null;
            try
            {
                result = (string)data["value"];
                string useComma = (string)template["useComma"];
                string symbol = (string)template["symbol"];
                string symbolPosition = (string)template["symbolPosition"];
                if (useComma != null && useComma == "Y")
                {
                    result = String.Format("{0:n}", result);
                }
                if (symbol != null)
                {
                    if (symbolPosition == "E")
                    {
                        result = result + symbol;
                    }
                    else
                    {
                        result = symbol + result;
                    }
                }
            }
            catch (Exception)
            {
                result = null;
            }
            return result;
        }
        public static string getDataFromCalendar(JObject data)
        {
            string result = null;
            try
            {
                result = (string)data["value"];
            }
            catch (Exception)
            {
                result = null;
            }
            return result;
        }
        public static string getDataFromTitle(JObject data)
        {
            string result = null;
            try
            {
                result = (string)data["value"];
            }
            catch (Exception)
            {
                result = null;
            }
            return result;
        }
        public static string getDataFromList(JObject data)
        {
            string result = null;
            try
            {
                result = (string)data["value"];
            }
            catch (Exception)
            {
                result = null;
            }
            return result;
        }
        public static string getDataFromCheckbox(JObject template, JObject data)
        {
            List<string> result = null;
            try
            {
                JArray itemsResult = (JArray)data["items"];
                JObject attribute = (JObject)template["attribute"];
                JArray items = (JArray)attribute["items"];
                result = new List<string>();
                for (int i = 0; i < itemsResult.Count; i++)
                {
                    string isChecked = (string)itemsResult[i];
                    if (isChecked == "Y")
                    {
                        JObject item = (JObject)items[i];
                        result.Add((string)item["item"]);
                    }
                }
            }
            catch (Exception)
            {
                result = null;
            }
            return result == null? null : result.ToString();
        }
        public static string getDataFromMultipleChoice(JObject data)
        {
            string result = null;
            try
            {
                result = (string)data["value"];
            }
            catch (Exception)
            {
                result = null;
            }
            return result;
        }
        //public static List<CustomReportBean> getDataFromTable(JObject template, JObject data)
        //{
        //    List<CustomReportBean> resultList = new List<CustomReportBean>();
        //    try
        //    {
        //        JObject attribute = (JObject)template["attribute"];
        //        JArray column = (JArray)attribute["column"];
        //        JArray row = (JArray)data["row"];
        //        for (int i = 0; i < column.Count; i++)
        //        {
        //            JObject subTemplate = (JObject)column[i];
        //            foreach (JArray subRow in row)
        //            {
        //                JObject subData = (JObject)subRow[i];
        //                resultList.AddRange(getData(subTemplate, subData));
        //            }
        //        }
                
        //    }
        //    catch (Exception ex)
        //    {
        //        resultList = null;
        //    }
        //    return resultList;
        //}
    }
}