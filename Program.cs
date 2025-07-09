
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Xml;
using WOLF_EMAIL_WITH_CALENDAR;
using WolfApprove.Model.CustomClass;
using WolfApprove.Model.Extension;
using Excel = Microsoft.Office.Interop.Excel;

namespace WOLF_EMAIL_WITH_CALENDAR
{
    class Program
    {
        private static readonly log4net.ILog log =
         log4net.LogManager.GetLogger(typeof(Program));

        private static string dbConnectionString
        {
            get
            {
                var dbConnectionString = ConfigurationManager.AppSettings["dbConnectionString"];
                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return "Integrated Security=SSPI;Initial Catalog=WolfApproveCoreISO;Data Source=172.168.1.14;User Id=sa;Password=pass@word1;";
            }
        }
        private static double iIntervalTime
        {
            get
            {
                var IntervalTime = ConfigurationManager.AppSettings["IntervalTimeMinute"];
                if (!string.IsNullOrEmpty(IntervalTime))
                {
                    return Convert.ToDouble(IntervalTime);
                }
                return -10;
            }
        }
        private static string TemplateDocumentCode
        {
            get
            {
                var TemplateDocumentCode = ConfigurationManager.AppSettings["TemplateDocumentCode"];
                if (!string.IsNullOrEmpty(TemplateDocumentCode))
                {
                    return TemplateDocumentCode;
                }
                return "";
            }
        }

        private static string TemplateEmailState
        {
            get
            {
                var TemplateEmailState = ConfigurationManager.AppSettings["TemplateEmailState"];
                if (!string.IsNullOrEmpty(TemplateEmailState))
                {
                    return TemplateEmailState;
                }
                return "";
            }
        }
        private static string TestEmail
        {
            get
            {
                var sTestEmail = ConfigurationManager.AppSettings["TestEmail"];
                if (!string.IsNullOrEmpty(sTestEmail))
                {
                    return sTestEmail;
                }
                return "";
            }
        }
        private static int MemoIDSpec
        {
            get
            {
                var iMemoIDSpec = ConfigurationManager.AppSettings["MemoIDSpec"];
                if (!string.IsNullOrEmpty(iMemoIDSpec))
                {
                    return Convert.ToInt32(iMemoIDSpec);
                }
                return 0;
            }
        }
        private static string _DebuggerMode = ConfigurationManager.AppSettings["DebuggerMode"];

        static void Main(string[] args)
        {
            try
            {

                log.Info("====== Start Process WOLF_EMAIL_WITH_CALENDAR ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                log.Info(string.Format("Run batch as :{0}", System.Security.Principal.WindowsIdentity.GetCurrent().Name));


                GetTransaction();
                Console.WriteLine(":COMPLETE");
                Console.WriteLine("exit 0");
                log.Info(":====== Completed Process WOLF_EMAIL_WITH_CALENDAR ====== ");
                log.Info("exit 0");

            }
            catch (Exception ex)
            {
                Console.WriteLine(":ERROR");
                Console.WriteLine("exit 1");

                log.Error(":ERROR");
                log.Error("message: " + ex.Message);
                log.Error("exit 1");
            }
            finally
            {
                log.Info("====== End Process WOLF_EMAIL_WITH_CALENDAR ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

            }
            if (_DebuggerMode == "T") Console.ReadKey();
        }
        public static List<CustomJsonAdvanceForm.BoxLayout_RefDoc> convertAdvanceFormToList(string advanceForm)
        {
            List<CustomJsonAdvanceForm.BoxLayout_RefDoc> listBoxLayout = new List<CustomJsonAdvanceForm.BoxLayout_RefDoc>();
            try
            {
                if (advanceForm != null)
                {
                    List<JObject> jsonAdvanceFormList = new List<JObject>();
                    JObject jsonAdvanceForm = JsonUtils.createJsonObject(advanceForm);
                    if (jsonAdvanceForm.ContainsKey("items"))
                    {
                        jsonAdvanceFormList.Add(jsonAdvanceForm);
                    }

                    if (jsonAdvanceFormList != null)
                    {
                        int iRunning = 1;

                        foreach (JObject json in jsonAdvanceFormList)
                        {
                            JArray itemsArray = (JArray)json["items"];
                            foreach (JObject jItems in itemsArray)
                            {
                                JArray jLayoutArray = (JArray)jItems["layout"];

                                CustomJsonAdvanceForm.BoxLayout_RefDoc iBoxLayout = new CustomJsonAdvanceForm.BoxLayout_RefDoc();
                                if (jLayoutArray.Count >= 1)
                                {
                                    iBoxLayout = new CustomJsonAdvanceForm.BoxLayout_RefDoc();
                                    JObject jTemplate = (JObject)jLayoutArray[0]["template"];
                                    if (jTemplate != null)
                                    {
                                        if (jTemplate.Count > 0)
                                        {

                                            iBoxLayout.Box_ID = iRunning.ToString(); iRunning++;
                                            iBoxLayout.Box_Column = "2";
                                            iBoxLayout.Box_ControlType = CustomJsonAdvanceForm.GetControlTypeByJSONKey((String)jTemplate["type"]);
                                            //iBoxLayout.Box_Control_ControlTypeCss = GetCssIconControlType(iBoxLayout.Box_ControlType.ToString());
                                            iBoxLayout.Box_Control_ControlTypeText = GetTextControlType(iBoxLayout.Box_ControlType.ToString());

                                            iBoxLayout.Box_Control_Label = (String)jTemplate["label"];
                                            iBoxLayout.Box_Control_AltLabel = (String)jTemplate["alter"];
                                            iBoxLayout.Box_Control_IsText = (String)jTemplate["istext"];
                                            iBoxLayout.Box_Control_TextValue = (String)jTemplate["textvalue"];
                                            if (jTemplate["description"] != null)
                                                iBoxLayout.Box_Control_Description = (String)jTemplate["description"];

                                            if (jTemplate["formula"] != null)
                                                iBoxLayout.Box_Control_Formula = (String)jTemplate["formula"];
                                            bool CheckAttribute = false;
                                            foreach (var valueInJtemp in jTemplate)
                                            {
                                                if (valueInJtemp.Key.ToString() == "attribute")
                                                {
                                                    CheckAttribute = true;
                                                }
                                            }
                                            if (CheckAttribute)
                                            {
                                                if (!String.IsNullOrEmpty(jTemplate["attribute"].ToString()))
                                                {
                                                    JObject jAttribute = (JObject)jTemplate["attribute"];
                                                    if (jAttribute != null)
                                                    {
                                                        if (String.IsNullOrEmpty(iBoxLayout.Box_Control_Description) && jAttribute["description"] != null)
                                                            iBoxLayout.Box_Control_Description = (String)jAttribute["description"];
                                                        iBoxLayout.Box_Control_DefaultValue = (String)jAttribute["default"];
                                                        iBoxLayout.Box_Control_MaxLength = (String)jAttribute["length"];
                                                        iBoxLayout.Box_Control_Required = (String)jAttribute["require"];

                                                        iBoxLayout.Box_Control_Min = (String)jAttribute["min"];
                                                        iBoxLayout.Box_Control_Max = (String)jAttribute["max"];
                                                        iBoxLayout.Box_Control_Comma = (String)jAttribute["useComma"];

                                                        iBoxLayout.Box_Control_Inline = (String)jAttribute["multipleLine"];
                                                        iBoxLayout.Box_Control_Summary = (String)jAttribute["summary"];

                                                        iBoxLayout.Box_Control_Decimal = (String)jAttribute["decimal"];
                                                        iBoxLayout.Box_Control_Symbol = (String)jAttribute["symbol"];
                                                        iBoxLayout.Box_Control_SymbolPosition = (String)jAttribute["symbolPosition"];
                                                        iBoxLayout.Box_Control_ValueAlign = (String)jAttribute["align"];

                                                        JObject objDate = (JObject)jAttribute["date"];
                                                        if (objDate != null)
                                                        {
                                                            iBoxLayout.Box_Control_Symbol = (String)objDate["symbol"];
                                                        }

                                                        if (jAttribute["items"] != null)
                                                        {
                                                            if (jAttribute["items"].HasValues)
                                                            {
                                                                JArray itemsItem = (JArray)jAttribute["items"];
                                                                if (itemsItem != null)
                                                                {
                                                                    foreach (JObject kItem in itemsItem)
                                                                    {
                                                                        if (!String.IsNullOrEmpty(iBoxLayout.Box_Control_Item))
                                                                            iBoxLayout.Box_Control_Item += ",";
                                                                        iBoxLayout.Box_Control_Item += $"{kItem["item"]}";
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        JArray itemsColumn = (JArray)jAttribute["column"];
                                                        if (itemsColumn != null)
                                                        {

                                                            JObject jDataColumn = (JObject)jLayoutArray[0]["data"];

                                                            if ((String)jTemplate["type"].ToString() != "an")
                                                            {
                                                                #region | Column |

                                                                if (itemsColumn.Count > 0)
                                                                {
                                                                    iBoxLayout.Box_Control_Column = new List<CustomJsonAdvanceForm.ColumnTable_RefDoc>();
                                                                    iBoxLayout.Box_Control_Table = new List<CustomJsonAdvanceForm.ColumnTable_RefDoc>();
                                                                    int columnCount = 0;
                                                                    foreach (JObject kItem in itemsColumn)
                                                                    {

                                                                        CustomJsonAdvanceForm.ColumnTable_RefDoc iColumnTable = new CustomJsonAdvanceForm.ColumnTable_RefDoc();

                                                                        iColumnTable.Column_Label = (String)kItem["label"];
                                                                        iColumnTable.Column_AltLabel = (String)kItem["alter"];
                                                                        if (!String.IsNullOrEmpty(kItem["control"]["template"].ToString()))
                                                                        {
                                                                            JObject ktemplate = (JObject)kItem["control"]["template"];
                                                                            if (ktemplate != null)
                                                                            {
                                                                                iColumnTable.Column_ControlType = CustomJsonAdvanceForm.GetControlTypeByJSONKey((String)ktemplate["type"]);
                                                                                //iColumnTable.Column_ControlTypeCss = GetCssIconControlType(iColumnTable.Column_ControlType.ToString());
                                                                                iColumnTable.Column_ControlTypeText = GetTextControlType(iColumnTable.Column_ControlType.ToString());
                                                                                if (ktemplate["description"] != null)
                                                                                    iColumnTable.Column_Description = (String)ktemplate["description"];

                                                                                if (!String.IsNullOrEmpty(ktemplate["attribute"].ToString()))
                                                                                {
                                                                                    JObject kAttribute = (JObject)ktemplate["attribute"];
                                                                                    if (kAttribute != null)
                                                                                    {
                                                                                        if (String.IsNullOrEmpty(iColumnTable.Column_Description))
                                                                                            iColumnTable.Column_Description = (String)kAttribute["description"];
                                                                                        iColumnTable.Column_DefaultValue = (String)kAttribute["default"];

                                                                                        if (kAttribute["items"] != null)
                                                                                        {
                                                                                            if (kAttribute["items"].HasValues)
                                                                                            {
                                                                                                JArray itemsItemInColumn = (JArray)kAttribute["items"];
                                                                                                if (itemsItemInColumn != null)
                                                                                                {
                                                                                                    foreach (JObject lItem in itemsItemInColumn)
                                                                                                    {
                                                                                                        if (!String.IsNullOrEmpty(iColumnTable.Column_Item))
                                                                                                            iColumnTable.Column_Item += ",";
                                                                                                        iColumnTable.Column_Item += $"{lItem["item"]}";
                                                                                                    }

                                                                                                }
                                                                                            }
                                                                                        }

                                                                                    }
                                                                                }
                                                                            }
                                                                        }

                                                                        //JValue tableRow = (JValue)jDataColumn["row"];
                                                                        //if (tableRow.Value != null)
                                                                        //{
                                                                        JArray tableRowArr = jDataColumn["row"].Count() > 0 ?
                                                                        (JArray)jDataColumn["row"] : null;
                                                                        if (tableRowArr != null)
                                                                        {
                                                                            iBoxLayout.Box_Control_TableRowCount = tableRowArr.Count;
                                                                            int rowCount = 0;
                                                                            foreach (JArray dataRow in tableRowArr)
                                                                            {
                                                                                CustomJsonAdvanceForm.ColumnTable_RefDoc iTable = new CustomJsonAdvanceForm.ColumnTable_RefDoc();
                                                                                iTable.Box_Control_RowIndex = rowCount;
                                                                                iTable.Table_Name = iBoxLayout.Box_Control_Label;
                                                                                iTable.Column_Label = (String)kItem["label"];
                                                                                iTable.Column_AltLabel = (String)kItem["alter"];
                                                                                iTable.Box_Control_Value = (String)dataRow[columnCount]["value"];
                                                                                iTable.Column_ControlType = iColumnTable.Column_ControlType;
                                                                                iTable.Column_ControlTypeCss = iColumnTable.Column_ControlTypeCss;
                                                                                iTable.Column_ControlTypeText = iColumnTable.Column_ControlTypeText;
                                                                                JArray DataItemValue = (JArray)dataRow[columnCount]["item"];
                                                                                JArray DataItem = (JArray)kItem["control"]["template"]["attribute"]["items"];
                                                                                if (DataItemValue != null)
                                                                                {
                                                                                    foreach (JValue kRowItem in DataItemValue)
                                                                                    {
                                                                                        if (!String.IsNullOrEmpty(iTable.Box_Control_ItemValue))
                                                                                            iTable.Box_Control_ItemValue += ",";
                                                                                        iTable.Box_Control_ItemValue += $"{kRowItem}";

                                                                                    }
                                                                                }
                                                                                if (DataItem != null)
                                                                                {
                                                                                    string[] selectedValue = (string[])null;
                                                                                    if (!string.IsNullOrEmpty(iTable.Box_Control_ItemValue))
                                                                                    {
                                                                                        selectedValue = iTable.Box_Control_ItemValue.Split(',');
                                                                                    }
                                                                                    int i = 0;
                                                                                    foreach (JObject kRowItem in DataItem)
                                                                                    {
                                                                                        if (!String.IsNullOrEmpty(iTable.Column_Item))
                                                                                            iTable.Column_Item += ",";
                                                                                        iTable.Column_Item += $"{kRowItem["item"]}";

                                                                                        if (selectedValue != null && selectedValue[i].ToUpper() == "Y")
                                                                                        {
                                                                                            if (!String.IsNullOrEmpty(iTable.Box_Control_Value))
                                                                                                iTable.Box_Control_Value += ",";
                                                                                            iTable.Box_Control_Value += $"{kRowItem["item"]}";
                                                                                        }
                                                                                        i++;
                                                                                    }
                                                                                }
                                                                                iBoxLayout.Box_Control_Table.Add(iTable);
                                                                                rowCount++;
                                                                            }
                                                                        }
                                                                        iColumnTable.Box_ID = iBoxLayout.Box_ID;
                                                                        iBoxLayout.Box_Control_Column.Add(iColumnTable);
                                                                        columnCount++;
                                                                    }
                                                                    //}
                                                                }

                                                                #endregion
                                                            }

                                                        }

                                                    }
                                                }
                                            }

                                        }
                                    }

                                    JObject jData = (JObject)jLayoutArray[0]["data"];
                                    if (jData != null)
                                    {
                                        iBoxLayout.Box_Control_Value = (String)jData["value"];
                                        JArray DataItem = (JArray)jData["item"];
                                        if (DataItem != null)
                                        {
                                            string[] selectedValue = (string[])null;
                                            if (!string.IsNullOrEmpty(iBoxLayout.Box_Control_Item))
                                            {
                                                selectedValue = iBoxLayout.Box_Control_Item.Split(',');
                                            }
                                            int i = 0;
                                            foreach (JValue kItem in DataItem)
                                            {
                                                if (!String.IsNullOrEmpty(iBoxLayout.Box_Control_ItemValue))
                                                    iBoxLayout.Box_Control_ItemValue += ",";
                                                iBoxLayout.Box_Control_ItemValue += $"{kItem}";

                                                if (selectedValue != null && $"{kItem}".ToUpper() == "Y")
                                                {
                                                    if (!String.IsNullOrEmpty(iBoxLayout.Box_Control_Value))
                                                        iBoxLayout.Box_Control_Value += ",";
                                                    iBoxLayout.Box_Control_Value += selectedValue[i];
                                                }
                                                i++;
                                            }
                                        }
                                    }

                                    listBoxLayout.Add(iBoxLayout);
                                }

                                if (jLayoutArray.Count == 2)
                                {
                                    iBoxLayout = new CustomJsonAdvanceForm.BoxLayout_RefDoc();
                                    iBoxLayout.Box_Column = "1";

                                    JObject jTemplate = (JObject)jLayoutArray[1]["template"];
                                    if (jTemplate != null)
                                    {
                                        iBoxLayout.Box_ID = iRunning.ToString(); iRunning++;
                                        iBoxLayout.Box_Column = "1";
                                        iBoxLayout.Box_ControlType = CustomJsonAdvanceForm.GetControlTypeByJSONKey((String)jTemplate["type"]);
                                        //iBoxLayout.Box_Control_ControlTypeCss = GetCssIconControlType(iBoxLayout.Box_ControlType.ToString());
                                        iBoxLayout.Box_Control_ControlTypeText = GetTextControlType(iBoxLayout.Box_ControlType.ToString());

                                        iBoxLayout.Box_Control_Label = (String)jTemplate["label"];
                                        iBoxLayout.Box_Control_AltLabel = (String)jTemplate["alter"];
                                        iBoxLayout.Box_Control_IsText = (String)jTemplate["istext"];
                                        iBoxLayout.Box_Control_TextValue = (String)jTemplate["textvalue"];
                                        if (jTemplate["description"] != null)
                                            iBoxLayout.Box_Control_Description = (String)jTemplate["description"];

                                        if (jTemplate["formula"] != null)
                                            iBoxLayout.Box_Control_Formula = (String)jTemplate["formula"];

                                        if (jTemplate["attribute"] != null)
                                        {
                                            if (!String.IsNullOrEmpty(jTemplate["attribute"].ToString()))
                                            {
                                                JObject jAttribute = (JObject)jTemplate["attribute"];
                                                if (jAttribute != null)
                                                {
                                                    if (String.IsNullOrEmpty(iBoxLayout.Box_Control_Description) && jAttribute["description"] != null)
                                                        iBoxLayout.Box_Control_Description = (String)jAttribute["description"];
                                                    iBoxLayout.Box_Control_DefaultValue = (String)jAttribute["default"];
                                                    iBoxLayout.Box_Control_MaxLength = (String)jAttribute["length"];
                                                    iBoxLayout.Box_Control_Required = (String)jAttribute["require"];

                                                    iBoxLayout.Box_Control_Min = (String)jAttribute["min"];
                                                    iBoxLayout.Box_Control_Max = (String)jAttribute["max"];
                                                    iBoxLayout.Box_Control_Comma = (String)jAttribute["useComma"];

                                                    iBoxLayout.Box_Control_Inline = (String)jAttribute["multipleLine"];
                                                    iBoxLayout.Box_Control_Summary = (String)jAttribute["summary"];

                                                    iBoxLayout.Box_Control_Decimal = (String)jAttribute["decimal"];
                                                    iBoxLayout.Box_Control_Symbol = (String)jAttribute["symbol"];
                                                    iBoxLayout.Box_Control_SymbolPosition = (String)jAttribute["symbolPosition"];
                                                    iBoxLayout.Box_Control_ValueAlign = (String)jAttribute["align"];

                                                    if (jAttribute["items"] != null)
                                                    {
                                                        if (jAttribute["items"].HasValues)
                                                        {
                                                            JArray itemsItem = (JArray)jAttribute["items"];
                                                            if (itemsItem != null)
                                                            {
                                                                foreach (JObject kItem in itemsItem)
                                                                {
                                                                    if (!String.IsNullOrEmpty(iBoxLayout.Box_Control_Item))
                                                                        iBoxLayout.Box_Control_Item += ",";
                                                                    iBoxLayout.Box_Control_Item += $"{kItem["item"]}";
                                                                }
                                                            }
                                                        }
                                                    }
                                                    JArray itemsColumn = (JArray)jAttribute["column"];
                                                    if (itemsColumn != null)
                                                    {
                                                        #region | Column |

                                                        if (itemsColumn.Count > 0)
                                                        {

                                                            JObject jDataColumn = (JObject)jLayoutArray[0]["data"];
                                                            JArray tableRow = (JArray)jDataColumn["row"];
                                                            iBoxLayout.Box_Control_Column = new List<CustomJsonAdvanceForm.ColumnTable_RefDoc>();
                                                            iBoxLayout.Box_Control_Table = new List<CustomJsonAdvanceForm.ColumnTable_RefDoc>();
                                                            int columnCount = 0;
                                                            foreach (JObject kItem in itemsColumn)
                                                            {

                                                                CustomJsonAdvanceForm.ColumnTable_RefDoc iColumnTable = new CustomJsonAdvanceForm.ColumnTable_RefDoc();

                                                                iColumnTable.Column_Label = (String)kItem["label"];
                                                                iColumnTable.Column_AltLabel = (String)kItem["alter"];
                                                                iColumnTable.Column_ControlType = CustomJsonAdvanceForm.GetControlTypeByJSONKey((String)kItem["control"]["template"]["type"]);
                                                                //iColumnTable.Column_ControlTypeCss = GetCssIconControlType(iColumnTable.Column_ControlType.ToString());
                                                                iColumnTable.Column_ControlTypeText = GetTextControlType(iColumnTable.Column_ControlType.ToString());
                                                                if (!String.IsNullOrEmpty(kItem["control"]["template"].ToString()))
                                                                {
                                                                    JObject ktemplate = (JObject)kItem["control"]["template"];
                                                                    if (ktemplate != null)
                                                                    {

                                                                        iColumnTable.Column_ControlType = CustomJsonAdvanceForm.GetControlTypeByJSONKey((String)ktemplate["type"]);
                                                                        //iColumnTable.Column_ControlTypeCss = GetCssIconControlType(iColumnTable.Column_ControlType.ToString());
                                                                        iColumnTable.Column_ControlTypeText = GetTextControlType(iColumnTable.Column_ControlType.ToString());
                                                                        if (ktemplate["attribute"] != null)
                                                                            iColumnTable.Column_Description = (String)ktemplate["attribute"];

                                                                        if (!String.IsNullOrEmpty(ktemplate["attribute"].ToString()))
                                                                        {
                                                                            JObject kAttribute = (JObject)ktemplate["attribute"];
                                                                            if (kAttribute != null)
                                                                            {
                                                                                if (String.IsNullOrEmpty(iColumnTable.Column_Description) && kAttribute["description"] != null)
                                                                                    iColumnTable.Column_Description = (String)kAttribute["description"];
                                                                                iColumnTable.Column_DefaultValue = (String)kAttribute["default"];
                                                                                JObject objkDate = (JObject)kAttribute["date"];

                                                                                if (kAttribute["items"] != null)
                                                                                {
                                                                                    if (kAttribute["items"].HasValues)
                                                                                    {
                                                                                        JArray itemsItemInColumn = (JArray)kAttribute["items"];
                                                                                        if (itemsItemInColumn != null)
                                                                                        {
                                                                                            foreach (JObject lItem in itemsItemInColumn)
                                                                                            {
                                                                                                if (!String.IsNullOrEmpty(iColumnTable.Column_Item))
                                                                                                    iColumnTable.Column_Item += ",";
                                                                                                iColumnTable.Column_Item += $"{lItem["item"]}";
                                                                                            }

                                                                                        }
                                                                                    }
                                                                                }

                                                                            }
                                                                        }

                                                                    }
                                                                }

                                                                if (tableRow != null)
                                                                {
                                                                    iBoxLayout.Box_Control_TableRowCount = tableRow.Count;
                                                                    int rowCount = 0;
                                                                    foreach (JArray dataRow in tableRow)
                                                                    {
                                                                        CustomJsonAdvanceForm.ColumnTable_RefDoc iTable = new CustomJsonAdvanceForm.ColumnTable_RefDoc();
                                                                        iTable.Box_Control_RowIndex = rowCount;
                                                                        iTable.Table_Name = iBoxLayout.Box_Control_Label;
                                                                        iTable.Column_Label = (String)kItem["label"];
                                                                        iTable.Column_AltLabel = (String)kItem["alter"];
                                                                        iTable.Box_Control_Value = (String)dataRow[columnCount]["value"];
                                                                        iTable.Column_ControlType = iColumnTable.Column_ControlType;
                                                                        iTable.Column_ControlTypeCss = iColumnTable.Column_ControlTypeCss;
                                                                        iTable.Column_ControlTypeText = iColumnTable.Column_ControlTypeText;
                                                                        JArray DataItem = (JArray)kItem["control"]["template"]["attribute"]["items"];
                                                                        JArray DataItemValue = (JArray)dataRow[columnCount]["item"];
                                                                        if (DataItemValue != null)
                                                                        {
                                                                            foreach (JValue kRowItem in DataItemValue)
                                                                            {
                                                                                if (!String.IsNullOrEmpty(iTable.Box_Control_ItemValue))
                                                                                    iTable.Box_Control_ItemValue += ",";
                                                                                iTable.Box_Control_ItemValue += $"{kRowItem}";
                                                                            }
                                                                        }
                                                                        if (DataItem != null)
                                                                        {
                                                                            string[] selectedValue = (string[])null;
                                                                            if (!string.IsNullOrEmpty(iTable.Box_Control_ItemValue))
                                                                            {
                                                                                selectedValue = iTable.Box_Control_ItemValue.Split(',');
                                                                            }
                                                                            int i = 0;
                                                                            foreach (JObject kRowItem in DataItem)
                                                                            {
                                                                                if (!String.IsNullOrEmpty(iTable.Column_Item))
                                                                                    iTable.Column_Item += ",";
                                                                                iTable.Column_Item += $"{kRowItem["item"]}";

                                                                                if (selectedValue[i].ToUpper() == "Y")
                                                                                {
                                                                                    if (!String.IsNullOrEmpty(iTable.Box_Control_Value))
                                                                                        iTable.Box_Control_Value += ",";
                                                                                    iTable.Box_Control_Value += $"{kRowItem["item"]}";
                                                                                }
                                                                                i++;
                                                                            }
                                                                        }
                                                                        iBoxLayout.Box_Control_Table.Add(iTable);
                                                                        rowCount++;
                                                                    }
                                                                }
                                                                iColumnTable.Box_ID = iBoxLayout.Box_ID;
                                                                iBoxLayout.Box_Control_Column.Add(iColumnTable);
                                                                columnCount++;
                                                            }
                                                        }

                                                        #endregion

                                                    }
                                                }
                                            }
                                        }
                                    }

                                    JObject jData = (JObject)jLayoutArray[1]["data"];
                                    if (jData != null)
                                    {
                                        iBoxLayout.Box_Control_Value = (String)jData["value"];
                                        JArray DataItem = (JArray)jData["item"];
                                        if (DataItem != null)
                                        {
                                            string[] selectedValue = (string[])null;
                                            if (!string.IsNullOrEmpty(iBoxLayout.Box_Control_Item))
                                            {
                                                selectedValue = iBoxLayout.Box_Control_Item.Split(',');
                                            }
                                            int i = 0;
                                            foreach (JValue kItem in DataItem)
                                            {
                                                if (!String.IsNullOrEmpty(iBoxLayout.Box_Control_ItemValue))
                                                    iBoxLayout.Box_Control_ItemValue += ",";
                                                iBoxLayout.Box_Control_ItemValue += $"{kItem}";

                                                if (selectedValue != null && $"{kItem}".ToUpper() == "Y")
                                                {
                                                    if (!String.IsNullOrEmpty(iBoxLayout.Box_Control_Value))
                                                        iBoxLayout.Box_Control_Value += ",";
                                                    iBoxLayout.Box_Control_Value += selectedValue[i];
                                                }
                                                i++;
                                            }
                                        }
                                    }

                                    listBoxLayout.Add(iBoxLayout);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
            return listBoxLayout;

        }
        static void GetTransaction()
        {

            DataClasses1DataContext db = new DataClasses1DataContext(dbConnectionString);

            if (db.Connection.State == ConnectionState.Open)
            {
                db.Connection.Close();
                db.Connection.Open();
            }
            else
            {
                try
                {
                    db.Connection.Open();
                    log.Info(string.Format("Get Email Template"));
                    List<MSTEmailTemplate> objListEmail = db.MSTEmailTemplates.Where(x => x.FormState == TemplateEmailState && x.IsActive == true).ToList();
                    if (objListEmail != null)
                    {
                        log.Info(string.Format("Get Template Data"));
                        List<MSTTemplate> objListTargetTemplate = db.MSTTemplates.Where(x => x.DocumentCode == TemplateDocumentCode).ToList();
                        if (objListTargetTemplate != null)
                        {
                            if (objListTargetTemplate.Count > 0)
                            {
                                var ListTemplateid = objListTargetTemplate.Select(a => a.TemplateId).ToArray();
                                log.Info(string.Format("Get Transaction Data"));
                                List<TRNMemo> objListMemo = new List<TRNMemo>();
                                if(MemoIDSpec != 0)
                                {
                                    objListMemo = db.TRNMemos.Where(x => x.MemoId == MemoIDSpec).ToList();
                                }
                                else
                                {
                                    objListMemo = db.TRNMemos.Where(x => x.StatusName == "Completed" && x.ModifiedDate >= DateTime.Now.AddMinutes(iIntervalTime)).ToList();
                                }
                             
                                if (objListMemo != null)
                                {
                                    List<TRNMemo> objTargetTransaction = objListMemo.Where(o => ListTemplateid.Contains(o.TemplateId ?? 0)).ToList();
                                    if (objTargetTransaction != null)
                                    {
                                        if (objTargetTransaction.Count > 0)
                                        {
                                            log.Info(string.Format("Begin Transaction Data : {0}", objTargetTransaction.Count));
                                            List<MSTEmployee> objActiveEmp = db.MSTEmployees.Where(x => x.IsActive == true).ToList();
                                            if (objActiveEmp != null)
                                            {

                                                foreach (TRNMemo ItemMemo in objTargetTransaction)
                                                {

                                                    string sSubjectEmail = objListEmail[0].EmailSubject;
                                                    string sContentEmail = objListEmail[0].EmailBody;
                                                    string sBU = "";
                                                    string sCompany = "";
                                                    string sDepartment = "";
                                                    string sLevel = "";
                                                    string sYear = "";
                                                    string sRound = "";
                                                    string sStandard = "";
                                                    string sReviseRound = "";
                                                    string sQuarter = "";
                                                    string sAuditType = "BSC";

                                                    //ItemRow Variable
                                                    string sStartDate = "";
                                                    string sEndDate = "";
                                                    string sStartTime = "";
                                                    string sEndTime = "";
                                                    string sAuditorLead = "";
                                                    string sAuditorTeam = "";
                                                    string sCoAudit = "";
                                                    string sAuditorLeadEmail = "";
                                                    string sAuditorTeamEmail = "";
                                                    string sItemDepartment = "";
                                                    string sAuditorCreatePlan = "";
                                                 
                                                    string sISONo = "";

                                                    DateTime dStartDate = DateTime.MinValue;
                                                    DateTime dEndDate = DateTime.MinValue;

                                                    List<CustomJsonAdvanceForm.BoxLayout_RefDoc> objJson = convertAdvanceFormToList(ItemMemo.MAdvancveForm).ToList();
                                                    log.Info(string.Format("Begin Transaction : PK - {0}", ItemMemo.MemoId));
                                                    foreach (CustomJsonAdvanceForm.BoxLayout_RefDoc item in objJson)
                                                    {
                                                        switch (item.Box_Control_Label)
                                                        {
                                                            case "Business Unit":
                                                                sBU = item.Box_Control_Value ?? string.Empty;
                                                                break;
                                                            case "ฝ่าย / กลุ่มงาน / แผนก ผู้จัดทำแผน":
                                                                sCompany = item.Box_Control_Value ?? string.Empty;
                                                                break;
                                                            case "ปี":
                                                                sYear = item.Box_Control_Value ?? string.Empty;
                                                                break;
                                                            case "ผู้จัดทำแผนงาน":
                                                                sAuditorCreatePlan = item.Box_Control_Value ?? string.Empty;
                                                                break;
                                                       
                                                            default:
                                                                break;
                                                        }

                                                        if (item.Box_ControlType == CustomJsonAdvanceForm.ControlTypeEnum.Table)
                                                        {
                                                            if (item.Box_Control_Label == "แผนการตรวจ")
                                                            {
                                                                var Table = item.Box_Control_Table.OrderBy(a => a.Box_Control_RowIndex).ToList();
                                                                var Index = Table.GroupBy(a => a.Box_Control_RowIndex).Select(s => s.Key).ToList();

                                                                foreach (var itemTableRow in Index)
                                                                {
                                                                    //���� Row � Table
                                                                    foreach (var itemTable in Table.FindAll(a => a.Box_Control_RowIndex == itemTableRow).ToList())
                                                                    {
                                                                        //���� Column
                                                                        if (itemTable.Column_Label == "วันที่เริ่มต้น")
                                                                        {
                                                                            sStartDate = itemTable.Box_Control_Value;
                                                                        }
                                                                        else if (itemTable.Column_Label == "วันที่สิ้นสุด")
                                                                        {
                                                                            sEndDate = itemTable.Box_Control_Value;
                                                                        }
                                                                        else if (itemTable.Column_Label == "หัวหน้าผู้ตรวจ")
                                                                        {
                                                                            sAuditorLead = itemTable.Box_Control_Value;
                                                                        }
                                                                        else if (itemTable.Column_Label == "ผู้รับผิดชอบ")
                                                                        {
                                                                            sAuditorTeam = itemTable.Box_Control_Value;
                                                                        }
                                                                        else if (itemTable.Column_Label == "มาตรฐานที่เกี่ยวข้อง")
                                                                        {
                                                                            sStandard = itemTable.Box_Control_Value;
                                                                        }
                                                                        else if (itemTable.Column_Label == "รหัสพื้นที ISO ที่ถูกตรวจ")
                                                                        {
                                                                            sISONo = itemTable.Box_Control_Value;
                                                                        }
                                                                    }
                                                                    log.Info("sAuditorLead :" + sAuditorLead);
                                                                    List<MSTEmployee> objEmpAuditLead = objActiveEmp.Where(x => x.NameEn.Replace(" ", "").Contains(sAuditorLead.Replace(" ", ""))).ToList();
                                                                    if (objEmpAuditLead != null)
                                                                    {
                                                                        if (objEmpAuditLead.Count > 0)
                                                                        {
                                                                            sAuditorLeadEmail = objEmpAuditLead[0].Email;

                                                                            List<MSTEmployee> objEmpAuditTeam = objActiveEmp.Where(x => x.NameEn.Replace(" ", "").Contains(sAuditorTeam.Replace(" ", ""))).ToList();
                                                                            if (objEmpAuditTeam != null)
                                                                            {
                                                                                sAuditorTeamEmail = objEmpAuditTeam[0].Email;
                                                                            }
                                                                            List<MSTEmployee> objEmpAuditCreate = objActiveEmp.Where(x => x.NameEn.Replace(" ", "").Contains(sAuditorCreatePlan.Replace(" ", ""))).ToList();
                                                                            if (objEmpAuditTeam != null)
                                                                            {
                                                                                sAuditorTeamEmail = ((sAuditorTeamEmail != string.Empty)? ";" : "") + objEmpAuditCreate[0].Email;
                                                                            }
                                                                            string sSubjectEmailFinal = "";
                                                                            string sContentEmailFinal = "";
                                                                            sSubjectEmailFinal = sSubjectEmail.Replace("[-Round-]", sRound)
                                                                                 .Replace("[-Year-]", sYear)
                                                                                .Replace("[-Standard-]", sStandard)
                                                                                .Replace("[-StartDate-]", sStartDate)
                                                                                .Replace("[-EndDate-]", sEndDate)
                                                                                .Replace("[-ISONo-]", sISONo);

                                                                            sContentEmailFinal = sContentEmail.Replace("[-Round-]", sRound)
                                                                                 .Replace("[-Year-]", sYear)
                                                                                .Replace("[-Standard-]", sStandard)
                                                                                .Replace("[-StartDate-]", sStartDate)
                                                                                .Replace("[-EndDate-]", sEndDate)
                                                                                .Replace("[-ISONo-]", sISONo);

                                                                            try
                                                                            {
                                                                                dStartDate = DateTime.ParseExact(sStartDate + " 09:00", "dd MMM yyyy HH:mm", new System.Globalization.CultureInfo("en-GB"));
                                                                                if (sEndDate == "") sEndDate = sStartDate;
                                                                                dEndDate = DateTime.ParseExact(sEndDate + " 17:00", "dd MMM yyyy HH:mm", new System.Globalization.CultureInfo("en-GB"));

                                                                                string sResultEmail = MailUtilities.CreateCalendarEntry(ItemMemo.MemoId.ToString(), dStartDate, dEndDate, sSubjectEmailFinal, sContentEmailFinal, "-", sAuditorLeadEmail, sAuditorTeamEmail,TestEmail);
                                                                                //string sResultEmail = MailUtilities.CreateCalendarEntry(ItemMemo.MemoId.ToString(), dStartDate, dEndDate, sSubjectEmail, sSubjectCalendar, sContentEmail, sContentCalendar, "-", "nichapa@techconsbiz.com", "siroat@techconsbiz.com;thitipon@techconsbiz.com");
                                                                                log.Info(string.Format("Result Send Email to {1} Transaction Result : {0}", sResultEmail, sAuditorLeadEmail));
                                                                                log.Info(string.Format("Result Send Email cc {1} Transaction Result : {0}", sResultEmail, sAuditorTeamEmail));
                                                                            }
                                                                            catch (Exception ex)
                                                                            {
                                                                                log.Info(string.Format("Error Send : {0} \r\n {1}", sAuditorLeadEmail, ex.Message.ToString()));
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                log.Info(string.Format("Not Found Active Employee"));
                                            }
                                        }
                                        else
                                        {
                                            log.Info(string.Format("Not Found Memo (zero row)"));
                                        }
                                    }
                                    else
                                    {
                                        log.Info(string.Format("Not Found Memo"));
                                    }
                                }
                                else
                                {
                                    log.Info(string.Format("Not Found Target Memo "));
                                }
                            }
                        }
                        else
                        {
                            log.Info(string.Format("Not Found Template Code : {0}", TemplateDocumentCode));
                        }
                    }
                    else
                    {
                        log.Info(string.Format("Not Found Email Template {0} ", TemplateEmailState));
                    }

                }
                catch (Exception ex)
                {
                    log.Error(ex.Message, ex);

                    Console.WriteLine(":ERROR");
                    Console.WriteLine("exit 1");

                    log.Info(":ERROR");
                    log.Info("exit 1");
                }
            }

        }
        public static String GetTextControlType(String sType)
        {
            sType = System.Text.RegularExpressions.Regex.Replace(sType, "[A-Z]", " $0");
            return sType.Trim();
        }
    }


}
