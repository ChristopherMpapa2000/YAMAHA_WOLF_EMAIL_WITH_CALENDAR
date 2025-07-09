using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;

namespace WOLF_EMAIL_WITH_CALENDAR
{
    public class MailUtilities
    {
        static string sPathICS = ConfigurationManager.AppSettings["PathFileiCS"];
        static MailMessage mail;
        static SmtpClient SmtpServer;
        static int port;
        static string from;
        static string sMailAuthen;
        static string sPasswordAuthen;
        public static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static void sendMailCampaign(string to, string subject, string body)
        {
            try
            {
                if (readConfig())
                {

                    mail.IsBodyHtml = true;
                    mail.From = new MailAddress(from, "WOLF Notification");
                    if (to.Contains(';'))
                    {
                        foreach (string mto in to.Split(';'))
                        {
                            if (!string.IsNullOrWhiteSpace(mto))
                            {
                                mail.To.Add(mto);
                            }
                        }
                    }
                    else
                    { mail.To.Add(to); }
                    mail.Subject = subject;
                    mail.Body = body;
                    SmtpServer.Port = port;
                    SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    SmtpServer.EnableSsl = EnableSsl; //.EnableSsl;
                    try
                    {
                        SmtpServer.Send(mail);
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Email :" + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("sendMailAssignlead", ex);
                throw new Exception("Email From : " + from + " Email To: " + to + "Subject :" + subject + "Body :" + body, ex);
                //throw ex ;
            }
        }

        public static string CreateCalendarEntry(string sPK, DateTime? start, DateTime? end, string title, string description, string location, string sRecipient, string sCC ,string  TestEmail)
        {
            try
            {
                Calendar iCal = new Calendar();

                iCal.Method = "PUBLISH";
                // Create the event, and add it to the iCalendar

                CalendarEvent evt = iCal.Create<CalendarEvent>();

                // Set information about the event
                evt.Start = new CalDateTime(start.Value);
                evt.End = new CalDateTime(end.Value); // This also sets the duration  

                evt.Description = description;
                //evt.Location = location;
                evt.Summary = title;
                // Create a reminder 24h before the event
                Alarm reminder = new Alarm();
                reminder.Action = AlarmAction.Display;
                reminder.Trigger = new Trigger(new TimeSpan(-24, 0, 0));
                evt.Alarms.Add(reminder);

                var property = new CalendarProperty("X-ALT-DESC;FMTTYPE=text/html", description);
                iCal.Calendar.AddProperty(property);

                CalendarSerializer serializer = new CalendarSerializer(new SerializationContext());
                string sRR = serializer.SerializeToString(iCal);
                sRR = sRR.Replace("X-ALT-DESC:FMTTYPE=text/html", "X-ALT-DESC;FMTTYPE=text/html");

                //System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType("text/calendar");
                //ct.Parameters.Add("method", "REQUEST");
                //AlternateView avCal = AlternateView.CreateAlternateViewFromString(serializer.SerializeToString(iCal), ct);
                string sTime = DateTime.Now.ToString("ddHHss");
                string sPathFileICS = string.Format("{0}Calendar{1}.ics", sPathICS, sPK);
                File.WriteAllText(sPathFileICS, sRR);
                //System.Net.Mime.ContentType contype = new System.Net.Mime.ContentType("text/calendar");
                System.Net.Mail.Attachment attach = new System.Net.Mail.Attachment(sPathFileICS);
                //attach.ContentDisposition.FileName = "myFile.ics";


                //Response.Write(str);
                // sc.ServicePoint.MaxIdleTime = 2;
                if (readConfig())
                {


                    MailMessage msg = new MailMessage();
                    System.Net.Mail.AlternateView htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(description, null,
                       System.Net.Mime.MediaTypeNames.Text.Html);
                    msg.AlternateViews.Add(htmlView);

                    msg.Attachments.Add(attach);
                    ///msg.Body = description;
                    msg.IsBodyHtml = true;
                    msg.From = new MailAddress(from);
                    if (sRecipient.Contains(';'))
                    {
                        foreach (string mto in sRecipient.Split(';'))
                        {
                            if (!string.IsNullOrWhiteSpace(mto))
                            {
                                msg.To.Add(mto);
                            }
                        }
                    }
                    else
                    {
                        msg.To.Add(sRecipient);
                    }
                    if (!string.IsNullOrEmpty(sCC))
                    {

                        if (sCC.Contains(';'))
                        {
                            foreach (string mCCo in sCC.Split(';'))
                            {
                                if (!string.IsNullOrWhiteSpace(mCCo))
                                {
                                    msg.CC.Add(mCCo);
                                }
                            }
                        }
                        else { msg.To.Add(sCC); }
                    }
                    if (TestEmail != string.Empty)
                    {
                        Console.WriteLine(":TestEmail : "+ TestEmail);
                        msg.To.Clear();
                        msg.To.Add(TestEmail);
                        msg.CC.Clear();
                        msg.CC.Add(TestEmail);
                    }
                    msg.Subject = title;
                    //msg.AlternateViews.Add(avCal);

                    logger.Info("Start sendmail : " + msg.To);
                    SmtpServer.Port = port;
                    SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    if (!string.IsNullOrEmpty(sMailAuthen))
                    {
                        SmtpServer.UseDefaultCredentials = false;
                        SmtpServer.Credentials = new System.Net.NetworkCredential(sMailAuthen, sPasswordAuthen);
                    }
                    else
                    {
                        SmtpServer.UseDefaultCredentials = true;
                    }
                    logger.Info("End sendmail" );
                    SmtpServer.EnableSsl = EnableSsl;

                    SmtpServer.Send(msg);
                }

                return "";
            }
            catch (Exception ex)
            {
                logger.Info("Error sendmail : "+ex.Message);
                return ex.ToString();
            }
        }
        private static bool readConfig()
        {
            bool status = false;
            try
            {
                mail = new MailMessage();
                SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SMTPServer"]); //Call Config - > SMTPServer
                port = int.Parse(ConfigurationManager.AppSettings["SMTPPort"]); //Call Config - > SMTPPort Default 25
                from = ConfigurationManager.AppSettings["EmailFrom"]; //Call Config - > SMTPFrom default T-CRM@thanachart.co.th
                sMailAuthen = ConfigurationManager.AppSettings["sEmailAuthen"];
                sPasswordAuthen = ConfigurationManager.AppSettings["sPassAuthen"];
                status = true;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return status;
        }
        private static bool EnableSsl
        {
            get
            {
                string EnableSsl_ = ConfigurationManager.AppSettings["EnableSsl"];
                if (!string.IsNullOrWhiteSpace(EnableSsl_))
                {
                    if (EnableSsl_ == "true")
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                return false;
            }
        }

    }
}
