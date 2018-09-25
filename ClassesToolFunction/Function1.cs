using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using ClassesToolGraph;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;

// ClassesToolFunction - Azure Function
// Demo for Ignite 2018: Modernize your apps with Graph
// Toni Pohl, atwork.at, @atwork
// Modern app: Send an email to a resource mailbox with Graph and create an appointment out of it.
//
// Subject: must contain "ID:"
//      ID45: Modern Workplace Conference
// Body: must contain the startdate and enddate as 1st and 2nd line as here:
//      20180902 09:00
//      20180902 15:00
//
namespace ClassesToolFunction
{
    public static class Function1
    {
        private static GraphServiceClient graphService;

        [FunctionName("ClassesToolFunction")]
        // Run all n minutes or HTTP (Webhook) or queue triggered...
        //https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer#cron-expressions
        // "0 */5 * * * *"  - all 5 minutes
        // "0 0 * * * *"    - once at the top of every hour
        // "0 0 */2 * * *"  - once every two hours
        // "0 30 9 * * *"   - at 9:30 AM every day
        // "0 30 2 * * 1"   - each first weekday at 2:30AM
        // "0 30 2 * * 1-5" - each weekday at 2:30
        public static async Task Run([TimerTrigger("0 */5 * * * *")]TimerInfo myTimer, TraceWriter log)
        {
            log.Info($"ClassesTool-Graph {DateTime.Now}");
            // https://docs.microsoft.com/en-us/azure/azure-functions/functions-dotnet-class-library#functions-class-library-project
            var mailboxes = Enviroments.GetEnvironmentVariable("Inbox").Split(new[] { ';' }).ToList();

            // Connect to Graph
            graphService = GraphClient.Create();

            // Do something
            log.Info("Loop through mailboxes");
            foreach (var box in mailboxes)
            {
                await SendEmail(box, log);
                //await GetEmails(box, log);
                //await GetCalendar(box, log);
            }

            log.Info("-- Done. --");
        }
        
        // methods start
        private static async Task SendEmail(string mailbox, TraceWriter log)
        {
            // Generate n demo emails
            for (int i = 0; i < 3; i++)
            {
                Random rnd = new Random();
                int randomId = rnd.Next(10, 100); // Generate a random ID 10..99 for the new appointment
                string sender = Enviroments.GetEnvironmentVariable("SenderEMail");

                Message msg = new Message();

                var emailAddress = new EmailAddress { Address = mailbox };
                Recipient recipient = new Recipient() { EmailAddress = emailAddress };
                msg.ToRecipients = new List<Recipient>() { recipient };
                msg.From = new Recipient() { EmailAddress = new EmailAddress() { Address = sender } };
                msg.Subject = "ID:" + randomId.ToString() + " Graph Function Workshop";
                msg.Body = new ItemBody()
                {
                    ContentType = BodyType.Text,
                    Content = (DateTime.Now.AddHours(randomId).ToString("yyyyMMdd HH:mm") + "\n\r" +
                               DateTime.Now.AddHours(randomId + 1).ToString("yyyyMMdd HH:mm") + "\n\r" +
                               "This is a generated seminar entry with ID:" + randomId.ToString() +
                               " sent by the ClassesTool-Graph.\n\r")
                };

                await graphService.SendMailAsync(msg, sender);
                log.Info($"Sent: {msg.Subject}");
            }
        }


        private static async Task GetEmails(string mailbox, TraceWriter log)
        {
            var messages = await graphService.GetMailAsync(mailbox);
            foreach (var msg in messages)
            {
                log.Info($"E-Mail: {msg.SentDateTime} {msg.Subject}");
                // valid appointment? does it contain "ID:" ?
                if (msg.Subject.ToLower().Contains("id:"))
                {
                    if (msg.Subject.ToLower().StartsWith("delete"))
                    {
                        //DeleteCalendarEntry(mailbox, msg.Subject);
                    }
                    else
                    {
                        await AddAppointment(mailbox, msg,log);
                    }
                    // remove that email
                    await graphService.RemoveMailAsync(msg, mailbox);
                }
            }
        }

        private static async Task AddAppointment(string mailbox, Message msg, TraceWriter log)
        {
            string body = msg.Body.Content.ToString().Replace("\n", ""); // Remove linefeed
            string[] lines = body.Split('\r'); // split by line

            if (lines.Count() > 1)
            {
                try
                {
                    string startdate = GetDate(lines[0].Trim(), log);
                    string enddate = GetDate(lines[1].Trim(), log);
                    log.Info($"AddAppointment: {startdate}, {enddate}, {msg.Subject}");

                    Event myevent = new Event();

                    myevent.Start = new DateTimeTimeZone()
                    {
                        DateTime = startdate,
                        TimeZone = "Eastern Standard Time"
                    };
                    myevent.End = new DateTimeTimeZone()
                    {
                        DateTime = enddate,
                        TimeZone = "Eastern Standard Time"
                    };
                    myevent.Subject = msg.Subject;
                    myevent.Body = new ItemBody()
                    {
                        ContentType = BodyType.Html,
                        Content = msg.Body.Content.ToString()
                    };

                    await graphService.AddCalendarEntryAsync(myevent, mailbox);
                }
                catch (Exception ex)
                {
                    log.Info("Error AddAppointment: " + ex.Message);
                }
            }
        }

        private static string GetDate(string date, TraceWriter log)
        {
            // Just a helper to convert a string to a DateTime
            var appointmentDate = "";
            try
            {
                date = date.Replace("T", "");
                appointmentDate = DateTime.ParseExact(date, "yyyyMMdd HH:mm", CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal).ToString(CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                log.Info("Error GetDate: " + ex.Message);
            }
            return appointmentDate;
        }

        private static async Task GetCalendar(string mailbox, TraceWriter log)
        {
            var appointments = await graphService.GetCalendarAsync(mailbox);
            foreach (var app in appointments)
            {
                log.Info(app.Subject);
            }
        }

        // end of methods
    }
}
