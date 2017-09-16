using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Net.Mail;
using System.Windows;

namespace Pauls_Contest_Mailer
{
    public class SpreadsheetAnalysis
    {
        public bool mailNotFound = false;

        private Dictionary<string, SmtpClient> mailClients = null;
        private List<List<MailMessage>> contests = new List<List<MailMessage>>();

        public void ProcessWorkbook(string filename)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook wkb = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                wkb = excelApp.Workbooks.Open(filename);

                foreach (Worksheet ws in wkb.Sheets)
                {
                    if (ws.Name == "Emails")
                    {
                        Console.WriteLine("We found the emails tab");
                        mailClients = ParseEmails(ws);
                    }
                    else if (ws.Name.Substring(0, 11) == "Gewinnspiel")
                    {
                        Console.WriteLine("This is a contest tab");
                        List<MailMessage> contest = ParseContest(ws);
                        if (mailNotFound)
                        {
                            return;
                        }
                        contests.Add(contest);
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private Dictionary<string, SmtpClient> ParseEmails(Worksheet ws)
        {
            Dictionary<string, SmtpClient> mailSetup = new Dictionary<string, SmtpClient>();

            // Parse WorkSheet
            Range rows = ws.UsedRange;

            object[,] valueArray = (object[,])rows.get_Value(XlRangeValueDataType.xlRangeValueDefault);

            for (int row = 2; row <= ws.UsedRange.Rows.Count; ++row)
            {
                SmtpClient tempClient = new SmtpClient
                {
                    Port = 25,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new System.Net.NetworkCredential($"{valueArray[row, 2]}", $"{valueArray[row, 3]}"),
                    Host = $"{valueArray[row, 4]}"
                };
                try
                {
                    mailSetup.Add($"{valueArray[row, 1]}", tempClient);
                }
                catch (ArgumentException)
                {
                    Console.WriteLine("Duplicate credentials for Email " + $"{valueArray[row, 1]}");
                }
            }

            return mailSetup;
        }

        private List<MailMessage> ParseContest(Worksheet ws)
        {
            List<MailMessage> contestMails = new List<MailMessage>();
            Range rows = ws.UsedRange;

            object[,] valueArray = (object[,])rows.get_Value(XlRangeValueDataType.xlRangeValueDefault);

            for (int row = 2; row <= ws.UsedRange.Rows.Count; ++row)
            {
                if (valueArray[row, 1] != null)
                {
                    MailMessage mail = new MailMessage($"{valueArray[row, 1]}", $"{valueArray[row, 2]}")
                    {
                        Subject = $"{valueArray[row, 3]}",
                        Body = $"{valueArray[row, 4]}"
                    };
                    

                    // If Sender Email doesn't have config, do not add it and complain!
                    if (mailClients.ContainsKey($"{valueArray[row, 1]}"))
                    {
                        contestMails.Add(mail);
                    }
                    else
                    {
                        MessageBoxResult result = MessageBox.Show("Für die angegebene Absenderadresse existiert keine passende Konfiguration im Emails Tab: " + $"{valueArray[row, 1]}" + Environment.NewLine + "Möchtest du trotzdem weiter machen? (Überspringt diesen Eintrag)", "Fehlende Email Konfiguration", MessageBoxButton.YesNo, MessageBoxImage.Question);

                        if (result == MessageBoxResult.Yes)
                        {
                            continue;
                        }
                        else
                        {
                            mailNotFound = true;
                            return null;
                        }
                    }
                }
            }

            return contestMails;
        }

        public int CountMails()
        {
            int total = 0;
            foreach (List<MailMessage> contest in contests)
            {
                total += contest.Count;
            }

            return total;            
        }

        public void SendMails(System.Windows.Controls.DataGrid grid, double interval)
        {
            List<object> gridSource = null;
            int idx = 1;

            grid.ItemsSource = gridSource;

            int intervaMillilSeconds = Convert.ToInt32(interval * 60 * 1000);

            foreach (List<MailMessage> contestMails in contests)
            {
                foreach (MailMessage contestMail in contestMails)
                {
                    try
                    {
                        System.Threading.Thread.Sleep(intervaMillilSeconds);
                        mailClients[contestMail.From.ToString()].Send(contestMail);
                        gridSource.Add(new { Index = idx++, Absender = contestMail.From.ToString(), Empfaenger = contestMail.To.ToString(), Betreff = contestMail.Subject.ToString() });
                    }
                    catch (SmtpException smtpEx)
                    {
                        MessageBox.Show(smtpEx.Message, "Fehler beim Mailversand", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }
    }
}
