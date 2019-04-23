using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Net;
using System.Text;


namespace SendEmail
{
    public static class HEmail
    {

        //public static string EmailFrom { get; set; }
        //public static string EmailFromPassword { get; set; }
        //public static string EmailTo { get; set; }
        //public static string EmailHost { get; set; }
        //public static int EmailPort { get; set; }
        //public static bool EmailEnableSsl { get; set; }
        //public static int EmailTimeOut { get; set; }
        //public static bool EmailIsHTML { get; set; }
        //public static string EmailSubject { get; set; }
        //public static string EmailBody { get; set; }
        public static string EmailTo;
        public static string body = "<p><b>Berkeley INNOVATION INDEX Beta Release V1.6</b> <br><br>Assalamu'alaikum Wr Wb.<br>Morning guys, thank you for joining us as our respondent. There is your Innovation Mindset Score, based on your answer in Google Form.<br><br>This is not a fixed level, anyone can grow their innovation mindset. Your level has been estimated using an analysis based on the Berkeley Method for Entrepreneurship & Innovation and fundamental testing methods in social psychology.<br><br> <b>INNOVATION MINDSET </b>:<br>Your personal Innovation Mindset Level is currently <b> @IM </b> out of 10 <br><br>The following factors are components of your innovation mindset: <br> - TRUST level: @QT of 10. This is your ability to trust others. <br> - RESILIENCE level: @QF of 10. This is your ability to overcome failure. <br> - DIVERSITY level: @QD of 10. This is your ability to overcome social barriers. <br> - MENTAL STRENGTH level: @QB of 10. This is a measure of your confidence and belief that you can succeed. <br> - COLLABORATION level: @QC of 10. This is your ability to work with everyone including competitors when needed. <br> - RESOURCE AWARENESS level: @QP of 10. This is your ability to balance your resources across multiple objectives. <br><br>These scores are normalized.  <b>The average INNOVATION MINDSET score of the general population in STIMIK ESQ is 6.38 of 10</b>. <br><br>Learn more, come and join us on Thursday, August 10th 2017 <link poster> <br><br>Wassalamu'alaikum Wr Wb. <br>Regards,<br>Cak Khoiron and team. <br>http://berkeleyinnovationindex.org</p>";

        public static void SendEmail()
        {
            try
            {

                var fromAddress = new MailAddress("youremail@gmail.com");
                var toAddress = new MailAddress(EmailTo);
                string fromPassword = "*********";
                string subject = "Your Innovation Mindset Score";

                Console.WriteLine("Start sending to " + toAddress);

                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };
                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = Program.bodySend
                })
                {
                    MailAddress addressBCC = new MailAddress("m.khoiron@students.esqbs.ac.id");
                    message.Bcc.Add(addressBCC);
                    message.IsBodyHtml = true;


                    smtp.Send(message);
                }

                Console.WriteLine("Message is Sent to " + toAddress);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error" + ex);
            }

        }

    }
}
