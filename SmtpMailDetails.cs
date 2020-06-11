using System;
using System.Collections.Generic;

namespace word_mailmerge
{
    public class SmtpMailDetails
    {
        public SmtpMailDetails()
        {
            mime_attachment_list = new List<string>();
            mime_mail_to_list = new List<Tuple<string, string>>();
        }
        public int smtp_mail_batch_id;
        public List<Tuple<string, string>> mime_mail_to_list;
        public List<string> mime_attachment_list;
    }
}
