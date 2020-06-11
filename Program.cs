using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Exports;

namespace word_mailmerge
{
    class Program
    {
        static void Main(string[] args)
        {
            int mailBatchId = 1;
            var app = new Application();
            Merger mrg = new Merger(ref app);
            var mailingList = mrg.Merge(mailBatchId);
            app.Quit();

            //TODO: insert list to database
            var db = new Exports.SqlExport();
            foreach(var m in mailingList)
            {
                if (db.Exists(m) == false)
                {
                    db.Insert(m);
                }
            }

            Console.WriteLine("Done");
        }
    }
}
