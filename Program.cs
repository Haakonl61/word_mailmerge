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
            var app = new Application();
            Merger mrg = new Merger(ref app);
            var mailingList = mrg.Merge(Settings.Default.BatchId);
            app.Quit();

            var db = new Exports.SqlExport();
            db.Export(mailingList);

            Console.WriteLine("Done");
        }
    }
}
