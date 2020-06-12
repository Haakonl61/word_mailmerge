using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace word_mailmerge
{
    interface IExport
    {
        void Create<T>(T param);
        Object Read();
        void Update<T>(T param);
        void Delete<T>(T param);
    }
}
