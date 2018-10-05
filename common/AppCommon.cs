using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MyDBQuery.common
{
    public class AppCommon
    {
        public static readonly string IniFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Setting.ini");

        public static readonly string SEC_IMPTABHIS = "ImpTabHis";
    }
}
