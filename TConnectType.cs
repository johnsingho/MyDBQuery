using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace MyDBQuery
{
    public enum DbType {
        [Description("SQLServer")]
        SQLServer = 0,
        [Description("Oracle")]
        Oracle = 1
    };

    public class TConnectType
    {
        public DbType Type { get; set; }
        public string ConnStr { get; set; }
    }
}
