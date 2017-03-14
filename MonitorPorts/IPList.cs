using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MonitorPorts
{
    class IPList
    {
        public string Name { get; set; }
        public string IP { get; set; }
        public string Port { get; set; }
        public string Type { get; set; }
        public bool IsChoose { get; set; }
    }
}
