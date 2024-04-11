using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OphiussauPnpPortForward.Models
{
    internal class PortFowardRules
    {
        public string Name      { get; set; }
        public string IpAdress  { get; set; }
        public int    StartPort { get; set; }
        public int    EndPort   { get; set; }
        public string Message  { get; set; }
    }
}
