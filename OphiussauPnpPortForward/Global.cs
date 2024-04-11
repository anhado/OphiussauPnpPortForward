using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FirewallManager;
using NetFwTypeLib;
using Newtonsoft.Json;
using Open.Nat;
using OphiussauPnpPortForward.Models;
using OphiussauPnpPortForward.Properties;
using Action = FirewallManager.Action;

namespace OphiussauPnpPortForward
{
    internal static class Global
    {
        internal static List<PortFowardRules> Rules = new List<PortFowardRules>();

        public static void Initialize()
        {
            try
            {
                Rules = JsonConvert.DeserializeObject<List<PortFowardRules>>(File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rules.json")));

            }
            catch (Exception e)
            {
                MessageBox.Show("Error Loading Rules");
            }
        }

        public static async void OpenRouterPort(PortFowardRules obj)
        {
            try
            {
                var discoverer = new NatDiscoverer();
                var device = await discoverer.DiscoverDeviceAsync();


                for (int port = obj.StartPort; port <= obj.EndPort; port++)
                {
                    var mapping = await device.GetSpecificMappingAsync(Protocol.TcpUpd, port);

                    IPAddress ip = IPAddress.Parse(obj.IpAdress);

                    if (mapping == null)
                        await device.CreatePortMapAsync(new Mapping(Protocol.TcpUpd, ip, port, port, 9999, obj.Name));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding mapping :" + ex.Message);
            }
        }

        public static async void RemoveRouterPort(PortFowardRules obj)
        {
            try
            {
                var discoverer = new NatDiscoverer();
                var device = await discoverer.DiscoverDeviceAsync();


                for (int port = obj.StartPort; port <= obj.EndPort; port++)
                {
                    var serverMapping = await device.GetSpecificMappingAsync(Protocol.TcpUpd, port);
                    if (serverMapping != null)
                        await device.DeletePortMapAsync(new Mapping(Protocol.TcpUpd, port, port, obj.Name));

                }
                 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding mapping :" + ex.Message);
            }
        }

        public static void OpenFireWallPort(PortFowardRules obj)
        {
            var fw = new FirewallCom();

            //TCP
            try
            {
                var rulesTcp = fw.GetRules().ToList().FindAll(r => r.Name == obj.Name + " TCP");

                if (rulesTcp.Count > 0)
                    rulesTcp.ForEach(r => { fw.RemoveRule(r.Name); });

                string ports = "";


                for (int port = obj.StartPort; port <= obj.EndPort; port++)
                {
                    if (ports != "")
                        ports += ",";

                    ports += $"{port}";
                }


                var rule = new Rule
                {
                    Action = Action.Allow,
                    Description = "Ophiussa Server Manager - " + obj.Name,
                    Direction = Direction.In,
                    Protocol = ProtocolPort.Tcp,
                    Name = obj.Name + " TCP",
                    RemotePorts = "*",
                    InterfaceTypes = "All",
                    Profiles = ProfileType.All,
                    EdgeTraversal = false,
                    LocalAddresses = "*",
                    RemoteAddresses = "*",
                    LocalPorts = ports,
                    Enabled = true
                };
                AddRule(rule);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding TCP Rule :" + ex.Message);
            }

            //UDP
            try
            {
                var rulesTcp = fw.GetRules().ToList().FindAll(r => r.Name == obj.Name + " UDP");

                if (rulesTcp.Count > 0)
                    rulesTcp.ForEach(r => { fw.RemoveRule(r.Name); });

                string ports = "";
                for (int port = obj.StartPort; port <= obj.EndPort; port++)
                {
                    if (ports != "")
                        ports += ",";

                    ports += $"{port}";
                }

                var rule = new Rule
                {
                    Action = Action.Allow,
                    Description = "Ophiussa uPnP - " + obj.Name,
                    Direction = Direction.In,
                    Protocol = ProtocolPort.Udp,
                    Name = obj.Name + " UDP",
                    RemotePorts = "*",
                    InterfaceTypes = "All",
                    Profiles = ProfileType.All,
                    EdgeTraversal = false,
                    LocalAddresses = "*",
                    RemoteAddresses = "*",
                    LocalPorts = ports,
                    Enabled = true
                };

                AddRule(rule);

                MessageBox.Show("Mapping added/Updated");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding UDP Rule :" + ex.Message);
            }
        }

        public static void RemoveFirewallRules(PortFowardRules obj)
        {
            var fw = new FirewallCom();

            //TCP
            try
            {
                var rulesTcp = fw.GetRules().ToList().FindAll(r => r.Name == obj.Name + " TCP");

                if (rulesTcp.Count > 0)
                    rulesTcp.ForEach(r => { fw.RemoveRule(r.Name); });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error delete TCP Rule :" + ex.Message);
            }

            //UDP
            try
            {
                var rulesTcp = fw.GetRules().ToList().FindAll(r => r.Name == obj.Name + " UDP");

                if (rulesTcp.Count > 0)
                    rulesTcp.ForEach(r => { fw.RemoveRule(r.Name); });
                MessageBox.Show("Mapping deleted");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error delete UDP Rule :" + ex.Message);
            }
        }


        public static bool IsPortOpen(string name, int port)
        {
            try
            {
                var fw = new FirewallCom();
                var rules = fw.GetRules().ToList().FindAll(r => r.Name == name + " TCP");

                bool find = false;
                if (rules.Count > 0)
                    rules.ForEach(r =>
                    {
                        if (r.RemotePorts.Contains(port.ToString()) || r.LocalPorts.Contains(port.ToString())) find = true;
                    });
                if (!find) return false;
                rules = fw.GetRules().ToList().FindAll(r => r.Name == name + " UDP");

                if (rules.Count > 0)
                    rules.ForEach(r =>
                    {
                        if (r.RemotePorts.Contains(port.ToString()) || r.LocalPorts.Contains(port.ToString())) find = true;
                    });

                return find;
            }
            catch
            {
                return false;
            }
        }

        private static void AddRule(Rule rule)
        {
            var firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
            ;

            var rule1 = ConvertRule(() =>
            {
                // ISSUE: variable of a compiler-generated type
                var instance = (INetFwRule)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
                return instance;
            }, rule);

            firewallPolicy.Rules.Add(rule1);
        }

        private static INetFwRule ConvertRule(Func<INetFwRule> func, Rule item)
        {
            // ISSUE: variable of a compiler-generated type
            var netFwRule1 = func();
            netFwRule1.Action = Tools.Convert(item.Action);
            netFwRule1.Protocol = (int)Tools.Convert(item.Protocol);
            netFwRule1.ApplicationName = item.ApplicationName;
            netFwRule1.Description = item.Description;
            netFwRule1.Direction = Tools.Convert(item.Direction);
            netFwRule1.EdgeTraversal = item.EdgeTraversal;
            netFwRule1.Enabled = item.Enabled;
            netFwRule1.Grouping = item.Grouping;
            if (item.IcmpTypesAndCodes != null)
                netFwRule1.IcmpTypesAndCodes = item.IcmpTypesAndCodes;
            netFwRule1.InterfaceTypes = item.InterfaceTypes;
            netFwRule1.Interfaces = item.Interfaces;
            netFwRule1.LocalAddresses = item.LocalAddresses;
            if (item.LocalPorts != null)
                netFwRule1.LocalPorts = item.LocalPorts;
            netFwRule1.Name = item.Name;
            netFwRule1.Profiles = (int)Tools.Convert(item.Profiles);
            netFwRule1.RemoteAddresses = item.RemoteAddresses;
            if (item.RemotePorts != null)
                netFwRule1.RemotePorts = item.RemotePorts;
            netFwRule1.serviceName = item.ServiceName;
            // ISSUE: variable of a compiler-generated type
            var netFwRule2 = netFwRule1;
            return netFwRule2;
        }

        public static int ToInt(this string prop)
        {
            if (int.TryParse(prop, NumberStyles.Any, CultureInfo.InvariantCulture, out int val)) return val;
            return 0;
        }

    }
}
