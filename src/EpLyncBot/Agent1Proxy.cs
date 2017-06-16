using System;
using System.Net;

namespace EpLyncBot
{
    public class Agent1Proxy : IWebProxy
    {
        public Uri GetProxy(Uri destination)
        {
            return new Uri("http://srv-agent-01:8080");
        }

        public bool IsBypassed(Uri host)
        {
            return false;
        }

        public ICredentials Credentials { get; set; }
    }
}