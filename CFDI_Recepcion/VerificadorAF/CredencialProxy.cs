using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

namespace VerificadorAF
{
    public class CredencialProxy:IWebProxy
    {
        public ICredentials Credentials
        {
            get { return new NetworkCredential("blancasaf", "Password199","waldos"); }
            //or get { return new NetworkCredential("user", "password","domain"); }
            set { }
        }

        public Uri GetProxy(Uri destination)
        {
            return new Uri("http://my.proxy:8080");
        }

        public bool IsBypassed(Uri host)
        {
            return false;
        }
    }
}
