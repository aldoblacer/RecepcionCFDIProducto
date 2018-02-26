using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

//Ingresar en NuGET: Install-Package Microsoft.Exchange.WebServices -Version 2.2.0
namespace VerificadorAF.Mail
{
    internal class mdlVariable
    {
        private string o_nombre = String.Empty;
        private string o_descripcion = String.Empty;

        internal mdlVariable(string NOMBRE, string DESCRIPCION)
        {
            o_nombre = NOMBRE;
            o_descripcion = DESCRIPCION;
        }

        internal string Nombre { get { return o_nombre; } set { o_nombre = value; } }
        internal string Descripcion { get { return o_descripcion; } set { o_descripcion = value; } }
    }
}
