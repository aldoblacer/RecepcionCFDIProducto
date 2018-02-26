using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using VerificadorAF.Mail;

namespace VerificadorAF.Contenedor
{
    internal class SeleccionArchivos
    {
        public string RegresaError { get; set; }

        internal bool VerificaXML(string myfile, string o_nombre, string o_id, string file)
        {
            StringBuilder sb = new StringBuilder();
            List<mdlParametros> lstparm = new List<mdlParametros>();
            IngresaBuzon ig = new IngresaBuzon();
            SQLCnx grd = new SQLCnx();
            XmlReader reader;
            string o_version = string.Empty;
            string o_tipodoc = string.Empty;
            string o_uuid = string.Empty;
            string o_po = string.Empty;

            try
            {
                if (Path.GetExtension(myfile).ToUpper().Equals(".XML"))
                {
                    try
                    {
                        reader = XmlReader.Create(myfile);
                        while (reader.Read())
                        {
                            if (reader.IsStartElement())
                            {
                                switch (reader.Name)
                                {
                                    case "cfdi:Comprobante":
                                        o_version = reader["version"] ?? "00.00";
                                        if (string.IsNullOrEmpty(o_version) || o_version == "00.00")
                                            o_version = reader["Version"] ?? "00.00";
                                        o_tipodoc = reader["tipoDeComprobante"] ?? "";
                                        if (string.IsNullOrEmpty(o_tipodoc))
                                            o_tipodoc = reader["TipoDeComprobante"] ?? "";
                                        break;
                                    case "tfd:TimbreFiscalDigital":
                                        o_uuid = reader["UUID"] ?? "";
                                        if (string.IsNullOrEmpty(o_uuid))
                                            o_uuid = reader["Uuid"] ?? "";
                                        break;
                                }
                            }
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        grd.GuardaBitacoraProcesoInterno("VerificaXML", file + " documento con formato invalido.", o_id, "0", string.Empty, string.Empty, "0", ig.MIPublico[0].ToString());
                        throw ex;
                    }
                }
                else
                    throw new Exception("No es un archivo XML");

                if (o_version != "3.3")
                    throw new Exception("Version incorrecta de CFDI, Archivo: " + file + "Version: " + o_version.ToString() + "UUID: " + o_uuid);

                if (string.IsNullOrEmpty(o_uuid))
                {
                    grd.GuardaBitacoraProcesoInterno("VerificaXML", file + " documento sin UUID.", o_id, "0", string.Empty, string.Empty, "0", ig.MIPublico[0].ToString());
                    throw new Exception(file + " documento sin UUID.");
                }

                if (o_tipodoc != "P" && o_tipodoc != "I" && o_tipodoc != "E")
                {
                    grd.GuardaBitacoraProcesoInterno("VerificaXML", file + " archivo con tipo de documento invalido.", o_id, "0", string.Empty, string.Empty, "0", ig.MIPublico[0].ToString());
                    throw new Exception(file + " archivo con tipo de documento invalido.");
                }

                RegresaError = null;
                return true;
            }
            catch (Exception ex)
            {
                RegresaError = ex.Message;
                grd.GuardaBitacora("VerificaXML", ex.Message + "Archivo: " + file + "Version: " + o_version.ToString() + "UUID: " + o_uuid);
                return false;
            }
        }
    }
}
