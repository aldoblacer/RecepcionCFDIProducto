using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VerificadorAF.Mail
{
    internal class Parametros
    {
        internal class Formatos
        {
            internal const string frmMSGRECIB= "<div>{3}<table><tr><td style=\"font-weight:bold\" align=\"right\">{4}</td><td>{0}</td></tr><tr><td style=\"font-weight:bold\" align=\"right\">{5}</td><td>{1}</td></tr><tr><td style=\"font-weight:bold\" align=\"right\">{6}</td><td>{2}</td></tr></table>{7}<br />{8}</div>";
            internal const string frmMSGNOATT = "<div>{3}<table><tr><td style=\"font-weight:bold\" align=\"right\">{4}</td><td>{0}</td></tr><tr><td style=\"font-weight:bold\" align=\"right\">{5}</td><td>{1}</td></tr><tr><td style=\"font-weight:bold\" align=\"right\">{6}</td><td>{2}</td></tr></table>{7}<br />{8}<br />{9}</div>";
            internal const string frmMSGERRORH = "<table widt=\"100%\"><tr><td colspan=\"7\">A continuación le comunicamos los Errores enontrados&nbsp; en las siguientes facturas:</td></tr><tr><td colspan=\"7\">&nbsp;</td></tr><tr><td style=\"text-align: center; background-color:blue; color:white;\">CORREO</td><td style=\"text-align: center; background-color:blue; color:white;\">EMISOR</td><td style=\"text-align: center; background-color:blue; color:white;\">BUZON</td><td style=\"text-align: center; background-color:blue; color:white;\">FECHA_PROCESO</td><td style=\"text-align: center; background-color:blue; color:white;\">ASUNTO</td><td style=\"text-align: center; background-color:blue; color:white;\">NOMBRE</td><td style=\"text-align: center; background-color:blue; color:white;\">ARCHIVO</td></tr>{0}<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td colspan=\"7\">Favor de revisar.</td></tr></table>";
            internal const string frmMSGERRORD = "<tr><td style=\"text-align: center\">{0}</td><td style=\"text-align: center\">{1}</td><td style=\"text-align: center\">{2}</td><td style=\"text-align: center\">{3}</td><td style=\"text-align: center\">{4}</td><td style=\"text-align: center\">{5}</td><td style=\"text-align: center\">{6}</td></tr>";
            internal const string frmMSGERRHIN = "<table widt=\"100%\"><tr><td colspan=\"4\">A continuación le comunicamos los Errores enontrados&nbsp; en las siguientes facturas:</td></tr><tr><td colspan=\"7\">&nbsp;</td></tr><tr><td style=\"text-align: center; background-color:blue; color:white;\">FECHA_LLEGADA</td><td style=\"text-align: center; background-color:blue; color:white;\">ASUNTO</td><td style=\"text-align: center; background-color:blue; color:white;\">FOLIO</td><td style=\"text-align: center; background-color:blue; color:white;\">DESCRIPCION</td></tr>{0}<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td colspan=\"4\">Favor de revisar.</td></tr></table>";
            internal const string frmMSGERRDIN = "<tr><td style=\"text-align: center\">{0}</td><td style=\"text-align: center\">{1}</td><td style=\"text-align: center\">{2}</td><td style=\"text-align: center\">{3}</td></tr>";

            //internal const string mailMSGERRORH = "<div><br /><table style=\"width:50%; background-color:lightgray; border:double;\"><tr><td  style=\"background-color:gray; color:white\">ASUNTO</td><td>{0}</td></tr></table><br /><table style=\"border:double; width:100%;\"><tr><td align=\"center\" style=\"background-color:gray; color:white\">NO.</td><td align=\"center\" style=\"background-color:gray; color:white\">ARCHIVO</td><td align=\"center\" style=\"background-color:gray; color:white\">FOLIO INTERNO</td><td align=\"center\" style=\"background-color:gray; color:white\">SERIE/FOLIO</td><td align=\"center\" style=\"background-color:gray; color:white\">UUID</td><td align=\"center\" style=\"background-color:gray; color:white\">RESPUESTA</td></tr> {99}</table><br /></div>";
            internal const string mailMSGERRORH = "<div><br /><table style=\"width:50%; background-color:lightgray; border:double;\"><tr><td  style=\"background-color:gray; color:white\">ASUNTO</td><td>{0}</td></tr></table><br /><table style=\"border:double; width:100%;\"> <tr><td align=\"center\" style=\"background-color:gray; color:white\">NO.</td><td align=\"center\" style=\"background-color:gray; color:white\">ARCHIVO</td><td align=\"center\" style=\"background-color:gray; color:white\">FOLIO INTERNO</td><td align=\"center\" style=\"background-color:gray; color:white\">SERIE/FOLIO</td> <td align=\"center\" style=\"background-color:gray; color:white\">UUID</td><td align=\"center\" style=\"background-color:gray; color:white\">ESTATUS</td><td align=\"center\" style=\"background-color:gray; color:white\">PROBLEMAS ENCONTRADOS</td></tr> {99}</table><br /></div>";
            //internal const string mailMSGERRORD = "<tr><td style=\"color:darkgrey\" align=\"right\">{1}</td><td style=\"color:darkgrey\" align=\"left\">{5}</td><td style=\"color:darkgrey\" align=\"right\">{2}</td><td style=\"color:darkgrey\" align=\"left\">{3}</td><td style=\"color:darkgrey\" align=\"left\">{4}</td><td style=\"color:darkgrey\" align=\"left\">{6}</td></tr>";
            internal const string mailMSGERRORD = "<tr><td style=\"color:darkgrey\" align=\"right\">{1}</td><td style=\"color:darkgrey\" align=\"left\">{5}</td><td style=\"color:darkgrey\" align=\"right\">{2}</td><td style=\"color:darkgrey\" align=\"left\">{3}</td><td style=\"color:darkgrey\" align=\"left\">{4}</td><td style=\"color:darkgrey\" align=\"left\">{7}</td><td style=\"color:darkgrey\" align=\"left\">{6}</td></tr>";
            internal const string mailMSGERRORB = "<tr><td style=\"color:cornflowerblue\" align=\"right\">{1}</td><td style=\"color:cornflowerblue\" align=\"right\">{2}</td><td style=\"color:cornflowerblue\" align=\"left\">{3}</td><td style=\"color:cornflowerblue\" align=\"left\">{4}</td><td style=\"color:cornflowerblue\" align=\"left\">{5}</td><td style=\"color:cornflowerblue\" align=\"left\">{6}</td></tr>";
            internal const string mailReporteCifras = "Se envía resumen del Proceso Recepción CFDI Producto.<br><div><br><table style=\"width:50%; background-color:lightgray; border:double;\"><tbody><tr><td align=\"center\" style=\"background-color:gray; color:white\">CONCEPTO</td><td align=\"center\" style=\"background-color:gray; color:white\">CIFRA</td></tr><tr><td style=\"color:black\" align=\"left\">CORREOS PROCESADOS:</td><td style=\"color:black\" align=\"center\">{0}</td></tr><tr><td style=\"color:black\" align=\"left\">ARCHIVOS PROCESADOS:</td><td style=\"color:black\" align=\"center\">{1}</td></tr><tr><td style=\"color:black\" align=\"left\">ACEPTADOS CON PRORROGA EXT:</td><td style=\"color:black\" align=\"center\">{2}</td></tr><tr><td style=\"color:black\" align=\"left\">ACEPTADOS CON PRORROGA INT:</td><td style=\"color:black\" align=\"center\">{3}</td></tr><tr><td style=\"color:black\" align=\"left\">ACEPTADOS SIN PRORROGA:</td><td style=\"color:black\" align=\"center\">{4}</td></tr><tr><td style=\"color:black\" align=\"left\">RECHAZADOS CON PRORROGA INTERNO:</td><td style=\"color:black\" align=\"center\">{5}</td></tr><tr><td style=\"color:black\" align=\"left\">RECHAZADOS CON PRORROGA EXTERNO:</td><td style=\"color:black\" align=\"center\">{6}</td></tr><tr><td style=\"color:black\" align=\"left\">RECHAZADOS SIN PRORROGA:</td><td style=\"color:black\" align=\"center\">{7}</td></tr><tr><td style=\"color:black\" align=\"left\">NC SIN FACTURA CORRECTAS:</td><td style=\"color:black\" align=\"center\">{8}</td></tr><tr><td style=\"color:black\" align=\"left\">NC TOTALES SIN FACUTA:</td><td style=\"color:black\" align=\"center\">{9}</td></tr><tr><td style=\"color:black\" align=\"left\">ENVIADOS A RETEK:</td><td style=\"color:black\" align=\"center\">{10}</td></tr><tr><td style=\"color:black\" align=\"left\">ENVIADOS A RETEK EN PROCESO:</td><td style=\"color:black\" align=\"center\">{11}</td></tr><tr><td style=\"color:black\" align=\"left\">ENVIADOS A RETEK NO PROCESO:</td><td style=\"color:black\" align=\"center\">{12}</td></tr><tr><td style=\"color:black\" align=\"left\">CON ERROR A REVISAR:</td><td style=\"color:black\" align=\"center\">{13}</td></tr></table><br></div>";
        }

        internal class Conexiones
        {
            //PRUEBAS
            //internal const string CNXSQL_CORONAP =  "Server=WTMXCFDI01\\RECEPCIONCFDI; Database=CFDI-MAIL; User ID=CFDI_REC_USR; Password=RL3v4wdFPGCZ9RQp";
            
            //PRODUCCION
            internal const string CNXSQL_CORONAP = "Server=WPMXDB01\\RECEPCIONCFDI; Database=CFDI-MAIL; User ID=CFDI_REC_USR; Password=RL3v4wdFPGCZ9RQp";
        }

        internal class Buzones
        {
            //PRUEBAS
            //internal const string EMAIL = "facturas_producto_prueba>P4sw0rd2017>waldos.com>http://wmx4exc10.waldos.com/EWS/Exchange.asmx";
            //internal const string EMAILEX = "facturaelectronica>Password500>waldos.com>http://wmx4exc10.waldos.com/EWS/Exchange.asmx";
            
            //PRODUCCION
            internal const string EMAIL = "facturas_producto>Password600>waldos.com>http://wmx4exc10.waldos.com/EWS/Exchange.asmx";
            internal const string EMAILEX = "facturaelectronica>Password500>waldos.com>http://wmx4exc10.waldos.com/EWS/Exchange.asmx";
        }

        internal class Asuntos
        {
            internal const string ASURECIB = "Acuse Recibo Factura.";
            internal const string ASUSINAT = "Correo sin Archivos Adjuntos.";
            internal const string ASUCONER = "Correo {0} con Excepcion.";
        }

        internal class Acuses
        {
            internal const string MSGRECIB = " Su correo con los siguientes datos: ";
            internal const string MSGRECIB1 = " Asunto: ";
            internal const string MSGRECIB2 = " Enviado (díay hora): ";
            internal const string MSGRECIB3 = " Recibido (díay hora): ";
            internal const string MSGRECIB4 = " Ha sido recibido. ";
            internal const string MSGRECIB5 = " Saludos.";
        }

        internal class Adjuntos
        {
            internal const string MSGNOATT = " Su correo con los siguientes datos: ";
            internal const string MSGNOATT1 = " Asunto: ";
            internal const string MSGNOATT2 = " Enviado (díay hora): ";
            internal const string MSGNOATT3 = " Recibido (díay hora): ";
            internal const string MSGNOATT4 = " No hemos recibido Archivos Adjuntos en su correo. ";
            internal const string MSGNOATT5 = " Favor de revisar. ";
            internal const string MSGNOATT6 = " Saludos. ";
        }

        internal class Errores
        {
            internal const string MSGRERRO1 = " Ha sido rechazado con la siguiente excepcion: {9}. ";
            internal const string MSGRERRO2 = "  Saludos. ";
        }

        //DIRECTORIOS MAILS
        internal class Direcciones
        {

            internal const bool PRUEBAS = false;

            ////SERVIDOR                    D:\WSI-CFDI\CFDI_Recepcion\Archivos
            internal const string ARCHTMP = "D:\\WSI-CFDi\\CFDI_Recepcion\\Archivos\\TMP\\";
            internal const string ARCHXML = "D:\\WSI-CFDi\\CFDI_Recepcion\\Archivos\\XML\\";

             //ESTOS NO SE ESTÁN USANDO
            internal const string ARCHPDF = "D:\\coronap\\Desktop\\SAT 3_3\\juegos\\PDF\\";
            internal const string ARCHPCP = "D:\\coronap\\Desktop\\SAT 3_3\\juegos\\PCP\\";
            internal const string ARCHPPW = "D:\\coronap\\Desktop\\SAT 3_3\\juegos\\PPW\\";

            //internal const string ARCHTMP = "D:\\XML_STORAGE\\Archivos\\TMP\\";
            //internal const string ARCHXML = "D:\\XML_STORAGE\\Archivos\\XML\\";


        }

        internal class Winr
        {
            //internal const string WINRARF = "C:\\Program Files (x86)\\WinRAR";
            internal const string WINRARF = "C:\\Program Files\\WinRAR";
            
        }

        internal class Carpeta
        {
            internal const string CARRES = "Respaldo";
        }
    }
}
