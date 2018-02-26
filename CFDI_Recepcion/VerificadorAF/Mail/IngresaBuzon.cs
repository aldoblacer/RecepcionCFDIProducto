using Ionic.Zip;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using VerificadorAF.Contenedor;

namespace VerificadorAF.Mail
{
    /// <summary>
    /// Conecta a buzones especificados
    /// </summary>
    internal class IngresaBuzon
    {
        private ExchangeService servicio = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
        private MailItem mi = new MailItem();
        private static string[] o_mi;
        private static StringBuilder o_XML;
        public string[] MIPublico { get { return o_mi; } set { o_mi = value; } }

        /// <summary>
        /// Extrae correos a leer de Config.irx y Discrimina de acuerdo a attachments
        /// </summary>
        /// <param name="lstConst">Lista de Constantes</param>
        internal void IdentificaBuzones()
        {
            List<mdlParametros> lstparm = new List<mdlParametros>();
            MailEnvio mlsndr = new MailEnvio();
            SQLCnx grd = new SQLCnx();
            int MaxID = 0;
            string o_parte = string.Empty;
            string fFecha = FormatoFechaLanguage();
            try
            {

                string o_correos = Parametros.Buzones.EMAIL;

                o_parte = " (1) "; //Extrae cuentas de correo a revisar.
                string[] IdCorreos = o_correos.Split('¬');
                foreach (string itm in IdCorreos)
                {
                    string[] correos = itm.Split('>');
                    MailItem[] mailstr = LeeBuzon(correos[0].ToString(), correos[0].ToString(), correos[1].ToString(), correos[2].ToString(), correos[3].ToString());
                    if (mailstr == null)
                        throw new Exception("Sin Correos");
                    if (mailstr.Count() > 0)
                    {
                        o_parte = " (2) "; //Revisa del buzon leído los correos uno por uno
                        foreach (MailItem Mitem in mailstr)
                        {
                            if (Mitem.Atach)
                            {
                                mi = Mitem;
                                o_parte = " (3) "; //Envía Acuse de Recibido
                                string o_body = Parametros.Formatos.frmMSGRECIB;
                                string repl = o_body.Replace("{0}", Mitem.Subject);
                                repl = repl.Replace("{1}", Mitem.Datein.ToString());
                                repl = repl.Replace("{2}", Mitem.Dateout.ToString());
                                repl = repl.Replace("{3}", Parametros.Acuses.MSGRECIB);
                                repl = repl.Replace("{4}", Parametros.Acuses.MSGRECIB1);
                                repl = repl.Replace("{5}", Parametros.Acuses.MSGRECIB2);
                                repl = repl.Replace("{6}", Parametros.Acuses.MSGRECIB3);
                                repl = repl.Replace("{7}", Parametros.Acuses.MSGRECIB4);
                                repl = repl.Replace("{8}", Parametros.Acuses.MSGRECIB5);
                                if (Mitem.Enviar != null && Mitem.Enviar.Count() > 0)
                                {
                                    o_mi = Mitem.Enviar;
                                    //mlsndr.EnviaMail("ASURECIB", repl, Mitem.Enviar, null);
                                }
                                else
                                {
                                    string[] recip = { Mitem.From };
                                    o_mi = recip;
                                    //mlsndr.EnviaMail("ASURECIB", repl, recip, null);
                                }

                                o_parte = " (4) "; //Prepara Query para guardar en header
                                if (lstparm != null)
                                    lstparm.Clear();
                                lstparm.Add(new mdlParametros("@Receptor", "MAIL"));
                                lstparm.Add(new mdlParametros("@Emisor", !string.IsNullOrEmpty(Mitem.From) ? Mitem.From : " "));
                                lstparm.Add(new mdlParametros("@Buzon", correos[0].ToString()));
                                lstparm.Add(new mdlParametros("@FechaMovimiento", DateTime.Now.ToString(fFecha + " HH:mm:ss")));
                                lstparm.Add(new mdlParametros("@FechaEnvio", Mitem.Dateout.ToString(fFecha + " HH:mm:ss")));
                                lstparm.Add(new mdlParametros("@FechaLlegada", Mitem.Datein.ToString(fFecha + " HH:mm:ss")));
                                lstparm.Add(new mdlParametros("@FechaProceso", Mitem.Datein.ToString(fFecha + " HH:mm:ss")));

                                //lstparm.Add(new mdlParametros("@FechaEnvio", Mitem.Dateout.ToString("yyyy/MM/dd HH:mm:ss")));
                                //lstparm.Add(new mdlParametros("@FechaLlegada", Mitem.Datein.ToString("yyyy/MM/dd HH:mm:ss")));
                                //lstparm.Add(new mdlParametros("@FechaProceso", Mitem.Datein.ToString("yyyy/MM/dd HH:mm:ss")));

                                lstparm.Add(new mdlParametros("@Asunto", !string.IsNullOrEmpty(Mitem.Subject) ? Mitem.Subject : " "));
                                lstparm.Add(new mdlParametros("@Respuesta", !string.IsNullOrEmpty(Mitem.From) ? Mitem.From : " "));
                                lstparm.Add(new mdlParametros("@Estatus", "1"));
                                lstparm.Add(new mdlParametros("@Recibido", !string.IsNullOrEmpty(Mitem.Recibido) ? Mitem.Recibido : " "));
                                o_parte = " (5) "; //Guarda en Base de Datos "HEADER" y obtiene el ID del registro

                                try
                                {
                                    bool mueve = true;
                                    DataSet dsret = new DataSet();
                                    dsret = grd.EjecutaSP("INS_CFDI_HEAD_MAIL", lstparm);
                                    if (dsret != null)
                                        if (dsret.Tables[0].Rows.Count > 0)
                                        {
                                            MaxID = Convert.ToInt32(dsret.Tables[0].Rows[0][0].ToString());
                                            if (MaxID > 0)
                                            {
                                                mueve = ObtieneAttachments(Mitem.MailItemID, null, MaxID, mi);
                                                if (!mueve)
                                                {
                                                    o_parte = " (6) "; //Envía Mensaje de No attachments
                                                    o_body = Parametros.Formatos.frmMSGNOATT;
                                                    repl = o_body.Replace("{0}", Mitem.Subject);
                                                    repl = repl.Replace("{1}", Mitem.Datein.ToString());
                                                    repl = repl.Replace("{2}", Mitem.Dateout.ToString());
                                                    repl = repl.Replace("{3}", Parametros.Adjuntos.MSGNOATT);
                                                    repl = repl.Replace("{4}", Parametros.Adjuntos.MSGNOATT1);
                                                    repl = repl.Replace("{5}", Parametros.Adjuntos.MSGNOATT2);
                                                    repl = repl.Replace("{6}", Parametros.Adjuntos.MSGNOATT3);
                                                    repl = repl.Replace("{7}", Parametros.Adjuntos.MSGNOATT4);
                                                    repl = repl.Replace("{8}", Parametros.Adjuntos.MSGNOATT5);
                                                    repl = repl.Replace("{9}", Parametros.Adjuntos.MSGNOATT6);

                                                    if (Mitem.Enviar != null && Mitem.Enviar.Count() > 0)
                                                    {
                                                        o_mi = Mitem.Enviar;
                                                        //mlsndr.EnviaMail("ASURECIB", repl, Mitem.Enviar, null); //Mitem.Enviar, null);
                                                    }
                                                    else
                                                    {
                                                        string[] recip = { Mitem.From };
                                                       // mlsndr.EnviaMail("ASURECIB", repl, recip, null);
                                                    }
                                                }
                                            }
                                        }
                                }
                                catch { continue; }

                                o_parte = " (6) "; //mueve mensaje a folder de respaldos
                                FolderId fid = LeeFolder(correos[0].ToString(), correos[0].ToString(), correos[1].ToString(), correos[2].ToString(), correos[3].ToString(), "Respaldo", null);
                                if (fid != null)
                                {
                                    EmailMessage message = EmailMessage.Bind(servicio, Mitem.MailItemID);
                                    message.Move(fid);
                                }
                            }
                            else
                            {
                                o_parte = " (6) "; //Envía Mensaje de No attachments
                                string o_body = Parametros.Formatos.frmMSGNOATT;
                                string repl = o_body.Replace("{0}", Mitem.Subject);
                                repl = repl.Replace("{1}", Mitem.Datein.ToString());
                                repl = repl.Replace("{2}", Mitem.Dateout.ToString());
                                repl = repl.Replace("{3}", Parametros.Adjuntos.MSGNOATT);
                                repl = repl.Replace("{4}", Parametros.Adjuntos.MSGNOATT1);
                                repl = repl.Replace("{5}", Parametros.Adjuntos.MSGNOATT2);
                                repl = repl.Replace("{6}", Parametros.Adjuntos.MSGNOATT3);
                                repl = repl.Replace("{7}", Parametros.Adjuntos.MSGNOATT4);
                                repl = repl.Replace("{8}", Parametros.Adjuntos.MSGNOATT5);
                                repl = repl.Replace("{9}", Parametros.Adjuntos.MSGNOATT6);

                                if (Mitem.Enviar != null && Mitem.Enviar.Count() > 0)
                                {
                                    o_mi = Mitem.Enviar;
                                    //mlsndr.EnviaMail("ASURECIB", repl, Mitem.Enviar, null); //Mitem.Enviar, null);
                                }
                                else
                                {
                                    string[] recip = { Mitem.From };
                                    //mlsndr.EnviaMail("ASURECIB", repl, recip, null);
                                }

                                o_parte = " (7) "; //Prepara query para Insertar en header con estatus 2
                                if (lstparm != null)
                                    lstparm.Clear();
                                lstparm.Add(new mdlParametros("@Receptor", "MAIL"));
                                lstparm.Add(new mdlParametros("@Emisor", !string.IsNullOrEmpty(Mitem.From) ? Mitem.From : " "));
                                lstparm.Add(new mdlParametros("@Buzon", correos[0].ToString()));
                                lstparm.Add(new mdlParametros("@FechaMovimiento", DateTime.Now.ToString(fFecha + " HH:mm:ss")));
                                lstparm.Add(new mdlParametros("@FechaEnvio", Mitem.Dateout.ToString(fFecha + " HH:mm:ss")));
                                lstparm.Add(new mdlParametros("@FechaLlegada", Mitem.Datein.ToString(fFecha + " HH:mm:ss")));

                                //lstparm.Add(new mdlParametros("@FechaEnvio", Mitem.Dateout.ToString("dd/MM/yyyy HH:mm:ss")));
                                //lstparm.Add(new mdlParametros("@FechaLlegada", Mitem.Datein.ToString("dd/MM/yyyy HH:mm:ss")));

                                lstparm.Add(new mdlParametros("@FechaProceso", ""));
                                lstparm.Add(new mdlParametros("@Asunto", !string.IsNullOrEmpty(Mitem.Subject) ? Mitem.Subject : " "));
                                lstparm.Add(new mdlParametros("@Respuesta", !string.IsNullOrEmpty(Mitem.From) ? Mitem.From : " "));
                                lstparm.Add(new mdlParametros("@Estatus", "2"));
                                lstparm.Add(new mdlParametros("@Recibido", !string.IsNullOrEmpty(Mitem.Recibido) ? Mitem.Recibido : " "));

                                o_parte = " (8) "; //Guarda en Base de Datos "HEADER" no nos interesa el regreso
                                grd.EjecutaSP("INS_CFDI_HEAD_MAIL", lstparm);

                                o_parte = " (9) "; //se borra el mensaje.
                                FolderId fid = LeeFolder(correos[0].ToString(), correos[0].ToString(), correos[1].ToString(), correos[2].ToString(), correos[3].ToString(), "Respaldo", null);
                                if (fid != null)
                                {
                                    EmailMessage message = EmailMessage.Bind(servicio, Mitem.MailItemID);
                                    message.Move(fid);
                                }

                                //EmailMessage message = EmailMessage.Bind(servicio, Mitem.MailItemID);
                                //message.Delete(DeleteMode.HardDelete);
                            }
                        }
                    }
                    else
                        throw new Exception("Sin Correos.");
                }
            }
            catch (Exception ex)
            {
                grd.GuardaBitacora("IdentificaBuzones", ex.Message);
            }
        }

        /// <summary>
        /// Funcion que devuelve el Foranto fecha de acuerdo al idioma de la instania de SQL
        /// </summary>
        /// <returns></returns>
        private string FormatoFechaLanguage()
        {
            SQLCnx grd = new SQLCnx();
            List<mdlParametros> lstparm = new List<mdlParametros>();
            try
            {
                DataSet dtsIdiomaBD = grd.EjecutaSPP("SEL_CFDI_CONSULTA_LENGUAJE", lstparm);
                if (dtsIdiomaBD.Tables[0].Rows[0][0].ToString() == "us_english")
                {
                    return "MM/dd/yyyy";
                }
                else
                {
                    return "dd/MM/yyyy";
                }
            }
            catch (Exception)
            {
                return "MM/dd/yyyy";
            }
        }
        /// <summary>
        /// Extrae Correos de Buzón
        /// </summary>
        /// <param name="i_mail">Correo</param>
        /// <param name="i_usr">Usuario</param>
        /// <param name="i_pass">Contraseña</param>
        /// <returns></returns>        
        private MailItem[] LeeBuzon(string i_mail, string i_usr, string i_pass, string i_dom, string i_uri)
        {
            SQLCnx grd = new SQLCnx();
            List<mdlParametros> lstparm = new List<mdlParametros>();
            string o_parte = string.Empty;
            try
            {
                o_parte = " (1) "; //Conecta con buzon
                //Con Autodiscover
                //servicio.Credentials = new WebCredentials("facturas_producto", "Password600", "waldos.com");
                //servicio.AutodiscoverUrl("facturas_producto@waldos.com", RedirectionCallback);

                servicio.Credentials = new WebCredentials(i_usr, i_pass, i_dom);
                servicio.Url = new Uri(i_uri);

                FindItemsResults<Item> findResults = null;
                o_parte = " (2) "; //Lee coleccion de correos
                findResults = servicio.FindItems(WellKnownFolderName.Inbox, new ItemView(int.MaxValue));
                ServiceResponseCollection<GetItemResponse> items = servicio.BindToItems(findResults.Select(item => item.Id), new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.From, EmailMessageSchema.ToRecipients));
                if (items == null)
                    throw new Exception("No se encontraron correos.");


                o_parte = " (3) "; //Si hay correos tómalos y guarda sus atributos
                List<MailItem> lstrevisa = new List<MailItem>();
                List<string> lstErrores = new List<string>();
                foreach (var algo in items)
                {
                    try
                    {
                        if (algo != null)
                        {
                            MailItem nvo = new MailItem();
                            nvo.MailItemID = algo.Item.Id;
                            nvo.From = (((Microsoft.Exchange.WebServices.Data.EmailAddress)algo.Item[EmailMessageSchema.From]).Address != null) ? ((Microsoft.Exchange.WebServices.Data.EmailAddress)algo.Item[EmailMessageSchema.From]).Address : string.Empty;
                            nvo.Recipients = ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)algo.Item[EmailMessageSchema.ToRecipients]).Select(recipient => recipient.Address).ToArray();
                            nvo.Subject = (algo.Item.Subject != null) ? algo.Item.Subject : string.Empty;
                            nvo.Body = (algo.Item.Body.ToString() != null) ? algo.Item.Body.ToString() : string.Empty;
                            nvo.Dateout = (algo.Item.DateTimeReceived != null) ? algo.Item.DateTimeReceived : DateTime.Now.Date;
                            nvo.Datein = (algo.Item.DateTimeSent != null) ? algo.Item.DateTimeSent : DateTime.Now.Date;
                            nvo.Atach = algo.Item.HasAttachments;
                            nvo.Recibido = (algo.Item.DisplayTo != null) ? algo.Item.DisplayTo : string.Empty;
                            nvo.Enviar = ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)algo.Item[EmailMessageSchema.ReplyTo]).Select(recibe => recibe.Address).ToArray();
                            lstrevisa.Add(nvo);
                        }
                    }
                    catch (Exception)
                    {
                        //NO SE IDENTIFICÓ EL ERROR
                        //lstErrores.Add((algo.Item.Subject != null) ? algo.Item.Subject : string.Empty);                 
                    }
                }
                var alista = lstrevisa.ToArray();

                //ESTA PERTE SE QUITÓ PORQUE UN CORREO ESTABA GENERANDO ERROR POR NULL EXCEPTION, SE QUIZO
                // AVERIGUAR CUAL ERA EL PROBLEMA PERO NO SE DETECTO AY QUE NO DABA NINGUNA REFERENCIA DE CUAL ERA EL CORREO.


                //o_parte = " (3) "; //Si hay correos tómalos y guarda sus atributos
                //var alista = items.Select(item =>
                //{
                //    return new MailItem()
                //    {
                //        MailItemID = item.Item.Id,
                //        From = (((Microsoft.Exchange.WebServices.Data.EmailAddress)item.Item[EmailMessageSchema.From]).Address != null) ? ((Microsoft.Exchange.WebServices.Data.EmailAddress)item.Item[EmailMessageSchema.From]).Address : string.Empty,
                //        Recipients = ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)item.Item[EmailMessageSchema.ToRecipients]).Select(recipient => recipient.Address).ToArray(),
                //        Subject = (item.Item.Subject != null) ? item.Item.Subject : string.Empty,
                //        Body = (item.Item.Body.ToString() != null) ? item.Item.Body.ToString() : string.Empty,
                //        Dateout = (item.Item.DateTimeReceived != null) ? item.Item.DateTimeReceived : DateTime.Now.Date,
                //        Datein = (item.Item.DateTimeSent != null) ? item.Item.DateTimeSent : DateTime.Now.Date,
                //        Atach = item.Item.HasAttachments,
                //        Recibido = (item.Item.DisplayTo != null) ? item.Item.DisplayTo : string.Empty,
                //        Enviar = ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)item.Item[EmailMessageSchema.ReplyTo]).Select(recibe => recibe.Address).ToArray()

                //        //    MailItemID = item.Item.Id,
                //        //From = ((Microsoft.Exchange.WebServices.Data.EmailAddress)item.Item[EmailMessageSchema.From]).Address,
                //        //Recipients = ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)item.Item[EmailMessageSchema.ToRecipients]).Select(recipient => recipient.Address).ToArray(),
                //        //Subject = item.Item.Subject,
                //        //Body = item.Item.Body.ToString(),
                //        //Dateout = item.Item.DateTimeReceived,
                //        //Datein = item.Item.DateTimeSent,
                //        //Atach = item.Item.HasAttachments,
                //        //Recibido = item.Item.DisplayTo,
                //        //Enviar = ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)item.Item[EmailMessageSchema.ReplyTo]).Select(recibe => recibe.Address).ToArray()


                //    };
                //}).ToArray();

                o_parte = " (4) "; //Regresa coleccion de correos
                return alista;
            }
            catch (Exception ex)
            {
                grd.GuardaBitacora("LeeBuzon", ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Extrae el id de carpeta
        /// </summary>
        /// <param name="i_mail">correo</param>
        /// <param name="i_usr">usuario</param>
        /// <param name="i_pass">password</param>
        /// <param name="i_nombre">nomre de carpeta a extraer id</param>
        /// <param name="lst">lista de constantes</param>
        /// <returns></returns>
        private FolderId LeeFolder(string i_mail, string i_usr, string i_pass, string i_dom, string i_uri, string i_nombre, List<mdlVariable> lst)
        {
            SQLCnx grd = new SQLCnx();
            List<mdlParametros> lstparm = new List<mdlParametros>();
            string o_parte = string.Empty;
            try
            {
                o_parte = " (1) "; //Conecta con buzon
                servicio.Credentials = new WebCredentials(i_usr, i_pass, i_dom);
                servicio.Url = new Uri(i_uri);

                o_parte = " (2) "; //Lee coleccion de folders
                FolderView fv = new FolderView(int.MaxValue);
                FindFoldersResults findResults = servicio.FindFolders(WellKnownFolderName.Inbox, fv);

                if (findResults == null)
                    throw new Exception("No se encontraron folders.");

                o_parte = " (3) "; //obtener id del folder
                var alista = from xfol in findResults
                             where xfol.DisplayName == i_nombre
                             select xfol;

                o_parte = " (4) "; //Regresa id de carpeta
                return alista.ElementAt(0).Id;
            }
            catch (Exception ex)
            {
                grd.GuardaBitacora("LeeFolder", ex.Message);
                return null;
            }
        }

        internal StringBuilder LeeFile(string myfile)
        {
            SQLCnx grd = new SQLCnx();
            string sLine = String.Empty;
            StringBuilder Contenido = new StringBuilder();

            try
            {
                if (File.Exists(myfile))
                {
                    StreamReader w = new StreamReader(myfile);

                    while (sLine != null)
                    {
                        sLine = w.ReadLine();
                        if (sLine != null)
                            Contenido.Append(sLine);
                    }

                    w.Close();
                    return Contenido;
                }
                else
                    throw new Exception("File not found.");
            }
            catch (Exception ex)
            {
                grd.GuardaBitacora("LeeFolder", ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Lee Adjuntos correo por correo
        /// </summary>
        /// <param name="itemId">id de correo</param>
        /// <param name="lstConst">lista de constantes</param>
        /// <param name="MailItemID">ID de Mail</param>
        internal Boolean ObtieneAttachments(ItemId itemId, List<mdlVariable> lstConst, int MailItemID, MailItem mio)
        {
            SQLCnx grd = new SQLCnx();
            List<mdlParametros> lstparm = new List<mdlParametros>();
            string o_parte = string.Empty;
            try
            {
                o_parte = " (1) "; //extrae constantes
                SeleccionArchivos sa = new SeleccionArchivos();
                string o_rutaXml = Parametros.Direcciones.ARCHXML;
                string o_rutaXmlDate = o_rutaXml + DateTime.Now.ToString("ddMMyyyy") + "\\";

                o_parte = " (2) "; //verifica directorios
                if (!Directory.Exists(o_rutaXmlDate))
                    Directory.CreateDirectory(o_rutaXmlDate);

                o_parte = " (3) "; //Extrae Attachments del correo
                int o_ctvo = 1, ctvo22 = 0;
                EmailMessage message = EmailMessage.Bind(servicio, itemId, new PropertySet(ItemSchema.Attachments));
                foreach (Attachment attachment in message.Attachments)
                {
                    if (attachment is FileAttachment)
                    {
                        o_parte = " (4) "; //transforma a attachment
                        FileAttachment fileAttachment = attachment as FileAttachment;

                        o_parte = " (5) "; //identifica el attachment
                        switch (Path.GetExtension(fileAttachment.Name.ToUpper()))
                        {
                            case ".XML":
                                o_parte = " (7) "; //si el attach es XML
                                fileAttachment.Load(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml");
                                if (sa.VerificaXML(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml", MailItemID + "_" + o_ctvo + ".xml", MailItemID.ToString(), fileAttachment.Name))
                                {
                                    o_XML = LeeFile(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml");
                                    if (lstparm != null)
                                        lstparm.Clear();
                                    lstparm.Add(new mdlParametros("@Id", MailItemID.ToString()));
                                    lstparm.Add(new mdlParametros("@Orden", o_ctvo.ToString()));
                                    lstparm.Add(new mdlParametros("@Nombre", fileAttachment.Name));
                                    lstparm.Add(new mdlParametros("@Generado", MailItemID + "_" + o_ctvo + ".xml"));
                                    lstparm.Add(new mdlParametros("@Direccion", o_rutaXmlDate));
                                    lstparm.Add(new mdlParametros("@Estatus", "2"));
                                    lstparm.Add(new mdlParametros("@Xml", o_XML.ToString()));
                                    lstparm.Add(new mdlParametros("@XmlEstatus", string.Empty));
                                    grd.EjecutaSP("INS_CFDI_DETAIL_MAIL", lstparm);
                                    o_ctvo++;
                                    ctvo22++;
                                }
                                else
                                {
                                    if (lstparm != null)
                                        lstparm.Clear();
                                    lstparm.Add(new mdlParametros("@IdFactura", MailItemID.ToString()));
                                    //lstparm.Add(new mdlParametros("@Fecha", DateTime.Now.ToString("yyyy/MM/dd")));
                                    lstparm.Add(new mdlParametros("@Fecha", DateTime.Now.ToString(FormatoFechaLanguage())));
                                    lstparm.Add(new mdlParametros("@Tipo", "I"));
                                    lstparm.Add(new mdlParametros("@Descripcion", fileAttachment.Name + sa.RegresaError));
                                    if (mio.Enviar != null && mio.Enviar.Count() > 0)
                                        lstparm.Add(new mdlParametros("@Correo", mio.Enviar[0].ToString()));
                                    else
                                        lstparm.Add(new mdlParametros("@Correo", mio.From));
                                    grd.EjecutaSP("INS_CFDI_LOG_MAIL", lstparm);
                                    try { File.Delete(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml"); }
                                    catch { }
                                    goto brinca;
                                }
                                break;
                            case ".ZIP":
                                o_parte = " (8) "; //si el attach es ZIP
                                fileAttachment.Load(Path.GetTempPath() + Path.GetFileNameWithoutExtension(fileAttachment.Name) + "_" + o_ctvo + ".zip");
                                DescomprimeZip(Path.GetTempPath() + Path.GetFileNameWithoutExtension(fileAttachment.Name) + "_" + o_ctvo + ".zip", ref o_ctvo, MailItemID);
                                break;
                            case ".RAR":
                                o_parte = " (9) "; //si el attach es RAR
                                fileAttachment.Load(Path.GetTempPath() + Path.GetFileNameWithoutExtension(fileAttachment.Name) + "_" + o_ctvo + ".rar");
                                DescomprimeRar(Path.GetTempPath() + Path.GetFileNameWithoutExtension(fileAttachment.Name) + "_" + o_ctvo + ".rar", lstConst, ref o_ctvo, MailItemID);
                                break;
                            default:
                            brinca:
                                o_parte = " (10) "; //si el attach es otro
                                if (lstparm != null)
                                    lstparm.Clear();
                                lstparm.Add(new mdlParametros("@Id", MailItemID.ToString()));
                                lstparm.Add(new mdlParametros("@Orden", o_ctvo.ToString()));
                                lstparm.Add(new mdlParametros("@Nombre", fileAttachment.Name));
                                lstparm.Add(new mdlParametros("@Generado", "X"));
                                lstparm.Add(new mdlParametros("@Direccion", "X"));
                                lstparm.Add(new mdlParametros("@Estatus", "1"));
                                lstparm.Add(new mdlParametros("@Xml", string.Empty));
                                lstparm.Add(new mdlParametros("@XmlEstatus", string.Empty));
                                grd.EjecutaSP("INS_CFDI_DETAIL_MAIL", lstparm);
                                o_ctvo++;
                                break;
                        }
                    }
                }

                if (ctvo22 == 0)
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                grd.GuardaBitacora("ObtieneAttachments", ex.Message);
                return true;
            }
        }

        /// <summary>
        /// Descomprime Formato Zip
        /// </summary>
        /// <param name="o_zipath">Ruta del archivo zip</param>
        /// <param name="lstConst">lista de Constantes</param>
        /// <param name="o_ctvo">objeto consecutivo por referencia</param>
        /// <param name="MailItemID">ID de ITEM</param>
        internal Exception DescomprimeZip(string o_zipath, ref int o_ctvo, int MailItemID)
        {
            List<mdlParametros> lstparm = new List<mdlParametros>();
            SeleccionArchivos sa = new SeleccionArchivos();
            SQLCnx grd = new SQLCnx();
            string o_parte = string.Empty;

            try
            {
                o_parte = " (1) "; //extrae constantes
                string o_rutaTmp = Parametros.Direcciones.ARCHTMP;
                string o_temp = o_rutaTmp + "Unzip\\";
                string o_rutaXml = Parametros.Direcciones.ARCHXML;
                string o_rutaXmlDate = o_rutaXml + DateTime.Now.ToString("ddMMyyyy") + "\\";

                o_parte = " (2) "; // Lee archivo
                using (ZipFile zip = ZipFile.Read(o_zipath))
                {
                    FileInfo o_archivo = new FileInfo(o_zipath);

                    o_parte = " (3) "; //Verifica directorios
                    if (Directory.Exists(o_temp))
                        Directory.Delete(o_temp, true);
                    if (!Directory.Exists(o_temp))
                        Directory.CreateDirectory(o_temp);

                    o_parte = " (4) "; //Extrae zip
                    zip.ExtractAll(o_temp);
                    zip.Dispose();

                    o_parte = " (5) "; //Lee archivos de directorio temporal
                    DirectoryInfo di = new DirectoryInfo(o_temp);
                    foreach (var fi in di.GetFiles("*", SearchOption.AllDirectories))
                    {
                        o_parte = " (6) "; //archivo uno por uno
                        switch (Path.GetExtension(fi.Name.ToUpper()))
                        {
                            case ".XML":
                                o_parte = " (8) "; //caso xml
                                File.Move(fi.FullName, o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml");
                                if (sa.VerificaXML(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml", MailItemID + "_" + o_ctvo + ".xml", MailItemID.ToString(), fi.Name.ToString()))
                                {
                                    o_XML = LeeFile(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml");
                                    if (lstparm != null)
                                        lstparm.Clear();
                                    lstparm.Add(new mdlParametros("@Id", MailItemID.ToString()));
                                    lstparm.Add(new mdlParametros("@Orden", o_ctvo.ToString()));
                                    lstparm.Add(new mdlParametros("@Nombre", fi.Name));
                                    lstparm.Add(new mdlParametros("@Generado", MailItemID + "_" + o_ctvo + ".xml"));
                                    lstparm.Add(new mdlParametros("@Direccion", o_rutaXmlDate));
                                    lstparm.Add(new mdlParametros("@Estatus", "2"));
                                    lstparm.Add(new mdlParametros("@Xml", o_XML.ToString()));
                                    lstparm.Add(new mdlParametros("@XmlEstatus", string.Empty));
                                    grd.EjecutaSP("INS_CFDI_DETAIL_MAIL", lstparm);
                                    o_ctvo++;
                                }
                                else
                                {
                                    if (lstparm != null)
                                        lstparm.Clear();
                                    lstparm.Add(new mdlParametros("@IdFactura", MailItemID.ToString()));
                                    lstparm.Add(new mdlParametros("@Fecha", DateTime.Now.ToString()));
                                    lstparm.Add(new mdlParametros("@Tipo", "I"));
                                    lstparm.Add(new mdlParametros("@Descripcion", Path.GetFileName(fi.Name) + " no es un archivo válido. "));
                                    if (mi.Enviar != null && mi.Enviar.Count() > 0)
                                        lstparm.Add(new mdlParametros("@Correo", mi.Enviar[0].ToString()));
                                    else
                                        lstparm.Add(new mdlParametros("@Correo", mi.From));
                                    grd.EjecutaSP("INS_CFDI_LOG_MAIL", lstparm);
                                    File.Delete(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml");
                                    goto brinca;
                                }
                                break;
                            default:
                            brinca:
                                o_parte = " (9) "; //caso otro
                                if (lstparm != null)
                                    lstparm.Clear();
                                lstparm.Add(new mdlParametros("@Id", MailItemID.ToString()));
                                lstparm.Add(new mdlParametros("@Orden", o_ctvo.ToString()));
                                lstparm.Add(new mdlParametros("@Nombre", fi.Name));
                                lstparm.Add(new mdlParametros("@Generado", "X"));
                                lstparm.Add(new mdlParametros("@Direccion", "X"));
                                lstparm.Add(new mdlParametros("@Estatus", "1"));
                                lstparm.Add(new mdlParametros("@Xml", string.Empty));
                                lstparm.Add(new mdlParametros("@XmlEstatus", string.Empty));
                                grd.EjecutaSP("INS_CFDI_DETAIL_MAIL", lstparm);
                                o_ctvo++;
                                break;
                        }
                    }
                    o_parte = " (10) "; //borra directorio temporal
                    if (Directory.Exists(o_temp))
                        Directory.Delete(o_temp, true);
                }

                return null;
            }
            catch (Exception ex)
            {
                grd.GuardaBitacora("DescomprimeZip", ex.Message);
                return ex;
            }
        }

        /// <summary>
        /// Descomprime formato Rar >> Utiliza Winrar
        /// </summary>
        /// <param name="o_rarpath">Ruta del archivo zip</param>
        /// <param name="lstConst">lista de Constantes</param>
        /// <param name="o_ctvo">objeto consecutivo por referencia</param>
        /// <param name="MailItemID">ID de ITEM</param>
        internal Exception DescomprimeRar(string o_rarpath, List<mdlVariable> lstConst, ref int o_ctvo, int MailItemID)
        {
            List<mdlParametros> lstparm = new List<mdlParametros>();
            SeleccionArchivos sa = new SeleccionArchivos();
            SQLCnx grd = new SQLCnx();
            string o_parte = string.Empty;

            try
            {
                o_parte = " (1) "; // extrae Constantes
                string o_rutaTmp = Parametros.Direcciones.ARCHTMP;
                string o_temp = o_rutaTmp + "Unzip\\";
                string o_rutaXml = Parametros.Direcciones.ARCHXML;
                string o_rutaXmlDate = o_rutaXml + DateTime.Now.ToString("ddMMyyyy") + "\\";
                string o_rutawinr = Parametros.Winr.WINRARF;

                o_parte = " (2) "; //verifica directorios
                if (Directory.Exists(o_temp))
                    Directory.Delete(o_temp, true);
                if (!Directory.Exists(o_temp))
                    Directory.CreateDirectory(o_temp);

                o_parte = " (3) "; //llama winrar y extrae a dir temporal
                string destinationFolder = o_rarpath.Remove(o_rarpath.LastIndexOf('.'));
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.FileName = o_rutawinr + "winrar.exe";
                p.StartInfo.Arguments = string.Format(@"x -s ""{0}"" *.* ""{1}\""", o_rarpath, o_temp);
                p.Start();
                p.WaitForExit();

                o_parte = " (4) "; //obtiene archivos de directorio temporal
                DirectoryInfo di = new DirectoryInfo(o_temp);
                foreach (var fi in di.GetFiles("*", SearchOption.AllDirectories))
                {
                    o_parte = " (5) "; //discrimina archivos
                    string o_replqry = string.Empty;
                    switch (Path.GetExtension(fi.Name.ToUpper()))
                    {
                        case ".XML":
                            o_parte = " (7) "; // caso xml
                            File.Move(fi.FullName, o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml");
                            if (sa.VerificaXML(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml", MailItemID + "_" + o_ctvo + ".xml", MailItemID.ToString(), fi.Name.ToString()))
                            {
                                o_XML = LeeFile(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml");
                                if (lstparm != null)
                                    lstparm.Clear();
                                lstparm.Add(new mdlParametros("@Id", MailItemID.ToString()));
                                lstparm.Add(new mdlParametros("@Orden", o_ctvo.ToString()));
                                lstparm.Add(new mdlParametros("@Nombre", fi.Name));
                                lstparm.Add(new mdlParametros("@Generado", MailItemID + "_" + o_ctvo + ".xml"));
                                lstparm.Add(new mdlParametros("@Direccion", o_rutaXmlDate));
                                lstparm.Add(new mdlParametros("@Estatus", "2"));
                                lstparm.Add(new mdlParametros("@Xml", o_XML.ToString()));
                                lstparm.Add(new mdlParametros("@XmlEstatus", string.Empty));
                                grd.EjecutaSP("INS_CFDI_DETAIL_MAIL", lstparm);
                                o_ctvo++;
                            }
                            else
                            {
                                if (lstparm != null)
                                    lstparm.Clear();
                                lstparm.Add(new mdlParametros("@IdFactura", MailItemID.ToString()));
                                lstparm.Add(new mdlParametros("@Fecha", DateTime.Now.ToString()));
                                lstparm.Add(new mdlParametros("@Tipo", "I"));
                                lstparm.Add(new mdlParametros("@Descripcion", Path.GetFileName(fi.Name) + " no es un archivo válido. "));
                                if (mi.Enviar != null && mi.Enviar.Count() > 0)
                                    lstparm.Add(new mdlParametros("@Correo", mi.Enviar[0].ToString()));
                                else
                                    lstparm.Add(new mdlParametros("@Correo", mi.From));
                                grd.EjecutaSP("INS_CFDI_LOG_MAIL", lstparm);
                                File.Delete(o_rutaXmlDate + MailItemID + "_" + o_ctvo + ".xml");
                                goto brinca;
                            }
                            break;
                        default:
                        brinca:
                            o_parte = " (8) "; //caso otro
                            if (lstparm != null)
                                lstparm.Clear();
                            lstparm.Add(new mdlParametros("@Id", MailItemID.ToString()));
                            lstparm.Add(new mdlParametros("@Orden", o_ctvo.ToString()));
                            lstparm.Add(new mdlParametros("@Nombre", fi.Name));
                            lstparm.Add(new mdlParametros("@Generado", "X"));
                            lstparm.Add(new mdlParametros("@Direccion", "X"));
                            lstparm.Add(new mdlParametros("@Estatus", "1"));
                            lstparm.Add(new mdlParametros("@Xml", string.Empty));
                            lstparm.Add(new mdlParametros("@XmlEstatus", string.Empty));
                            grd.EjecutaSP("INS_CFDI_DETAIL_MAIL", lstparm);
                            o_ctvo++;
                            break;
                    }
                }
                o_parte = " (9) "; //elimina directorio temporal
                if (Directory.Exists(o_temp))
                    Directory.Delete(o_temp, true);

                return null;
            }
            catch (Exception ex)
            {
                grd.GuardaBitacora("DescomprimeRar", ex.Message);
                return ex;
            }
        }
    }

    /// <summary>
    /// Envío de Correos a buzón de salida
    /// </summary>
    internal class MailEnvio
    {
        private ExchangeService servicio = new ExchangeService(ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Envío de Correos
        /// </summary>
        /// <param name="i_sub">ID en Config irx de Asunto</param>
        /// <param name="i_Body">Cuerpo de Correo</param>
        /// <param name="i_mail">Correo Destino</param>
        /// <param name="lst">Lista de Constantes</param>
        internal Exception EnviaMail(string i_sub, string i_Body, string[] i_mail, List<mdlVariable> lst)
        {
            SQLCnx grd = new SQLCnx();
            List<mdlParametros> lstparm = new List<mdlParametros>();
            string o_parte = string.Empty;
            string correos = string.Empty ;
            string conTRIM = string.Empty;
            try
            {
                o_parte = " (1) "; //Extrae Constantes
                string o_mailext = Parametros.Buzones.EMAILEX;
                string o_subject = string.Empty;
                switch (i_sub)
                {
                    case "ASURECIB":
                        o_subject = Parametros.Asuntos.ASURECIB;
                        break;
                    case "ASUSINAT":
                        o_subject = Parametros.Asuntos.ASUSINAT;
                        break;
                    case "ASUCONER":
                        o_subject = Parametros.Asuntos.ASUCONER;
                        break;
                    default:
                        o_subject = i_sub;
                        break;
                }

                o_parte = " (2) "; //separa descripcion
                string[] Menvio = o_mailext.Split('>');

                o_parte = " (3) "; //Conecta con mail de envío
                servicio.Credentials = new WebCredentials(Menvio[0].ToString(), Menvio[1].ToString(), Menvio[2].ToString());
                servicio.Url = new Uri(Menvio[3].ToString());

                o_parte = " (4) "; //Arma correo y envía
                EmailMessage message = new EmailMessage(servicio);
                message.Subject = o_subject;
                message.Body = i_Body;

                message.BccRecipients.Add("blancasaf@waldos.com");
                message.BccRecipients.Add("vieyrad@waldos.com");
             
                foreach (string itm in i_mail)
                {
                    if (itm != null && !string.IsNullOrEmpty(itm))
                    {
                        string[] subMails = itm.Split(';');
                        foreach (string mail in subMails)
                        {
                            //conTRIM = mail.Trim();
                            if (mail != null && !string.IsNullOrEmpty(mail))
                            {
                                if (isValidEmail(mail))
                                {
                                    correos += ";" + mail;
                                    message.ToRecipients.Add(mail.Trim());
                                }                               
                            }
                        }   
                    }
                }
                message.SendAndSaveCopy();
                //message.Send();
                return null;
            }
            catch (Exception ex)
            {
                grd.GuardaBitacoraEnvioMail("EnviaMail", ex.Message, correos);
                return ex;
            }
        }

        internal Exception EnviaMailReporte(List<string> lst_recep, string i_Body)
        {
            SQLCnx grd = new SQLCnx();
            List<mdlParametros> lstparm = new List<mdlParametros>();
            string o_parte = string.Empty;
            try
            {
                string o_mailext = Parametros.Buzones.EMAILEX;
                o_parte = " (2) "; //separa descripcion
                string[] Menvio = o_mailext.Split('>');

                o_parte = " (3) "; //Conecta con mail de envío
                servicio.Credentials = new WebCredentials(Menvio[0].ToString(), Menvio[1].ToString(), Menvio[2].ToString());
                servicio.Url = new Uri(Menvio[3].ToString());

                o_parte = " (4) "; //Arma correo y envía
                EmailMessage message = new EmailMessage(servicio);
                message.Subject = "Resumen Proceso Recepción CFDI Producto" + DateTime.Now.Date.ToString("dd/MM/yyyy");
                message.Body = i_Body;



                foreach (string itm in lst_recep)
                {
                    if (itm != null && !string.IsNullOrEmpty(itm))
                    {
                        string[] subMails = itm.Split(';');
                        foreach (string mail in subMails)
                        {
                            if (mail != null && !string.IsNullOrEmpty(mail))
                            {
                                if (isValidEmail(mail))
                                {
                                    message.ToRecipients.Add(mail);
                                }
                            }
                        }
                    }
                }
                message.SendAndSaveCopy();

                return null;
            }
            catch (Exception ex)
            {
                grd.GuardaBitacora("EnviaMail_Reporte", ex.Message);
                return ex;
            }
        }
        
        internal bool isValidEmail(string input)
        {
            try
            {
                var email = new System.Net.Mail.MailAddress(input);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

   
    /// <summary>
    /// Modelo para recuperar correos
    /// </summary>
    internal class MailItem
    {
        internal ItemId MailItemID;
        internal string From;
        internal string[] Recipients;
        internal string Subject;
        internal string Body;
        internal DateTime Datein;
        internal DateTime Dateout;
        internal bool Atach;
        internal string Recibido;
        internal string[] Enviar;
    }

    /// <summary>
    /// Modelo para recuperar carpetas
    /// </summary>
    internal class FolderBind
    {
        internal ExchangeService service;
        internal FolderId id;
        internal PropertySet propertySet;
    }
}
