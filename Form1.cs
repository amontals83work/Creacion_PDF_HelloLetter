using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Contracts;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.StyledXmlParser.Jsoup.Nodes;
using static iText.IO.Util.IntHashtable;
using iText.Layout.Element;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Action;
using System.Data.SqlClient;
using Microsoft.Win32;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography;

//vigilar 749065 573829
namespace Creacion_PDF_HelloLetter
{
    //CRISALIDA
    public partial class Form1 : Form
    {
        ConexionDB conn = new ConexionDB();
        MCCommand mcComm = new MCCommand();

        //Ruta para guardar archivo
        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public Form1()
        {
            InitializeComponent();
            CenterToScreen();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CrearPDF();
        }

        public string CodigoPostal(int cp)
        {
            string provincia = "";
            if (cp < 1999) return provincia = "ÁLAVA";
            else if (cp < 2999) return provincia = "ALBACETE";
            else if (cp < 3999) return provincia = "ALICANTE";
            else if (cp < 4999) return provincia = "ALMERÍA";
            else if (cp < 5999) return provincia = "ÁVILA";
            else if (cp < 6999) return provincia = "BADAJOZ";
            else if (cp < 7999) return provincia = "BALEARES";
            else if (cp < 8999) return provincia = "BARCELONA";
            else if (cp < 9999) return provincia = "BURGOS";
            else if (cp < 10999) return provincia = "CÁCERES";
            else if (cp < 11999) return provincia = "CÁDIZ";
            else if (cp < 12999) return provincia = "CASTELLÓN";
            else if (cp < 13999) return provincia = "CIUDAD REAL";
            else if (cp < 14999) return provincia = "CÓRDOBA";
            else if (cp < 15999) return provincia = "LA CORUÑA";
            else if (cp < 16999) return provincia = "CUENCA";
            else if (cp < 17999) return provincia = "GERONA";
            else if (cp < 18999) return provincia = "GRANADA";
            else if (cp < 19999) return provincia = "GUADALAJARA";
            else if (cp < 20999) return provincia = "GUIPÚZCOA";
            else if (cp < 21999) return provincia = "HUELVA";
            else if (cp < 22999) return provincia = "HUESCA";
            else if (cp < 23999) return provincia = "JAÉN";
            else if (cp < 24999) return provincia = "LEÓN";
            else if (cp < 25999) return provincia = "LÉRIDA";
            else if (cp < 26999) return provincia = "LA RIOJA";
            else if (cp < 27999) return provincia = "LUGO";
            else if (cp < 28999) return provincia = "MADRID";
            else if (cp < 29999) return provincia = "MÁLAGA";
            else if (cp < 30999) return provincia = "MURCIA";
            else if (cp < 31999) return provincia = "NAVARRA";
            else if (cp < 32999) return provincia = "OURENSE";
            else if (cp < 33999) return provincia = "ASTURIAS";
            else if (cp < 34999) return provincia = "PALENCIA";
            else if (cp < 35999) return provincia = "LAS PALMAS";
            else if (cp < 36999) return provincia = "PONTEVEDRA";
            else if (cp < 37999) return provincia = "SALAMANCA";
            else if (cp < 38999) return provincia = "S. C. DE TENERIFE";
            else if (cp < 39999) return provincia = "CANTABRIA";
            else if (cp < 40999) return provincia = "SEGOVIA";
            else if (cp < 41999) return provincia = "SEVILLA";
            else if (cp < 42999) return provincia = "SORIA";
            else if (cp < 43999) return provincia = "TARRAGONA";
            else if (cp < 44999) return provincia = "TERUEL";
            else if (cp < 45999) return provincia = "TOLEDO";
            else if (cp < 46999) return provincia = "VALENCIA";
            else if (cp < 47999) return provincia = "VALLADOLID";
            else if (cp < 48999) return provincia = "VIZCAYA";
            else if (cp < 49999) return provincia = "ZAMORA";
            else if (cp < 50999) return provincia = "ZARAGOZA";
            else if (cp < 51999) return provincia = "CEUTA";
            else if (cp < 52999) return provincia = "MELILLA";
            else return provincia;
        }

        private void CrearPDF()
        {
            ConexionDB.AbrirConexion();

            string rutaArchivos = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida");
            string rutaArchivosGeneral = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida/General");
            string rutaArchivosNavarra = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida/Navarra");
            string rutaArchivosValencia = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida/Valencia");
            if (!Directory.Exists(rutaArchivos)) Directory.CreateDirectory(rutaArchivos);
            if (!Directory.Exists(rutaArchivosGeneral)) Directory.CreateDirectory(rutaArchivosGeneral);
            if (!Directory.Exists(rutaArchivosNavarra)) Directory.CreateDirectory(rutaArchivosNavarra);
            if (!Directory.Exists(rutaArchivosValencia)) Directory.CreateDirectory(rutaArchivosValencia);

            mcComm.command.Connection = conn.ObtenerConexion();
            mcComm.CommandText = "SELECT OriginalAccountId,  REPLACE(FullName, '?', '') AS FullName_reemplazada, REPLACE(FullName2, '?', '') AS FullName2_reemplazada, REPLACE(Address1, '?', '') AS Address1_reemplazada, Zip, REPLACE(City, '?', '') AS City_reemplazada FROM TempPagantis";

            using (IDataReader reader = mcComm.ExecuteReader())
            {
                while (reader.Read())
                {
                    string referencia = reader["OriginalAccountId"].ToString();
                    string importe = string.Empty;
                    string importe2Decimales = string.Empty;

                    using (SqlCommand innerCommand = new SqlCommand())
                    {
                        innerCommand.Connection = conn.ObtenerConexion();
                        innerCommand.CommandText = "SELECT DeudaTotal FROM Expedientes WHERE RefCliente='" + referencia + "'";

                        importe = innerCommand.ExecuteScalar()?.ToString();
                    }

                    if (!string.IsNullOrEmpty(importe))
                    {
                        importe = importe.Replace(".", ",");
                        importe2Decimales = importe.Substring(0, importe.IndexOf(',') + 3);
                    }

                    string nombreCompleto = reader["FullName_reemplazada"].ToString();
                    string nombreCompleto2 = reader["FullName2_reemplazada"].ToString();
                    string valorNombreCompleto = string.IsNullOrEmpty(nombreCompleto2) ? nombreCompleto : nombreCompleto2;
                    string direccion = reader["Address1_reemplazada"].ToString();
                    int cp = Convert.ToInt32(reader["Zip"]);
                    string localidad = reader["City_reemplazada"].ToString();
                    string provincia = CodigoPostal(cp);

                    string[] palabrasNombre = valorNombreCompleto.ToLower().Split(' ');
                    for (int i = 0; i < palabrasNombre.Length; i++) if (palabrasNombre[i].Length > 2) palabrasNombre[i] = char.ToUpper(palabrasNombre[i][0]) + palabrasNombre[i].Substring(1);
                    string nombreFormateado = string.Join(" ", palabrasNombre);

                    string[] palabrasDireccion = direccion.ToLower().Split(' ');
                    for (int i = 0; i < palabrasDireccion.Length; i++) if (palabrasDireccion[i].Length > 2) palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                    string direccionFormateada = string.Join(" ", palabrasDireccion);

                    string[] palabrasLocalidad = localidad.ToLower().Split(' ');
                    for (int i = 0; i < palabrasLocalidad.Length; i++) if (palabrasLocalidad[i].Length > 2) palabrasLocalidad[i] = char.ToUpper(palabrasLocalidad[i][0]) + palabrasLocalidad[i].Substring(1);
                    string localidadFormateada = string.Join(" ", palabrasLocalidad);

                    var exportarPDF = "";
                    if (cp > 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarra, "Hello_" + referencia + ".pdf");//Creacion del destino con su nombre y extension                    
                    else if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValencia, "Hello_" + referencia + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneral, "Hello_" + referencia + ".pdf");//Creacion del destino con su nombre y extension                                

                    using (var writter = new PdfWriter(exportarPDF))
                    {
                        using (var pdf = new PdfDocument(writter))
                        {
                            var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                            doc.SetMargins(65, 70, 110, 70); //Margenes PDF

                            //Definimos Tipografia
                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            //-------------------------------------------------------------

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));//Recogemos la ruta del archivo
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);//Creacion de la imagen
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)//Utilizamos la imagen
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(70);
                            //Agregamos la imagen al PDF
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);
                            //
                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)//Creacion de la imagen
                                .SetFixedPosition(1, 380, 725)
                                .SetMaxWidth(160);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            //-------------------------------------------------------------                    

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                .SetRelativePosition(0, 40, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MC2 Legal S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -28, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, 29 de Noviembre de 2023");
                            doc.Add(prfDatosCliente);

                            //-------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            prfPrimero.Add("Referencia del crédito: ");
                            prfPrimero.Add(new Text(referencia + "\r\n\r\n").SetFont(boldFont));

                            prfPrimero.Add("Muy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 20 de noviembre de 2023, ");

                            prfPrimero.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cedente").SetFont(boldFont));
                            prfPrimero.Add("“) cedió a ");
                            prfPrimero.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cesionario").SetFont(boldFont));
                            prfPrimero.Add("”) una cartera de créditos y, entre ellos, el crédito de referencia ");
                            prfPrimero.Add(new Text(referencia).SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Crédito").SetFont(boldFont));
                            prfPrimero.Add("“), que ostenta frente a usted en su calidad de Titular, con un saldo pendiente a fecha de ");
                            prfPrimero.Add(new Text("20 de noviembre de 2023").SetFont(boldFont));
                            prfPrimero.Add(" de ");
                            prfPrimero.Add(new Text(importe2Decimales + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text("Pagantis").SetFont(boldFont));
                            prfPrimero.Add(".\r\n\r\n");

                            prfPrimero.Add("El cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MC2 Legal S.L").SetFont(boldFont));
                            prfPrimero.Add(". Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n"));
                            doc.Add(prfPrimero);

                            //-------------------------------------------------------------

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, 10, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 233, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            prfRecuadroInteriorIzq.Add("REFERENCIA DEL PAGO:\r\n");
                            prfRecuadroInteriorIzq.Add("IMPORTE:\r\n");
                            prfRecuadroInteriorIzq.Add("CUENTA DE PAGO:");
                            doc.Add(prfRecuadroInteriorIzq);

                            Paragraph prfRecuadroInteriorDch = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(330, 233, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfRecuadroInteriorDch.Add(new Text(referencia + "\r\n").SetFont(boldFont));
                            prfRecuadroInteriorDch.Add(new Text(importe2Decimales + " €\r\n").SetFont(boldFont));
                            prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                            doc.Add(prfRecuadroInteriorDch);

                            //-------------------------------------------------------------

                            Paragraph prfSegundo = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 30, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfSegundo.Add("También le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a ");
                            prfSegundo.Add(new Text("MC2 Legal S.L").SetFont(boldFont));
                            prfSegundo.Add(". en el teléfono ");
                            prfSegundo.Add(new Text("911175438").SetFont(boldFont));
                            prfSegundo.Add(" o en la dirección de correo electrónico ");

                            string link1 = "arbor@prejudicial.es";
                            PdfLinkAnnotation linkAnnotation1 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation1.SetAction(PdfAction.CreateURI("mailto:" + link1));
                            linkAnnotation1.SetBorder(new PdfArray());
                            Link linkMail = new Link(link1, linkAnnotation1);
                            prfSegundo.Add(linkMail.SetFontColor(colorAzul).SetUnderline());

                            prfSegundo.Add(".\r\n\r\nIgualmente, estaremos encantados de poder atenderle en nuestro número de teléfono arriba indicado, donde le podremos ofrecer diferentes formas de pago para facilitar la regularización de su deuda.");
                            doc.Add(prfSegundo);

                            //-------------------------------------------------------------

                            Paragraph prfTercero = new Paragraph()
                                .SetPageNumber(2)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -15, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            prfTercero.Add("Por último, ");
                            prfTercero.Add(new Text("AKCP Europe SCSp (“Arbor Knot”)").SetFont(boldFont));
                            prfTercero.Add(" le informa de que sus datos de carácter personal estarán sujetos a su política de protección de datos que puede consultar en la siguiente dirección web: ");

                            string link2 = "https://arborknot.com/privacy-policy/";
                            PdfLinkAnnotation linkAnnotation2 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation2.SetAction(PdfAction.CreateURI(link2));
                            linkAnnotation2.SetBorder(new PdfArray());
                            Link linkHtml1 = new Link(link2, linkAnnotation2);
                            prfTercero.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());

                            prfTercero.Add(". Puede encontrar un breve resumen sobre la cesión de sus datos, así como su tratamiento por Arbor Knot en el pie de la presente comunicación. Tenga en cuenta que algunos de los tratamientos descritos en la citada política pueden estar sujetos a su previa conformidad en el momento en que usted facilite voluntariamente datos de carácter personal adicionales para valorar alternativas que le permitan mejorar su capacidad financiera.");

                            prfTercero.Add("\r\n\r\nSin otro particular, aprovechamos la ocasión para saludarle atentamente.");

                            prfTercero.Add(new Text("\r\n\r\n\r\nPagamastarde S.L.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaPagamastarde = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Firma_Pagamasterde.jpg"));
                            ImageData imagenDataFirmaPagamastarde = ImageDataFactory.Create(rutaFirmaPagamastarde);
                            var firmaPagamastarde = new iText.Layout.Element.Image(imagenDataFirmaPagamastarde)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 10, 1, 1)
                                .SetMaxWidth(160);
                            Paragraph firmaIzq = new Paragraph("");
                            firmaIzq.Add(firmaPagamastarde);
                            doc.Add(firmaIzq);

                            string rutaFirmaAKCP = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Firma_AKCP_EUROPE_SCSP.jpg"));
                            ImageData imagenDataFirmaAKCP = ImageDataFactory.Create(rutaFirmaAKCP);
                            var firmaAKCP = new iText.Layout.Element.Image(imagenDataFirmaAKCP)
                                .SetPageNumber(2)
                                .SetRelativePosition(320, -45, 1, 1)
                                .SetMaxWidth(160);
                            Paragraph firmaDch = new Paragraph("");
                            firmaDch.Add(firmaAKCP);
                            doc.Add(firmaDch);

                            //-------------------------------------------------------------

                            if (cp > 31000 && cp < 32000) //Hello_747039
                            {
                                Paragraph prfNavarra = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -20, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);

                                prfNavarra.Add("Asimismo, ");
                                prfNavarra.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de doscientos mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);

                            }
                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfValencia = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -20, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(10)
                                    .SetFixedLeading(11);

                                prfValencia.Add("Asimismo, ");
                                prfValencia.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(" de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                                prfValencia.Add(" está domiciliada en Cornella de Llobregar (08940) Plaza Pau, S/N EDF 3, P.3, WTC ALMEDA PARK, constituido el 11 de Abril de 2.011 e Inscrita en el Registro Mercantil de Barcelona, al Tomo 42.465, Folio 94, Sección 8ª, Hoja B-409105, ii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de doscientos mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -10, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(7)
                            .SetFixedLeading(8);

                            prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                            string link3 = "www.aepd.es";
                            PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation3.SetAction(PdfAction.CreateURI(link2));
                            linkAnnotation3.SetBorder(new PdfArray());
                            Link linkHtml2 = new Link(link3, linkAnnotation3);
                            prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                            prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                            prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                            prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                            doc.Add(prfCuarto);

                            //-------------------------------------------------------------

                            string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Proteccion_Datos.png"));
                            ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                            var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfTabla = new Paragraph();
                            prfTabla.Add(Tabla);
                            doc.Add(prfTabla);
                        }
                    }
                }
                reader.Close();
            }
            ConexionDB.CerrarConexion();
        }
    }
    //PAGANTIS
    /*public partial class Form1 : Form
    {
        ConexionDB conn = new ConexionDB();
        MCCommand mcComm = new MCCommand();

        //Ruta para guardar archivo
        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public Form1()
        {
            InitializeComponent();
            CenterToScreen();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CrearPDF();
        }

        public string CodigoPostal(int cp)
        {
            string provincia = "";
            if (cp < 1999) return provincia = "ÁLAVA";
            else if (cp < 2999) return provincia = "ALBACETE";
            else if (cp < 3999) return provincia = "ALICANTE";
            else if (cp < 4999) return provincia = "ALMERÍA";
            else if (cp < 5999) return provincia = "ÁVILA";
            else if (cp < 6999) return provincia = "BADAJOZ";
            else if (cp < 7999) return provincia = "BALEARES";
            else if (cp < 8999) return provincia = "BARCELONA";
            else if (cp < 9999) return provincia = "BURGOS";
            else if (cp < 10999) return provincia = "CÁCERES";
            else if (cp < 11999) return provincia = "CÁDIZ";
            else if (cp < 12999) return provincia = "CASTELLÓN";
            else if (cp < 13999) return provincia = "CIUDAD REAL";
            else if (cp < 14999) return provincia = "CÓRDOBA";
            else if (cp < 15999) return provincia = "LA CORUÑA";
            else if (cp < 16999) return provincia = "CUENCA";
            else if (cp < 17999) return provincia = "GERONA";
            else if (cp < 18999) return provincia = "GRANADA";
            else if (cp < 19999) return provincia = "GUADALAJARA";
            else if (cp < 20999) return provincia = "GUIPÚZCOA";
            else if (cp < 21999) return provincia = "HUELVA";
            else if (cp < 22999) return provincia = "HUESCA";
            else if (cp < 23999) return provincia = "JAÉN";
            else if (cp < 24999) return provincia = "LEÓN";
            else if (cp < 25999) return provincia = "LÉRIDA";
            else if (cp < 26999) return provincia = "LA RIOJA";
            else if (cp < 27999) return provincia = "LUGO";
            else if (cp < 28999) return provincia = "MADRID";
            else if (cp < 29999) return provincia = "MÁLAGA";
            else if (cp < 30999) return provincia = "MURCIA";
            else if (cp < 31999) return provincia = "NAVARRA";
            else if (cp < 32999) return provincia = "OURENSE";
            else if (cp < 33999) return provincia = "ASTURIAS";
            else if (cp < 34999) return provincia = "PALENCIA";
            else if (cp < 35999) return provincia = "LAS PALMAS";
            else if (cp < 36999) return provincia = "PONTEVEDRA";
            else if (cp < 37999) return provincia = "SALAMANCA";
            else if (cp < 38999) return provincia = "S. C. DE TENERIFE";
            else if (cp < 39999) return provincia = "CANTABRIA";
            else if (cp < 40999) return provincia = "SEGOVIA";
            else if (cp < 41999) return provincia = "SEVILLA";
            else if (cp < 42999) return provincia = "SORIA";
            else if (cp < 43999) return provincia = "TARRAGONA";
            else if (cp < 44999) return provincia = "TERUEL";
            else if (cp < 45999) return provincia = "TOLEDO";
            else if (cp < 46999) return provincia = "VALENCIA";
            else if (cp < 47999) return provincia = "VALLADOLID";
            else if (cp < 48999) return provincia = "VIZCAYA";
            else if (cp < 49999) return provincia = "ZAMORA";
            else if (cp < 50999) return provincia = "ZARAGOZA";
            else if (cp < 51999) return provincia = "CEUTA";
            else if (cp < 52999) return provincia = "MELILLA";
            else return provincia;
        }

        private void CrearPDF()
        {
            ConexionDB.AbrirConexion();

            string rutaArchivos = Path.Combine(ruta, "PagantisPDF");
            string rutaArchivosGeneral = Path.Combine(ruta, "PagantisPDF/General");
            string rutaArchivosNavarra = Path.Combine(ruta, "PagantisPDF/Navarra");
            string rutaArchivosValencia = Path.Combine(ruta, "PagantisPDF/Valencia");
            if (!Directory.Exists(rutaArchivos)) Directory.CreateDirectory(rutaArchivos);
            if (!Directory.Exists(rutaArchivosGeneral)) Directory.CreateDirectory(rutaArchivosGeneral);
            if (!Directory.Exists(rutaArchivosNavarra)) Directory.CreateDirectory(rutaArchivosNavarra);
            if (!Directory.Exists(rutaArchivosValencia)) Directory.CreateDirectory(rutaArchivosValencia);

            mcComm.command.Connection = conn.ObtenerConexion();
            mcComm.CommandText = "SELECT OriginalAccountId,  REPLACE(FullName, '?', '') AS FullName_reemplazada, REPLACE(FullName2, '?', '') AS FullName2_reemplazada, REPLACE(Address1, '?', '') AS Address1_reemplazada, Zip, REPLACE(City, '?', '') AS City_reemplazada FROM TempPagantis";

            using (IDataReader reader = mcComm.ExecuteReader())
            {
                while (reader.Read())
                {
                    string referencia = reader["OriginalAccountId"].ToString();
                    string importe = string.Empty;
                    string importe2Decimales =  string.Empty;

                    using (SqlCommand innerCommand = new SqlCommand())
                    {
                        innerCommand.Connection = conn.ObtenerConexion();
                        innerCommand.CommandText = "SELECT DeudaTotal FROM Expedientes WHERE RefCliente='" + referencia + "'";

                        importe = innerCommand.ExecuteScalar()?.ToString();
                    }

                    if (!string.IsNullOrEmpty(importe))
                    {
                        importe = importe.Replace(".", ",");
                        importe2Decimales = importe.Substring(0, importe.IndexOf(',') + 3);
                    }

                    string nombreCompleto = reader["FullName_reemplazada"].ToString();
                    string nombreCompleto2 = reader["FullName2_reemplazada"].ToString();
                    string valorNombreCompleto = string.IsNullOrEmpty(nombreCompleto2) ? nombreCompleto : nombreCompleto2;
                    string direccion = reader["Address1_reemplazada"].ToString();
                    int cp = Convert.ToInt32(reader["Zip"]);
                    string localidad = reader["City_reemplazada"].ToString();
                    string provincia = CodigoPostal(cp);

                    string[] palabrasNombre = valorNombreCompleto.ToLower().Split(' ');
                    for (int i = 0; i < palabrasNombre.Length; i++) if (palabrasNombre[i].Length > 2) palabrasNombre[i] = char.ToUpper(palabrasNombre[i][0]) + palabrasNombre[i].Substring(1);
                    string nombreFormateado = string.Join(" ", palabrasNombre);

                    string[] palabrasDireccion = direccion.ToLower().Split(' ');
                    for (int i = 0; i < palabrasDireccion.Length; i++) if (palabrasDireccion[i].Length > 2) palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                    string direccionFormateada = string.Join(" ", palabrasDireccion);

                    string[] palabrasLocalidad = localidad.ToLower().Split(' ');
                    for (int i = 0; i < palabrasLocalidad.Length; i++) if (palabrasLocalidad[i].Length > 2) palabrasLocalidad[i] = char.ToUpper(palabrasLocalidad[i][0]) + palabrasLocalidad[i].Substring(1);
                    string localidadFormateada = string.Join(" ", palabrasLocalidad);

                    var exportarPDF = "";
                    if (cp > 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarra, "Hello_" + referencia + ".pdf");//Creacion del destino con su nombre y extension                    
                    else if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValencia, "Hello_" + referencia + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneral, "Hello_" + referencia + ".pdf");//Creacion del destino con su nombre y extension                                

                    using (var writter = new PdfWriter(exportarPDF))
                    {
                        using (var pdf = new PdfDocument(writter))
                        {
                            var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                            doc.SetMargins(65, 70, 110, 70); //Margenes PDF

                            //Definimos Tipografia
                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            //-------------------------------------------------------------

                            string rutaLogoArborknot = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Logo_Arborknot.jpg"));//Recogemos la ruta del archivo
                            ImageData imagenDataArborknot = ImageDataFactory.Create(rutaLogoArborknot);//Creacion de la imagen
                            var logoArborknot = new iText.Layout.Element.Image(imagenDataArborknot)//Utilizamos la imagen
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(70);
                            //Agregamos la imagen al PDF
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoArborknot);
                            doc.Add(encabezadoIzq);
                            //
                            string rutaLogoPagantis = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Logo_Pagantis.jpg"));
                            ImageData imagenDataPagantis = ImageDataFactory.Create(rutaLogoPagantis);
                            var logoPagantis = new iText.Layout.Element.Image(imagenDataPagantis)//Creacion de la imagen
                                .SetFixedPosition(1, 380, 725)
                                .SetMaxWidth(160);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoPagantis);
                            doc.Add(encabezadoDch);

                            //-------------------------------------------------------------                    

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                .SetRelativePosition(0, 40, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MC2 Legal S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -28, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, 29 de Noviembre de 2023");
                            doc.Add(prfDatosCliente);

                            //-------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            prfPrimero.Add("Referencia del crédito: ");
                            prfPrimero.Add(new Text(referencia + "\r\n\r\n").SetFont(boldFont));

                            prfPrimero.Add("Muy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 20 de noviembre de 2023, ");

                            prfPrimero.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cedente").SetFont(boldFont));
                            prfPrimero.Add("“) cedió a ");
                            prfPrimero.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cesionario").SetFont(boldFont));
                            prfPrimero.Add("”) una cartera de créditos y, entre ellos, el crédito de referencia ");
                            prfPrimero.Add(new Text(referencia).SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Crédito").SetFont(boldFont));
                            prfPrimero.Add("“), que ostenta frente a usted en su calidad de Titular, con un saldo pendiente a fecha de ");
                            prfPrimero.Add(new Text("20 de noviembre de 2023").SetFont(boldFont));
                            prfPrimero.Add(" de ");
                            prfPrimero.Add(new Text(importe2Decimales + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text("Pagantis").SetFont(boldFont));
                            prfPrimero.Add(".\r\n\r\n");

                            prfPrimero.Add("El cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MC2 Legal S.L").SetFont(boldFont));
                            prfPrimero.Add(". Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n"));
                            doc.Add(prfPrimero);

                            //-------------------------------------------------------------

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, 10, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 233, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            prfRecuadroInteriorIzq.Add("REFERENCIA DEL PAGO:\r\n");
                            prfRecuadroInteriorIzq.Add("IMPORTE:\r\n");
                            prfRecuadroInteriorIzq.Add("CUENTA DE PAGO:");
                            doc.Add(prfRecuadroInteriorIzq);

                            Paragraph prfRecuadroInteriorDch = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(330, 233, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfRecuadroInteriorDch.Add(new Text(referencia + "\r\n").SetFont(boldFont));
                            prfRecuadroInteriorDch.Add(new Text(importe2Decimales + " €\r\n").SetFont(boldFont));
                            prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                            doc.Add(prfRecuadroInteriorDch);

                            //-------------------------------------------------------------

                            Paragraph prfSegundo = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 30, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfSegundo.Add("También le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a ");
                            prfSegundo.Add(new Text("MC2 Legal S.L").SetFont(boldFont));
                            prfSegundo.Add(". en el teléfono ");
                            prfSegundo.Add(new Text("911175438").SetFont(boldFont));
                            prfSegundo.Add(" o en la dirección de correo electrónico ");

                            string link1 = "arbor@prejudicial.es";
                            PdfLinkAnnotation linkAnnotation1 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation1.SetAction(PdfAction.CreateURI("mailto:" + link1));
                            linkAnnotation1.SetBorder(new PdfArray());
                            Link linkMail = new Link(link1, linkAnnotation1);
                            prfSegundo.Add(linkMail.SetFontColor(colorAzul).SetUnderline());

                            prfSegundo.Add(".\r\n\r\nIgualmente, estaremos encantados de poder atenderle en nuestro número de teléfono arriba indicado, donde le podremos ofrecer diferentes formas de pago para facilitar la regularización de su deuda.");
                            doc.Add(prfSegundo);

                            //-------------------------------------------------------------

                            Paragraph prfTercero = new Paragraph()
                                .SetPageNumber(2)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -15, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            prfTercero.Add("Por último, ");
                            prfTercero.Add(new Text("AKCP Europe SCSp (“Arbor Knot”)").SetFont(boldFont));
                            prfTercero.Add(" le informa de que sus datos de carácter personal estarán sujetos a su política de protección de datos que puede consultar en la siguiente dirección web: ");

                            string link2 = "https://arborknot.com/privacy-policy/";
                            PdfLinkAnnotation linkAnnotation2 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation2.SetAction(PdfAction.CreateURI(link2));
                            linkAnnotation2.SetBorder(new PdfArray());
                            Link linkHtml1 = new Link(link2, linkAnnotation2);
                            prfTercero.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());

                            prfTercero.Add(". Puede encontrar un breve resumen sobre la cesión de sus datos, así como su tratamiento por Arbor Knot en el pie de la presente comunicación. Tenga en cuenta que algunos de los tratamientos descritos en la citada política pueden estar sujetos a su previa conformidad en el momento en que usted facilite voluntariamente datos de carácter personal adicionales para valorar alternativas que le permitan mejorar su capacidad financiera.");

                            prfTercero.Add("\r\n\r\nSin otro particular, aprovechamos la ocasión para saludarle atentamente.");

                            prfTercero.Add(new Text("\r\n\r\n\r\nPagamastarde S.L.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaPagamastarde = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Firma_Pagamasterde.jpg"));
                            ImageData imagenDataFirmaPagamastarde = ImageDataFactory.Create(rutaFirmaPagamastarde);
                            var firmaPagamastarde = new iText.Layout.Element.Image(imagenDataFirmaPagamastarde)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 10, 1, 1)
                                .SetMaxWidth(160);
                            Paragraph firmaIzq = new Paragraph("");
                            firmaIzq.Add(firmaPagamastarde);
                            doc.Add(firmaIzq);

                            string rutaFirmaAKCP = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Firma_AKCP_EUROPE_SCSP.jpg"));
                            ImageData imagenDataFirmaAKCP = ImageDataFactory.Create(rutaFirmaAKCP);
                            var firmaAKCP = new iText.Layout.Element.Image(imagenDataFirmaAKCP)
                                .SetPageNumber(2)
                                .SetRelativePosition(320, -45, 1, 1)
                                .SetMaxWidth(160);
                            Paragraph firmaDch = new Paragraph("");
                            firmaDch.Add(firmaAKCP);
                            doc.Add(firmaDch);

                            //-------------------------------------------------------------

                            if (cp > 31000 && cp < 32000) //Hello_747039
                            {
                                Paragraph prfNavarra = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -20, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);

                                prfNavarra.Add("Asimismo, ");
                                prfNavarra.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de doscientos mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);

                            }                           
                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfValencia = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -20, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(10)
                                    .SetFixedLeading(11);

                                prfValencia.Add("Asimismo, ");
                                prfValencia.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(" de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                                prfValencia.Add(" está domiciliada en Cornella de Llobregar (08940) Plaza Pau, S/N EDF 3, P.3, WTC ALMEDA PARK, constituido el 11 de Abril de 2.011 e Inscrita en el Registro Mercantil de Barcelona, al Tomo 42.465, Folio 94, Sección 8ª, Hoja B-409105, ii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de doscientos mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -10, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(7)
                            .SetFixedLeading(8);

                            prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                            string link3 = "www.aepd.es";
                            PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation3.SetAction(PdfAction.CreateURI(link2));
                            linkAnnotation3.SetBorder(new PdfArray());
                            Link linkHtml2 = new Link(link3, linkAnnotation3);
                            prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                            prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                            prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                            prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                            doc.Add(prfCuarto);

                            //-------------------------------------------------------------

                            string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "imagenes/Proteccion_Datos.png"));
                            ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                            var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfTabla = new Paragraph();
                            prfTabla.Add(Tabla);
                            doc.Add(prfTabla);
                        }
                    }
                }
                reader.Close();
            }
            ConexionDB.CerrarConexion();
        }
    }*/
}

