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
using iText.StyledXmlParser.Jsoup.Select;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace Creacion_PDF_HelloLetter
{
    //"" - Quartz Capital Fund II - ORANGE
    public partial class Form1 : Form
    {
        ConexionDB conn = new ConexionDB();
        MCCommand mcComm = new MCCommand();
        Comp comp = new Comp();
        private OpenFileDialog openFileDialog;

        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public Form1()
        {
            InitializeComponent();
            CenterToScreen();
            openFileDialog = new OpenFileDialog();
        }

        private void btnFichero_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK) txtFichero.Text = openFileDialog.FileName;
            else MessageBox.Show("Debe seleccionar un fichero para cargar los expedientes"); return;
        }

        private void button1_Click(object sender, EventArgs e) { CargarPDF(); }

        private void CargarPDF()
        {
            string hoja = nHoja();
            int count = 0;

            OleDbConnection oleConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim() + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter("SELECT * FROM [" + hoja + "$]", oleConnection);

            DataSet ds = new DataSet();
            oleAdapter.Fill(ds);
            oleConnection.Close();

            DataTable dt = ds.Tables[0];

            foreach (DataRow fila in dt.Rows)
            {
                string expediente = fila["EXPEDIENTE"].ToString();
                string contrato = fila["CUST_CODE"].ToString();
                string refEnvio = fila["REF_ENVIO"].ToString();
                string importe = fila["PENDIENTE"].ToString();
                string nombre = fila["NOMBRE"].ToString();
                string municipio = fila["MUNICIPIO"] != DBNull.Value ? fila["MUNICIPIO"].ToString() : "";
                string direccion = fila["DIRECCION"].ToString();
                int cp = fila["CP"] != DBNull.Value ? Convert.ToInt32(fila["CP"]) : 0;

                string[] nombreMinusculas = nombre.ToLower().Split(' ');
                for (int i = 0; i < nombreMinusculas.Length; i++) if (nombreMinusculas[i].Length > 2) nombreMinusculas[i] = char.ToUpper(nombreMinusculas[i][0]) + nombreMinusculas[i].Substring(1);
                string nombreFormateado = string.Join(" ", nombreMinusculas);

                string[] palabrasDireccion = direccion.ToLower().Split(' ');
                for (int i = 0; i < palabrasDireccion.Length; i++)
                {
                    if (palabrasDireccion[i].Length > 2)
                    {
                        palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                        if (palabrasDireccion[i] == "null" || palabrasDireccion[i] == "(null)") palabrasDireccion[i] = string.Empty;
                    }
                }
                string direccionFormateada = string.Join(" ", palabrasDireccion);

                string[] palabrasLocalidad = municipio.ToLower().Split(' ');
                for (int i = 0; i < palabrasLocalidad.Length; i++) if (palabrasLocalidad[i].Length > 2) palabrasLocalidad[i] = char.ToUpper(palabrasLocalidad[i][0]) + palabrasLocalidad[i].Substring(1);
                string localidadFormateada = string.Join(" ", palabrasLocalidad);

                string provincia = comp.CodigoPostal(cp);

                string rutaArchivos = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Quartz Capital Fund II");
                string rutaArchivosGeneral = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Quartz Capital Fund II/General");
                string rutaArchivosNavarra = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Quartz Capital Fund II/Navarra");
                string rutaArchivosValencia = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Quartz Capital Fund II/Valencia");
                if (!Directory.Exists(rutaArchivos)) Directory.CreateDirectory(rutaArchivos);
                if (!Directory.Exists(rutaArchivosGeneral)) Directory.CreateDirectory(rutaArchivosGeneral);
                if (!Directory.Exists(rutaArchivosNavarra)) Directory.CreateDirectory(rutaArchivosNavarra);
                if (!Directory.Exists(rutaArchivosValencia)) Directory.CreateDirectory(rutaArchivosValencia);

                var exportarPDF = "";
                if (cp > 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarra, "Hello_" + refEnvio + ".pdf");
                else if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValencia, "Hello_" + refEnvio + ".pdf");
                else exportarPDF = Path.Combine(rutaArchivosGeneral, "Hello_" + refEnvio + ".pdf");
                
                using (var writter = new PdfWriter(exportarPDF))
                {
                    using (var pdf = new PdfDocument(writter))
                    {
                        var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);
                        doc.SetMargins(50, 70, 50, 70);
                        
                        string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                        string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                        PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                        PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                        iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                        iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                        ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);
                        var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)
                            .SetRelativePosition(0, 0, 0, 0)
                            .SetMaxWidth(70)
                            .SetMarginBottom(48);                        
                        Paragraph encabezadoIzq = new Paragraph("");
                        encabezadoIzq.Add(logoIzq);
                        doc.Add(encabezadoIzq);                        
                        string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Quartz.jpg"));
                        ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                        var logoDch = new iText.Layout.Element.Image(imagenDataDch)
                            .SetFixedPosition(1, 370, 745)
                            .SetMaxWidth(160);
                        Paragraph encabezadoDch = new Paragraph("");
                        encabezadoDch.Add(logoDch);
                        doc.Add(encabezadoDch);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////                    

                        Paragraph prfMC2 = new Paragraph()
                            .SetPageNumber(1)
                            .SetRelativePosition(0, -20, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfMC2.Add("MCDOS LEGAL S.L.");
                        prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                        doc.Add(prfMC2);

                        Paragraph prfDatosCliente = new Paragraph()
                            .SetPageNumber(1)
                            .SetTextAlignment(TextAlignment.RIGHT)
                            .SetFixedPosition( 25, 580
                            , 500)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\n\r\n\r\nMadrid, 21 de febrero de 2024");
                        doc.Add(prfDatosCliente);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        Paragraph prfPrimero = new Paragraph()
                            .SetPageNumber(1)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, 0, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);

                        prfPrimero.Add("\r\nReferencia del crédito: ");
                        prfPrimero.Add(new Text(contrato).SetFont(boldFont));                        

                        prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\n\r\nPor la presente le comunicamos que con fecha 19 de febrero de 2024, ");
                        prfPrimero.Add(new Text("QUARTZ CAPITAL FUND II").SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Cedente").SetFont(boldFont));
                        prfPrimero.Add("“) cedió a ");
                        prfPrimero.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Cesionario").SetFont(boldFont));
                        prfPrimero.Add("”) una cartera de créditos y, entre ellos, el crédito de referencia ");
                        prfPrimero.Add(new Text(contrato).SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Crédito").SetFont(boldFont));
                        prfPrimero.Add("“), que ostenta frente a usted en su calidad de Titular, con un saldo pendiente a fecha de 19 de febrero de 2024 de ");
                        prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                        prfPrimero.Add(", cuyo origen es ");
                        prfPrimero.Add(new Text("Orange.").SetFont(boldFont));

                        prfPrimero.Add("\r\n\r\nEl Cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                        prfPrimero.Add(new Text("MCDOS LEGAL S.L.").SetFont(boldFont));
                        prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                        prfPrimero.Add("Por la presente, le requerimos para que - ");
                        prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                        prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                        prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont));
                        doc.Add(prfPrimero);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                        ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                        var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                            .SetRelativePosition(0, 0, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfRecuadro = new Paragraph();
                        prfRecuadro.Add(Recuadro);
                        doc.Add(prfRecuadro);

                        Paragraph prfRecuadroInteriorIzq = new Paragraph()
                            .SetPageNumber(1)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetFixedPosition(135, 266, 250)
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
                            .SetFixedPosition(330, 266, 250)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfRecuadroInteriorDch.Add(new Text(expediente + "\r\n").SetFont(boldFont));
                        prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                        prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                        doc.Add(prfRecuadroInteriorDch);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        Paragraph prfSegundo = new Paragraph()
                            .SetPageNumber(1)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, 0, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfSegundo.Add("\r\nTambién le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a la Agencia de Cobro en el teléfono ");
                        prfSegundo.Add(new Text("91 108 89 04").SetFont(boldFont));
                        prfSegundo.Add(" o en la dirección de correo electrónico ");

                        string link1 = "arbor@prejudicial.es";
                        PdfLinkAnnotation linkAnnotation1 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                        linkAnnotation1.SetAction(PdfAction.CreateURI("mailto:" + link1));
                        linkAnnotation1.SetBorder(new PdfArray());
                        Link linkMail = new Link(link1, linkAnnotation1);
                        prfSegundo.Add(linkMail.SetFontColor(colorAzul).SetUnderline());

                        prfSegundo.Add(".\r\n\r\nIgualmente, estaremos encantados de poder atenderle en nuestro número de teléfono arriba indicado, donde le podremos ofrecer diferentes formas de pago para facilitar la regularización de su deuda.");
                        doc.Add(prfSegundo);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        Paragraph prfTercero = new Paragraph()
                            .SetPageNumber(1)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, 0, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);

                        prfTercero.Add("\r\nPor último, ");
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

                        prfTercero.Add(new Text("\r\n\r\n\r\nQUARTZ CAPITAL FUND II\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                        doc.Add(prfTercero);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                        ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                        var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                            .SetPageNumber(2)
                            .SetRelativePosition(0, 10, 1, 1)
                            .SetMaxWidth(150);
                        Paragraph firmaPrfIzq = new Paragraph("");
                        firmaPrfIzq.Add(firmaIzq);
                        doc.Add(firmaPrfIzq);

                        string rutaFirmaDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                        ImageData imagenDataFirmaDch = ImageDataFactory.Create(rutaFirmaDch);
                        var firmaDch = new iText.Layout.Element.Image(imagenDataFirmaDch)
                            .SetPageNumber(2)
                            .SetRelativePosition(285, -45, 1, 1)
                            .SetMaxWidth(180);
                        Paragraph firmaPrfDch = new Paragraph("");
                        firmaPrfDch.Add(firmaDch);
                        doc.Add(firmaPrfDch);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        if (cp > 31000 && cp < 32000) //NAVARRA
                        {
                            Paragraph prfNavarra = new Paragraph()
                                .SetPageNumber(2)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            prfNavarra.Add("Asimismo, ");
                            prfNavarra.Add(new Text("QUARTZ CAPITAL FUND II").SetFont(boldFont));
                            prfNavarra.Add(" y ");
                            prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfNavarra.Add(", de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de trescientos quince mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
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
                            prfValencia.Add(new Text("QUARTZ CAPITAL FUND II").SetFont(boldFont));
                            prfValencia.Add(" cedio los creditos a ");
                            prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) QUARTZ CAPITAL FUND S.C.A. – QUARTZ CAPITAL II, una Sociedad de Inversión de Capital Variable, Fondo de Inversión Especializado organizado bajo las leyes del Gran Ducado de Luxemburgo en forma de una sociedad en comandita por acciones, con sede social en 6A Rue Gabriel Lippman, L-5365 Schuttrange-Munsbach, Gran Ducado de Luxemburgo, registrada en el Registro de Comercio y Sociedades de Luxemburgo (Registre de Commerce et des Sociétés o RCS) con el número 167191, representada por su Socio General QUARTZ MANAGEMENT GP S.A.R.L, una sociedad de responsabilidad limitada privada de Luxemburgo, con domicilio social en 16 Rue d’Epernay L-1616 Luxemburgo, Gran Ducado de Luxemburgo, registrada en el Registro de Comercio y Sociedades de Luxemburgo con el número B 211.727 y titular del número de identificación fiscal español N0076022C ii) el Crédito se encuentra identificado en la página 21 a 41 del Contrato de Cesión de la Cartera de Créditos (Anexo III) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de trescientos quince mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");                            
                            doc.Add(prfValencia);
                        }

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        Paragraph prfCuarto = new Paragraph()
                        .SetPageNumber(2)
                        .SetVerticalAlignment(VerticalAlignment.TOP)
                        .SetTextAlignment(TextAlignment.JUSTIFIED)
                        .SetRelativePosition(0, -15, 0, 0)
                        .SetFontColor(colorNegro)
                        .SetFont(regularFont)
                        .SetFontSize(7)
                        .SetFixedLeading(8);

                        prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                        string link3 = "www.aepd.es";
                        PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                        linkAnnotation3.SetAction(PdfAction.CreateURI(link3));
                        linkAnnotation3.SetBorder(new PdfArray());
                        Link linkHtml2 = new Link(link3, linkAnnotation3);
                        prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                        prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                        prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                        prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                        doc.Add(prfCuarto);

                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Proteccion_Datos.png"));
                        ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                        var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                            .SetPageNumber(2)
                            .SetRelativePosition(0, -10, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfTabla = new Paragraph();
                        prfTabla.Add(Tabla);
                        doc.Add(prfTabla);
                    }
                }                
                count++;
            }
            MessageBox.Show("Se han generado " + count + " pdfs.");
        }      

        private string nHoja()
        {
            string hoja;
            var xlsApp = new Excel.Application();
            xlsApp.Workbooks.Open(txtFichero.Text.Trim());

            hoja = xlsApp.Sheets[1].Name;

            xlsApp.DisplayAlerts = false;
            xlsApp.Workbooks.Close();
            xlsApp.DisplayAlerts = true;

            xlsApp.Quit();
            xlsApp = null;

            return hoja;
        }
    }
    
    //CRISALIDA III - ALERIN/AXACTOR ESPAÑA/AXACTOR INVEST - SANTANDER
    /*public partial class Form1 : Form
    {
        ConexionDB conn = new ConexionDB();
        MCCommand mcComm = new MCCommand();
        Comp comp = new Comp();
        private OpenFileDialog openFileDialog;

        //Ruta para guardar archivo
        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public Form1()
        {
            InitializeComponent();
            CenterToScreen();
            openFileDialog = new OpenFileDialog();
        }

        private void btnFichero_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFichero.Text = openFileDialog.FileName;
            }
            else
            {
                MessageBox.Show("Debe seleccionar un fichero para cargar los expedientes");
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CargarPDF();
        }        

        private void CargarPDF()
        {
            string hoja = nHoja();
            int count = 0;

            OleDbConnection oleConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim() + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter("SELECT * FROM [" + hoja + "$]", oleConnection);

            DataSet ds = new DataSet();
            oleAdapter.Fill(ds);
            oleConnection.Close();

            DataTable dt = ds.Tables[0];

            foreach (DataRow fila in dt.Rows)
            {
                //string contrato = fila["CONTRATO"].ToString();
                string refMC = fila["REF_MC"].ToString();
                string contrato = fila["CONTRATO"].ToString();
                string refEnvio = fila["REF_ENVIO"].ToString();
                string origen = fila["ORIGEN"].ToString();
                string importe = fila["IMPORTE"].ToString();
                string nombre = fila["NOMBRE"].ToString();
                string municipio = fila["MUNICIPIO"] != DBNull.Value ? fila["MUNICIPIO"].ToString() : "";
                //string municipio = fila["MUNICIPIO"].ToString();
                string direccion = fila["DIRECCION"].ToString();
                int cp = fila["CP"] != DBNull.Value ? Convert.ToInt32(fila["CP"]) : 0;
                //string provincia = fila["PROVINCIA"].ToString();

                string[] nombreMinusculas = nombre.ToLower().Split(' ');
                for (int i = 0; i < nombreMinusculas.Length; i++) if (nombreMinusculas[i].Length > 2) nombreMinusculas[i] = char.ToUpper(nombreMinusculas[i][0]) + nombreMinusculas[i].Substring(1);
                string nombreFormateado = string.Join(" ", nombreMinusculas);

                //string[] palabrasDireccion = direccion.ToLower().Split(' ');
                //for (int i = 0; i < palabrasDireccion.Length; i++) if (palabrasDireccion[i].Length > 2) palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                //string direccionFormateada = string.Join(" ", palabrasDireccion);

                string[] palabrasDireccion = direccion.ToLower().Split(' ');
                for (int i = 0; i < palabrasDireccion.Length; i++)
                {
                    if (palabrasDireccion[i].Length > 2)
                    {
                        palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                        if (palabrasDireccion[i] == "null" || palabrasDireccion[i] == "(null)") palabrasDireccion[i] = string.Empty;
                    } 
                }
                string direccionFormateada = string.Join(" ", palabrasDireccion);

                string[] palabrasLocalidad = municipio.ToLower().Split(' ');
                for (int i = 0; i < palabrasLocalidad.Length; i++) if (palabrasLocalidad[i].Length > 2) palabrasLocalidad[i] = char.ToUpper(palabrasLocalidad[i][0]) + palabrasLocalidad[i].Substring(1);
                string localidadFormateada = string.Join(" ", palabrasLocalidad);

                string provincia = comp.CodigoPostal(cp);

                string rutaArchivosA = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Alerin");
                string rutaArchivosGeneralA = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Alerin/General");
                string rutaArchivosNavarraA = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Alerin/Navarra");
                string rutaArchivosValenciaA = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Alerin/Valencia");
                if (!Directory.Exists(rutaArchivosA)) Directory.CreateDirectory(rutaArchivosA);
                if (!Directory.Exists(rutaArchivosGeneralA)) Directory.CreateDirectory(rutaArchivosGeneralA);
                if (!Directory.Exists(rutaArchivosNavarraA)) Directory.CreateDirectory(rutaArchivosNavarraA);
                if (!Directory.Exists(rutaArchivosValenciaA)) Directory.CreateDirectory(rutaArchivosValenciaA);

                string rutaArchivosE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor España");
                string rutaArchivosGeneralE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor España/General");
                string rutaArchivosNavarraE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor España/Navarra");
                string rutaArchivosValenciaE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor España/Valencia");
                if (!Directory.Exists(rutaArchivosE)) Directory.CreateDirectory(rutaArchivosE);
                if (!Directory.Exists(rutaArchivosGeneralE)) Directory.CreateDirectory(rutaArchivosGeneralE);
                if (!Directory.Exists(rutaArchivosNavarraE)) Directory.CreateDirectory(rutaArchivosNavarraE);
                if (!Directory.Exists(rutaArchivosValenciaE)) Directory.CreateDirectory(rutaArchivosValenciaE);

                string rutaArchivosI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor Invest");
                string rutaArchivosGeneralI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor Invest/General");
                string rutaArchivosNavarraI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor Invest/Navarra");
                string rutaArchivosValenciaI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor Invest/Valencia");
                if (!Directory.Exists(rutaArchivosI)) Directory.CreateDirectory(rutaArchivosA);
                if (!Directory.Exists(rutaArchivosGeneralI)) Directory.CreateDirectory(rutaArchivosGeneralI);
                if (!Directory.Exists(rutaArchivosNavarraI)) Directory.CreateDirectory(rutaArchivosNavarraI);
                if (!Directory.Exists(rutaArchivosValenciaI)) Directory.CreateDirectory(rutaArchivosValenciaI);

                var exportarPDF = "";
                if (origen == "Alerin")
                {
                    if (cp > 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraA, "Hello_" + refEnvio + ".pdf");
                    else if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaA, "Hello_" + refEnvio + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralA, "Hello_" + refEnvio + ".pdf");
                    string rutaArchivos = rutaArchivosA;
                    string rutaArchivosNavarra = rutaArchivosNavarraA;
                    string rutaArchivosValencia = rutaArchivosValenciaA;

                    using (var writter = new PdfWriter(exportarPDF))
                    {
                        using (var pdf = new PdfDocument(writter))
                        {
                            var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                            doc.SetMargins(65, 70, 30, 70); //Margenes PDF

                            //Definimos Tipografia
                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            //-------------------------------------------------------------

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Alerin.jpg"));//Recogemos la ruta del archivo
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);//Creacion de la imagen
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)//Utilizamos la imagen
                                .SetRelativePosition(0, -35, 0, 0)
                                .SetMaxWidth(70)
                                .SetMarginBottom(48);
                            //Agregamos la imagen al PDF
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);
                            //
                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)//Creacion de la imagen
                                .SetFixedPosition(1, 455, 740)
                                .SetMaxWidth(70);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            //-------------------------------------------------------------                    

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                .SetRelativePosition(0, -30, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MCdos Legal S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -110, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, 4 de enero de 2024");
                            doc.Add(prfDatosCliente);

                            //-------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -50, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                prfPrimero.Add(new Text("Referencia del crédito: " + refMC).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + contrato).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                            }
                            else
                            {
                                prfPrimero.Add("Referencia del crédito: ");
                                prfPrimero.Add(new Text(refMC).SetFont(boldFont));
                            }

                            prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 4 de enero de 2024, ");
                            prfPrimero.Add(new Text("ALERIN CONSULTING S.L.U.").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cedente").SetFont(boldFont));
                            prfPrimero.Add("“) cedió a ");
                            prfPrimero.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cesionario").SetFont(boldFont));
                            prfPrimero.Add("”) una cartera de créditos y, entre ellos, el crédito de referencia ");
                            prfPrimero.Add(new Text(contrato).SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Crédito").SetFont(boldFont));
                            prfPrimero.Add("“), que ostenta frente a usted en su calidad de Titular, con un saldo pendiente a fecha de ");
                            prfPrimero.Add(new Text("4 de enero de 2024").SetFont(boldFont));
                            prfPrimero.Add(" de ");
                            prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text("Banco Santander.").SetFont(boldFont));

                            prfPrimero.Add("\r\n\r\nEl Cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MCdos Legal S.L.").SetFont(boldFont));
                            prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                            doc.Add(prfPrimero);

                            //-------------------------------------------------------------

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, -40, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 194, 250)
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
                                .SetFixedPosition(330, 194, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }
                            else
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 218, 250)
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
                                    .SetFixedPosition(330, 218, 250)
                                    .SetFontColor(colorNegro)
                                    .SetFont(boldFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfSegundo = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -20, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfSegundo.Add("También le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a la Agencia de Cobro en el teléfono ");
                            prfSegundo.Add(new Text("91 108 89 04").SetFont(boldFont));
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
                                .SetPageNumber(1)
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

                            prfTercero.Add(new Text("\r\n\r\n\r\nALERIN CONSULTING S.L.U.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                            ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                            var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 1, 1)
                                .SetMaxWidth(150);
                            Paragraph firmaPrfIzq = new Paragraph("");
                            firmaPrfIzq.Add(firmaIzq);
                            doc.Add(firmaPrfIzq);

                            string rutaFirmaAKCP = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                            ImageData imagenDataFirmaAKCP = ImageDataFactory.Create(rutaFirmaAKCP);
                            var firmaAKCP = new iText.Layout.Element.Image(imagenDataFirmaAKCP)
                                .SetPageNumber(2)
                                .SetRelativePosition(270, -35, 1, 1)
                                .SetMaxWidth(180);
                            Paragraph firmaDch = new Paragraph("");
                            firmaDch.Add(firmaAKCP);
                            doc.Add(firmaDch);

                            //-------------------------------------------------------------

                            if (cp > 31000 && cp < 32000) //NAVARRA
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
                                prfNavarra.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de setenta y cinco mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
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
                                prfValencia.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfValencia.Add(" sociedad de nacionalidad española, con domicilio en Coslada (Madrid) 28021, Calle Olmo 7 Portal 10 7ª planta, inscrita en el Registro Mercantil de Madrid, al Tomo 42.096, Folio 110, Sección 8ª, Hoja M-745349, y con N.I.F.: B-06999957, ii) el Crédito se encuentra identificado en la página 54 del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de setenta y cinco mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -15, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(7)
                            .SetFixedLeading(8);

                            prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                            string link3 = "www.aepd.es";
                            PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation3.SetAction(PdfAction.CreateURI(link3));
                            linkAnnotation3.SetBorder(new PdfArray());
                            Link linkHtml2 = new Link(link3, linkAnnotation3);
                            prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                            prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                            prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                            prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                            doc.Add(prfCuarto);

                            //-------------------------------------------------------------

                            string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Proteccion_Datos.png"));
                            ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                            var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfTabla = new Paragraph();
                            prfTabla.Add(Tabla);
                            doc.Add(prfTabla);
                        }
                    }
                }
                else if (origen == "España")
                {
                    if (cp > 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraE, "Hello_" + refEnvio + ".pdf");
                    else if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaE, "Hello_" + refEnvio + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralE, "Hello_" + refEnvio + ".pdf");
                    string rutaArchivos = rutaArchivosE;
                    string rutaArchivosNavarra = rutaArchivosNavarraE;
                    string rutaArchivosValencia = rutaArchivosValenciaE;

                    using (var writter = new PdfWriter(exportarPDF))
                    {
                        using (var pdf = new PdfDocument(writter))
                        {
                            var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                            doc.SetMargins(65, 70, 80, 70); //Margenes PDF

                            //Definimos Tipografia
                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            //-------------------------------------------------------------

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));//Recogemos la ruta del archivo
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);//Creacion de la imagen
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)//Utilizamos la imagen
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(150)
                                .SetMarginBottom(48);
                            //Agregamos la imagen al PDF
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);
                            //
                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)//Creacion de la imagen
                                .SetFixedPosition(1, 455, 740)
                                .SetMaxWidth(70);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            //-------------------------------------------------------------                    

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MCdos Legal S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -75, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, 4 de enero de 2024");
                            doc.Add(prfDatosCliente);

                            //-------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                prfPrimero.Add(new Text("Referencia del crédito: " + refMC).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + contrato).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                            }
                            else
                            {
                                prfPrimero.Add("Referencia del crédito: ");
                                prfPrimero.Add(new Text(refMC).SetFont(boldFont));
                            }

                            prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 4 de enero de 2024, ");
                            prfPrimero.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cedente").SetFont(boldFont));
                            prfPrimero.Add("“) cedió a ");
                            prfPrimero.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cesionario").SetFont(boldFont));
                            prfPrimero.Add("”) una cartera de créditos y, entre ellos, el crédito de referencia ");
                            prfPrimero.Add(new Text(contrato).SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Crédito").SetFont(boldFont));
                            prfPrimero.Add("“), que ostenta frente a usted en su calidad de Titular, con un saldo pendiente a fecha de ");
                            prfPrimero.Add(new Text("4 de enero de 2024").SetFont(boldFont));
                            prfPrimero.Add(" de ");
                            prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text("Banco Santander.").SetFont(boldFont));

                            prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MCdos Legal S.L.").SetFont(boldFont));
                            prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                            doc.Add(prfPrimero);

                            //-------------------------------------------------------------

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 215, 250)
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
                                .SetFixedPosition(330, 215, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }
                            else
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 240, 250)
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
                                    .SetFixedPosition(330, 240, 250)
                                    .SetFontColor(colorNegro)
                                    .SetFont(boldFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfSegundo = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfSegundo.Add("También le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a la Agencia de Cobro en el teléfono ");
                            prfSegundo.Add(new Text("91 108 89 04").SetFont(boldFont));
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
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
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

                            prfTercero.Add(new Text("\r\n\r\n\r\nAXACTOR ESPAÑA , S.L.U.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                            ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                            var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 1, 1)
                                .SetMaxWidth(100);
                            Paragraph firmaPrfIzq = new Paragraph("");
                            firmaPrfIzq.Add(firmaIzq);
                            doc.Add(firmaPrfIzq);

                            string rutaFirmaDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                            ImageData imagenDataFirmaDch = ImageDataFactory.Create(rutaFirmaDch);
                            var firmaDch = new iText.Layout.Element.Image(imagenDataFirmaDch)
                                .SetPageNumber(2)
                                .SetRelativePosition(260, -65, 1, 1)
                                .SetMaxWidth(190);
                            Paragraph firmaPrfDch = new Paragraph("");
                            firmaPrfDch.Add(firmaDch);
                            doc.Add(firmaPrfDch);

                            //-------------------------------------------------------------

                            if (cp > 31000 && cp < 32000) //NAVARRA
                            {
                                Paragraph prfNavarra = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);

                                prfNavarra.Add("Asimismo, ");
                                prfNavarra.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de cuatrocientos dieciséis mil cuatrocientos cuarenta y dos con tres euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);
                            }
                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfValencia = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(10)
                                    .SetFixedLeading(11);

                                prfValencia.Add("Asimismo, ");
                                prfValencia.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfValencia.Add(" está domiciliada en Madrid (28007) Calle Doctor Esquerdo 136, 4ª planta, constituido el 09 de Julio de 2015 e inscrita en el Registro Mercantil de Madrid, al Tomo 33.781, Folio 113, Sección 8ª, Hoja M-607982, ii) el Crédito se encuentra identificado en la página 53 a 63 del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de cuatrocientos dieciséis mil cuatrocientos cuarenta y dos con tres euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -35, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(7)
                            .SetFixedLeading(8);

                            prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                            string link3 = "www.aepd.es";
                            PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation3.SetAction(PdfAction.CreateURI(link3));
                            linkAnnotation3.SetBorder(new PdfArray());
                            Link linkHtml2 = new Link(link3, linkAnnotation3);
                            prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                            prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                            prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                            prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                            doc.Add(prfCuarto);

                            //-------------------------------------------------------------

                            string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Proteccion_Datos.png"));
                            ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                            var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, -25, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfTabla = new Paragraph();
                            prfTabla.Add(Tabla);
                            doc.Add(prfTabla);
                        }
                    }
                }
                else if (origen == "Invest")
                {
                    if (cp > 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraI, "Hello_" + refEnvio + ".pdf");
                    else if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaI, "Hello_" + refEnvio + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralI, "Hello_" + refEnvio + ".pdf");
                    string rutaArchivos = rutaArchivosI;
                    string rutaArchivosNavarra = rutaArchivosNavarraI;
                    string rutaArchivosValencia = rutaArchivosValenciaI;

                    using (var writter = new PdfWriter(exportarPDF))
                    {
                        using (var pdf = new PdfDocument(writter))
                        {
                            var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                            doc.SetMargins(65, 70, 80, 70); //Margenes PDF

                            //Definimos Tipografia
                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            //-------------------------------------------------------------

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));//Recogemos la ruta del archivo
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);//Creacion de la imagen
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)//Utilizamos la imagen
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(150)
                                .SetMarginBottom(48);
                            //Agregamos la imagen al PDF
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);
                            //
                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)//Creacion de la imagen
                                .SetFixedPosition(1, 455, 740)
                                .SetMaxWidth(70);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            //-------------------------------------------------------------                    

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MCdos Legal S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -75, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, 4 de enero de 2024");
                            doc.Add(prfDatosCliente);

                            //-------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                prfPrimero.Add(new Text("Referencia del crédito: " + refMC).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + contrato).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                            }
                            else
                            {
                                prfPrimero.Add("Referencia del crédito: ");
                                prfPrimero.Add(new Text(refMC).SetFont(boldFont));
                            }

                            prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 4 de enero de 2024, ");
                            prfPrimero.Add(new Text(" AXACTOR INVEST 1, S.Á.R.L. ").SetFont(boldFont));
                            prfPrimero.Add(" (el \r\n“");
                            prfPrimero.Add(new Text("Cedente").SetFont(boldFont));
                            prfPrimero.Add("“) cedió a ");
                            prfPrimero.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Cesionario").SetFont(boldFont));
                            prfPrimero.Add("”) una cartera de créditos y, entre ellos, el crédito de referencia ");
                            prfPrimero.Add(new Text(contrato).SetFont(boldFont));
                            prfPrimero.Add(" (el “");
                            prfPrimero.Add(new Text("Crédito").SetFont(boldFont));
                            prfPrimero.Add("“), que ostenta frente a usted en su calidad de Titular, con un saldo pendiente a fecha de ");
                            prfPrimero.Add(new Text("4 de enero de 2024").SetFont(boldFont));
                            prfPrimero.Add(" de ");
                            prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text("Banco Santander.").SetFont(boldFont));

                            prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MCdos Legal S.L.").SetFont(boldFont));
                            prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                            doc.Add(prfPrimero);

                            //-------------------------------------------------------------

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 215, 250)
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
                                .SetFixedPosition(330, 215, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }
                            else
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 240, 250)
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
                                    .SetFixedPosition(330, 240, 250)
                                    .SetFontColor(colorNegro)
                                    .SetFont(boldFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfSegundo = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfSegundo.Add("También le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a la Agencia de Cobro en el teléfono ");
                            prfSegundo.Add(new Text("91 108 89 04").SetFont(boldFont));
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
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
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

                            prfTercero.Add(new Text("\r\n\r\n\r\nAXACTOR INVEST 1, S.Á.R.L.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                            ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                            var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 1, 1)
                                .SetMaxWidth(100);
                            Paragraph firmaPrfIzq = new Paragraph("");
                            firmaPrfIzq.Add(firmaIzq);
                            doc.Add(firmaPrfIzq);

                            string rutaFirmaDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                            ImageData imagenDataFirmaDch = ImageDataFactory.Create(rutaFirmaDch);
                            var firmaDch = new iText.Layout.Element.Image(imagenDataFirmaDch)
                                .SetPageNumber(2)
                                .SetRelativePosition(260, -65, 1, 1)
                                .SetMaxWidth(190);
                            Paragraph firmaPrfDch = new Paragraph("");
                            firmaPrfDch.Add(firmaDch);
                            doc.Add(firmaPrfDch);

                            //-------------------------------------------------------------

                            if (cp > 31000 && cp < 32000) //NAVARRA
                            {
                                Paragraph prfNavarra = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);

                                prfNavarra.Add("Asimismo, ");
                                prfNavarra.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de cincuenta y seis mil ochocientos ochenta y cinco con cero uno euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);
                            }
                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfValencia = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(10)
                                    .SetFixedLeading(11);

                                prfValencia.Add("Asimismo, ");
                                prfValencia.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfValencia.Add(" está domiciliada en 6, rue Eugène Ruppert, L-2453, Luxemburgo, ii) el Crédito se encuentra identificado en la página 53 a 54 del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de cincuenta y seis mil ochocientos ochenta y cinco con cero uno euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -35, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(7)
                            .SetFixedLeading(8);

                            prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                            string link3 = "www.aepd.es";
                            PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation3.SetAction(PdfAction.CreateURI(link3));
                            linkAnnotation3.SetBorder(new PdfArray());
                            Link linkHtml2 = new Link(link3, linkAnnotation3);
                            prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                            prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                            prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                            prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                            doc.Add(prfCuarto);

                            //-------------------------------------------------------------

                            string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Proteccion_Datos.png"));
                            ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                            var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, -25, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfTabla = new Paragraph();
                            prfTabla.Add(Tabla);
                            doc.Add(prfTabla);
                        }
                    }
                }             
                count++;
            }
            MessageBox.Show("Se han generado " + count + " pdfs.");
        }

        private void CrearPDF()
        {
            ConexionDB.AbrirConexion();

            string rutaArchivos = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Alerin");
            string rutaArchivosGeneral = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Alerin/General");
            string rutaArchivosNavarra = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Alerin/Navarra");
            string rutaArchivosValencia = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Alerin/Valencia");
            if (!Directory.Exists(rutaArchivos)) Directory.CreateDirectory(rutaArchivos);
            if (!Directory.Exists(rutaArchivosGeneral)) Directory.CreateDirectory(rutaArchivosGeneral);
            if (!Directory.Exists(rutaArchivosNavarra)) Directory.CreateDirectory(rutaArchivosNavarra);
            if (!Directory.Exists(rutaArchivosValencia)) Directory.CreateDirectory(rutaArchivosValencia);

            mcComm.command.Connection = conn.ObtenerConexion();
            //mcComm.CommandText = "SELECT OriginalAccountId,  REPLACE(FullName, '?', '') AS FullName_reemplazada, REPLACE(FullName2, '?', '') AS FullName2_reemplazada, REPLACE(Address1, '?', '') AS Address1_reemplazada, Zip, REPLACE(City, '?', '') AS City_reemplazada FROM TempPagantis";
            mcComm.CommandText = "SELECT Recibos.CodFactura AS CodFactura FROM Recibos INNER JOIN Expedientes ON Recibos.IdExpediente = Expedientes.IdExpediente WHERE Expedientes.IdCliente = 90 and IdSubcliente in (1052, 1053)";

            List<string> listDeudores = new List<string>();
            string importe = string.Empty;
            string importe2Decimales = string.Empty;

            using (IDataReader reader1 = mcComm.ExecuteReader())
            {
                while (reader1.Read())
                {
                    string referencia = reader1["CodFactura"].ToString();

                    using (SqlCommand innerCommand = new SqlCommand())
                    {
                        innerCommand.Connection = conn.ObtenerConexion();
                        innerCommand.CommandText = "SELECT Expedientes.DeudaTotal FROM Expedientes INNER JOIN Recibos ON Expedientes.IdExpediente = Recibos.IdExpediente WHERE IdCliente = 90 and IdSubcliente in (1052, 1053) and Recibos.CodFactura ='" + referencia + "'";

                        importe = innerCommand.ExecuteScalar()?.ToString();
                    }

                    if (!string.IsNullOrEmpty(importe))
                    {
                        importe = importe.Replace(".", ",");
                        importe2Decimales = importe.Substring(0, importe.IndexOf(',') + 3);
                    }
                    listDeudores.Add(referencia);
                }
                reader1.Close();
            }

            foreach (string referencia in listDeudores)
            {
                mcComm.CommandText = "SELECT Deudores.idDeudor AS Deudor, Deudores.NIF, REPLACE(Deudores.TITULAR, '?', '') AS Titular2, REPLACE(Deudores.Domicilio, '?', '') AS Domicilio2, REPLACE(Deudores.Localidad, '?', '') AS Localidad2, Deudores.CP FROM ExpedientesDeudores INNER JOIN Deudores ON ExpedientesDeudores.idDeudor = Deudores.idDeudor INNER JOIN Expedientes ON ExpedientesDeudores.IdExpediente = Expedientes.IdExpediente INNER JOIN Recibos ON Expedientes.IdExpediente = Recibos.IdExpediente WHERE Expedientes.IdCliente = 90 AND Expedientes.IdSubcliente IN (1052, 1053) AND Recibos.CodFactura ='" + referencia + "'";

                using (IDataReader reader2 = mcComm.ExecuteReader())
                {
                    while (reader2.Read())
                    {
                        string deudor = reader2["Deudor"].ToString();

                        mcComm.CommandText = "SELECT " +
                                                "NIF," +
                                                "REPLACE(TITULAR, '?', '') AS Titular2, " +
                                                "REPLACE(Domicilio, '?', '') AS Domicilio2, " +
                                                "REPLACE(Localidad, '?', '') AS Localidad2, " +
                                                "CP" +
                                                "FROM Deudores WHERE idDeudor ='" + deudor + "'";

                        string nombreMayusculas = reader2["Titular2"].ToString();
                        string direccion = reader2["Domicilio2"].ToString();
                        string localidad = reader2["Localidad2"].ToString();

                        int cp = reader2["CP"] != DBNull.Value ? Convert.ToInt32(reader2["CP"]) : 0;

                        string provincia = comp.CodigoPostal(cp);

                        string[] nombreMinusculas = nombreMayusculas.ToLower().Split(' ');
                        for (int i = 0; i < nombreMinusculas.Length; i++) if (nombreMinusculas[i].Length > 2) nombreMinusculas[i] = char.ToUpper(nombreMinusculas[i][0]) + nombreMinusculas[i].Substring(1);
                        string nombreFormateado = string.Join(" ", nombreMinusculas);

                        string[] palabrasDireccion = direccion.ToLower().Split(' ');
                        for (int i = 0; i < palabrasDireccion.Length; i++) if (palabrasDireccion[i].Length > 2) palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                        string direccionFormateada = string.Join(" ", palabrasDireccion);

                        string[] palabrasLocalidad = localidad.ToLower().Split(' ');
                        for (int i = 0; i < palabrasLocalidad.Length; i++) if (palabrasLocalidad[i].Length > 2) palabrasLocalidad[i] = char.ToUpper(palabrasLocalidad[i][0]) + palabrasLocalidad[i].Substring(1);
                        string localidadFormateada = string.Join(" ", palabrasLocalidad);

                        var exportarPDF = "";
                        if (cp > 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarra, "Hello_" + referencia + ".pdf");
                        else if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValencia, "Hello_" + referencia + ".pdf");
                        else exportarPDF = Path.Combine(rutaArchivosGeneral, "Hello_" + referencia + ".pdf");
                            
                        using (var writter = new PdfWriter(exportarPDF))
                        {
                            using (var pdf = new PdfDocument(writter))
                            {
                                var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                                doc.SetMargins(65, 70, 30, 70); //Margenes PDF

                                //Definimos Tipografia
                                string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                                string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                                PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                                PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                                iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                                iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                                //-------------------------------------------------------------

                                string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Alerin.jpg"));//Recogemos la ruta del archivo
                                ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);//Creacion de la imagen
                                var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)//Utilizamos la imagen
                                    .SetRelativePosition(0, -35, 0, 0)
                                    .SetMaxWidth(70)
                                    .SetMarginBottom(48);
                                //Agregamos la imagen al PDF
                                Paragraph encabezadoIzq = new Paragraph("");
                                encabezadoIzq.Add(logoIzq);
                                doc.Add(encabezadoIzq);
                                //
                                string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                                ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                                var logoDch = new iText.Layout.Element.Image(imagenDataDch)//Creacion de la imagen
                                    .SetFixedPosition(1, 455, 740)
                                    .SetMaxWidth(70);
                                Paragraph encabezadoDch = new Paragraph("");
                                encabezadoDch.Add(logoDch);
                                doc.Add(encabezadoDch);

                                //-------------------------------------------------------------                    

                                Paragraph prfMC2 = new Paragraph()
                                    .SetPageNumber(1)
                                    .SetRelativePosition(0, -30, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(boldFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                prfMC2.Add("MCdos Legal S.L.");
                                prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                                doc.Add(prfMC2);

                                Paragraph prfDatosCliente = new Paragraph()
                                    .SetPageNumber(1)
                                    .SetTextAlignment(TextAlignment.RIGHT)
                                    .SetRelativePosition(0, -110, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, 4 de enero de 2024");
                                doc.Add(prfDatosCliente);

                                //-------------------------------------------------------------

                                Paragraph prfPrimero = new Paragraph()
                                    .SetPageNumber(1)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -50, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);

                                if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                                {
                                    prfPrimero.Add(new Text("Referencia del crédito: " + referencia).SetFont(boldFont));
                                    prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + referencia).SetFont(boldFont));
                                    prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                                }
                                else
                                {
                                    prfPrimero.Add("Referencia del crédito: ");
                                    prfPrimero.Add(new Text(referencia).SetFont(boldFont));
                                }

                                prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 4 de enero de 2024, ");
                                prfPrimero.Add(new Text("ALERIN CONSULTING S.L.U.").SetFont(boldFont));
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
                                prfPrimero.Add(new Text("4 de enero de 2024").SetFont(boldFont));
                                prfPrimero.Add(" de ");
                                prfPrimero.Add(new Text(importe2Decimales + " €").SetFont(boldFont));
                                prfPrimero.Add(", cuyo origen es ");
                                prfPrimero.Add(new Text("Banco Santander.").SetFont(boldFont));

                                prfPrimero.Add("\r\n\r\nEl Cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                                prfPrimero.Add(new Text("MCdos Legal S.L.").SetFont(boldFont));
                                prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                                prfPrimero.Add("Por la presente, le requerimos para que - ");
                                prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                                prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                                prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                                doc.Add(prfPrimero);

                                //-------------------------------------------------------------

                                string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                                ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                                var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetMaxWidth(455);
                                Paragraph prfRecuadro = new Paragraph();
                                prfRecuadro.Add(Recuadro);
                                doc.Add(prfRecuadro);

                                if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                                {
                                    Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                    .SetPageNumber(1)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetFixedPosition(145, 194, 250)
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
                                    .SetFixedPosition(330, 194, 250)
                                    .SetFontColor(colorNegro)
                                    .SetFont(boldFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                    prfRecuadroInteriorDch.Add(new Text(referencia + "\r\n").SetFont(boldFont));
                                    prfRecuadroInteriorDch.Add(new Text(importe2Decimales + " €\r\n").SetFont(boldFont));
                                    prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                    doc.Add(prfRecuadroInteriorDch);
                                }
                                else
                                {
                                    Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                    .SetPageNumber(1)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetFixedPosition(145, 218, 250)
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
                                        .SetFixedPosition(330, 218, 250)
                                        .SetFontColor(colorNegro)
                                        .SetFont(boldFont)
                                        .SetFontSize(11)
                                        .SetFixedLeading(12);
                                    prfRecuadroInteriorDch.Add(new Text(referencia + "\r\n").SetFont(boldFont));
                                    prfRecuadroInteriorDch.Add(new Text(importe2Decimales + " €\r\n").SetFont(boldFont));
                                    prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                    doc.Add(prfRecuadroInteriorDch);
                                }

                                //-------------------------------------------------------------

                                Paragraph prfSegundo = new Paragraph()
                                    .SetPageNumber(1)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -20, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                prfSegundo.Add("También le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a la Agencia de Cobro en el teléfono ");
                                prfSegundo.Add(new Text("91 108 89 04").SetFont(boldFont));
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
                                    .SetPageNumber(1)
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

                                prfTercero.Add(new Text("\r\n\r\n\r\nALERIN CONSULTING S.L.U.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                                doc.Add(prfTercero);

                                //-------------------------------------------------------------

                                string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                                ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                                var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                    .SetPageNumber(2)
                                    .SetRelativePosition(0, 0, 1, 1)
                                    .SetMaxWidth(150);
                                Paragraph firmaPrfIzq = new Paragraph("");
                                firmaPrfIzq.Add(firmaIzq);
                                doc.Add(firmaPrfIzq);

                                string rutaFirmaAKCP = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                                ImageData imagenDataFirmaAKCP = ImageDataFactory.Create(rutaFirmaAKCP);
                                var firmaAKCP = new iText.Layout.Element.Image(imagenDataFirmaAKCP)
                                    .SetPageNumber(2)
                                    .SetRelativePosition(270, -35, 1, 1)
                                    .SetMaxWidth(180);
                                Paragraph firmaDch = new Paragraph("");
                                firmaDch.Add(firmaAKCP);
                                doc.Add(firmaDch);

                                //-------------------------------------------------------------

                                if (cp > 31000 && cp < 32000) //NAVARRA
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
                                    prfNavarra.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                    prfNavarra.Add(" y ");
                                    prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                    prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de setenta y cinco mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
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
                                    prfValencia.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                    prfValencia.Add(" cedio los creditos a ");
                                    prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                    prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                    prfValencia.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                    prfValencia.Add(" sociedad de nacionalidad española, con domicilio en Coslada (Madrid) 28021, Calle Olmo 7 Portal 10 7ª planta, inscrita en el Registro Mercantil de Madrid, al Tomo 42.096, Folio 110, Sección 8ª, Hoja M-745349, y con N.I.F.: B-06999957, ii) el Crédito se encuentra identificado en la página 54 del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de setenta y cinco mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                    doc.Add(prfValencia);
                                }

                                //-------------------------------------------------------------

                                Paragraph prfCuarto = new Paragraph()
                                .SetPageNumber(2)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -15, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(7)
                                .SetFixedLeading(8);

                                prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                                string link3 = "www.aepd.es";
                                PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                                linkAnnotation3.SetAction(PdfAction.CreateURI(link3));
                                linkAnnotation3.SetBorder(new PdfArray());
                                Link linkHtml2 = new Link(link3, linkAnnotation3);
                                prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                                prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                                prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                                prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                                doc.Add(prfCuarto);

                                //-------------------------------------------------------------

                                string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Proteccion_Datos.png"));
                                ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                                var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                                    .SetPageNumber(2)
                                    .SetRelativePosition(0, -10, 0, 0)
                                    .SetMaxWidth(455);
                                Paragraph prfTabla = new Paragraph();
                                prfTabla.Add(Tabla);
                                doc.Add(prfTabla);
                            }
                        }

                    }
                    reader2.Close();
                }                            
            }
            ConexionDB.CerrarConexion();
        }

        private string nHoja()
        {
            string hoja;
            var xlsApp = new Excel.Application();
            xlsApp.Workbooks.Open(txtFichero.Text.Trim());

            hoja = xlsApp.Sheets[1].Name;

            xlsApp.DisplayAlerts = false;
            xlsApp.Workbooks.Close();
            xlsApp.DisplayAlerts = true;

            xlsApp.Quit();
            xlsApp = null;

            return hoja;
        }
    }*/

    //CRISALIDA III - AXACTOR INVESTMENT - SANTANDER
    /*public partial class Form1 : Form
    {
        ConexionDB conn = new ConexionDB();
        MCCommand mcComm = new MCCommand();
        Comp comp = new Comp();

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

        private void CrearPDF()
        {
            ConexionDB.AbrirConexion();

            string rutaArchivos = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor_Invest");
            string rutaArchivosGeneral = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor_Invest/General");
            string rutaArchivosNavarra = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor_Invest/Navarra");
            string rutaArchivosValencia = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor_Invest/Valencia");
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
                    string provincia = comp.CodigoPostal(cp);

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
                            doc.SetMargins(65, 70, 80, 70); //Margenes PDF

                            //Definimos Tipografia
                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            //-------------------------------------------------------------

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));//Recogemos la ruta del archivo
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);//Creacion de la imagen
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)//Utilizamos la imagen
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(150)
                                .SetMarginBottom(48);
                            //Agregamos la imagen al PDF
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);
                            //
                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)//Creacion de la imagen
                                .SetFixedPosition(1, 455, 740)
                                .SetMaxWidth(70);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            //-------------------------------------------------------------                    

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MCdos Legal S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -75, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, 4 de enero de 2024");
                            doc.Add(prfDatosCliente);

                            //-------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                prfPrimero.Add(new Text("Referencia del crédito: " + referencia).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + referencia).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                            }
                            else
                            {
                                prfPrimero.Add("Referencia del crédito: ");
                                prfPrimero.Add(new Text(referencia).SetFont(boldFont));
                            }

                            prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 4 de enero de 2024, ");
                            prfPrimero.Add(new Text(" AXACTOR INVEST 1, S.Á.R.L. ").SetFont(boldFont));
                            prfPrimero.Add(" (el \r\n“");
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
                            prfPrimero.Add(new Text("4 de enero de 2024").SetFont(boldFont));
                            prfPrimero.Add(" de ");
                            prfPrimero.Add(new Text(importe2Decimales + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text("Banco Santander.").SetFont(boldFont));

                            prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MCdos Legal S.L.").SetFont(boldFont));
                            prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                            doc.Add(prfPrimero);

                            //-------------------------------------------------------------

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 227, 250)
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
                                .SetFixedPosition(330, 227, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(referencia + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe2Decimales + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }
                            else
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 252, 250)
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
                                    .SetFixedPosition(330, 252, 250)
                                    .SetFontColor(colorNegro)
                                    .SetFont(boldFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(referencia + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe2Decimales + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfSegundo = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfSegundo.Add("También le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a la Agencia de Cobro en el teléfono ");
                            prfSegundo.Add(new Text("91 108 89 04").SetFont(boldFont));
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
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
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

                            prfTercero.Add(new Text("\r\n\r\n\r\nAXACTOR INVEST 1, S.Á.R.L.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                            ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                            var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 1, 1)
                                .SetMaxWidth(100);
                            Paragraph firmaPrfIzq = new Paragraph("");
                            firmaPrfIzq.Add(firmaIzq);
                            doc.Add(firmaPrfIzq);

                            string rutaFirmaDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                            ImageData imagenDataFirmaDch = ImageDataFactory.Create(rutaFirmaDch);
                            var firmaDch = new iText.Layout.Element.Image(imagenDataFirmaDch)
                                .SetPageNumber(2)
                                .SetRelativePosition(260, -65, 1, 1)
                                .SetMaxWidth(190);
                            Paragraph firmaPrfDch = new Paragraph("");
                            firmaPrfDch.Add(firmaDch);
                            doc.Add(firmaPrfDch);

                            //-------------------------------------------------------------

                            if (cp > 31000 && cp < 32000) //NAVARRA
                            {
                                Paragraph prfNavarra = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);

                                prfNavarra.Add("Asimismo, ");
                                prfNavarra.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de cincuenta y seis mil ochocientos ochenta y cinco con cero uno euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);
                            }
                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfValencia = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(10)
                                    .SetFixedLeading(11);

                                prfValencia.Add("Asimismo, ");
                                prfValencia.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("AXACTOR INVEST 1, S.Á.R.L.").SetFont(boldFont));
                                prfValencia.Add(" está domiciliada en 6, rue Eugène Ruppert, L-2453, Luxemburgo, ii) el Crédito se encuentra identificado en la página 53 a 54 del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de cincuenta y seis mil ochocientos ochenta y cinco con cero uno euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -35, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(7)
                            .SetFixedLeading(8);

                            prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                            string link3 = "www.aepd.es";
                            PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation3.SetAction(PdfAction.CreateURI(link3));
                            linkAnnotation3.SetBorder(new PdfArray());
                            Link linkHtml2 = new Link(link3, linkAnnotation3);
                            prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                            prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                            prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                            prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                            doc.Add(prfCuarto);

                            //-------------------------------------------------------------

                            string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Proteccion_Datos.png"));
                            ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                            var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, -25, 0, 0)
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

    //CRISALIDA III - AXACTOR ESPAÑA - SANTANDER
    /*public partial class Form1 : Form
    {
        ConexionDB conn = new ConexionDB();
        MCCommand mcComm = new MCCommand();
        Comp comp = new Comp();

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

        private void CrearPDF()
        {
            ConexionDB.AbrirConexion();

            string rutaArchivos = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor_España");
            string rutaArchivosGeneral = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor_España/General");
            string rutaArchivosNavarra = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor_España/Navarra");
            string rutaArchivosValencia = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Axactor_España/Valencia");
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
                    string provincia = comp.CodigoPostal(cp);

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
                            doc.SetMargins(65, 70, 80, 70); //Margenes PDF

                            //Definimos Tipografia
                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            //-------------------------------------------------------------

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));//Recogemos la ruta del archivo
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);//Creacion de la imagen
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)//Utilizamos la imagen
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(150)
                                .SetMarginBottom(48);
                            //Agregamos la imagen al PDF
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);
                            //
                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)//Creacion de la imagen
                                .SetFixedPosition(1, 455, 740)
                                .SetMaxWidth(70);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            //-------------------------------------------------------------                    

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MCdos Legal S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911175438\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -75, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + cp + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, 4 de enero de 2024");
                            doc.Add(prfDatosCliente);

                            //-------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);                            

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                prfPrimero.Add(new Text("Referencia del crédito: " + referencia).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + referencia).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                            }
                            else
                            {
                                prfPrimero.Add("Referencia del crédito: ");
                                prfPrimero.Add(new Text(referencia).SetFont(boldFont));
                            }

                            prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 4 de enero de 2024, ");
                            prfPrimero.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
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
                            prfPrimero.Add(new Text("4 de enero de 2024").SetFont(boldFont));
                            prfPrimero.Add(" de ");
                            prfPrimero.Add(new Text(importe2Decimales + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text("Banco Santander.").SetFont(boldFont));

                            prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MCdos Legal S.L.").SetFont(boldFont));
                            prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                            doc.Add(prfPrimero);

                            //-------------------------------------------------------------

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, -10, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 227, 250)
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
                                .SetFixedPosition(330, 227, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(referencia + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe2Decimales + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }
                            else
                            {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 252, 250)
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
                                    .SetFixedPosition(330, 252, 250)
                                    .SetFontColor(colorNegro)
                                    .SetFont(boldFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(referencia + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe2Decimales + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfSegundo = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfSegundo.Add("También le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a la Agencia de Cobro en el teléfono ");
                            prfSegundo.Add(new Text("91 108 89 04").SetFont(boldFont));
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
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
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

                            prfTercero.Add(new Text("\r\n\r\n\r\nAXACTOR ESPAÑA , S.L.U.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                            ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                            var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 1, 1)
                                .SetMaxWidth(100);
                            Paragraph firmaPrfIzq = new Paragraph("");
                            firmaPrfIzq.Add(firmaIzq);
                            doc.Add(firmaPrfIzq);

                            string rutaFirmaDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                            ImageData imagenDataFirmaDch = ImageDataFactory.Create(rutaFirmaDch);
                            var firmaDch = new iText.Layout.Element.Image(imagenDataFirmaDch)
                                .SetPageNumber(2)
                                .SetRelativePosition(260, -65, 1, 1)
                                .SetMaxWidth(190);
                            Paragraph firmaPrfDch = new Paragraph("");
                            firmaPrfDch.Add(firmaDch);
                            doc.Add(firmaPrfDch);

                            //-------------------------------------------------------------

                            if (cp > 31000 && cp < 32000) //NAVARRA
                            {
                                Paragraph prfNavarra = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(11)
                                    .SetFixedLeading(12);

                                prfNavarra.Add("Asimismo, ");
                                prfNavarra.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de cuatrocientos dieciséis mil cuatrocientos cuarenta y dos con tres euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);
                            }
                            if (cp > 3000 && cp < 4000 || cp > 12000 && cp < 13000 || cp > 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                            {
                                Paragraph prfValencia = new Paragraph()
                                    .SetPageNumber(2)
                                    .SetVerticalAlignment(VerticalAlignment.TOP)
                                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                                    .SetRelativePosition(0, -40, 0, 0)
                                    .SetFontColor(colorNegro)
                                    .SetFont(regularFont)
                                    .SetFontSize(10)
                                    .SetFixedLeading(11);

                                prfValencia.Add("Asimismo, ");
                                prfValencia.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfValencia.Add(" está domiciliada en Madrid (28007) Calle Doctor Esquerdo 136, 4ª planta, constituido el 09 de Julio de 2015 e inscrita en el Registro Mercantil de Madrid, al Tomo 33.781, Folio 113, Sección 8ª, Hoja M-607982, ii) el Crédito se encuentra identificado en la página 53 a 63 del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de cuatrocientos dieciséis mil cuatrocientos cuarenta y dos con tres euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");                                
                                doc.Add(prfValencia);                               
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -35, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(7)
                            .SetFixedLeading(8);

                            prfCuarto.Add("De conformidad con lo previsto en la Ley Orgánica 3/2018 de 5 de diciembre, de Protección de Datos Personales y garantía de los derechos digitales, el Reglamento (UE) 2016/679 del Parlamento Europeo y del Consejo de 27 de abril de 2016 y demás normativa aplicable en materia de protección de datos, mediante la presente comunicación se le informa de que sus datos personales han sido cedidos por .a AKCP Europe SCSp (“Arbor Knot”) con domicilio, a estos efectos, en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgo que, en su condición de responsable, tratará los datos para la finalidad de poder ejercer y gestionar el derecho del crédito que ostenta frente a usted, así como, en su caso, la elaboración de perfiles que podrán dar lugar a la toma de decisiones automatizadas para facilitar el pago de la deuda pendiente. Usted podrá ejercer los derechos de acceso, rectificación, oposición, supresión, limitación del tratamiento, portabilidad de datos y a no ser objeto de decisiones individualizadas automatizadas y cualesquiera otros que resulten de aplicación, mediante el envío de una carta dirigida a Arbor Knot en la dirección del encabezado de esta carta acompañando copia de su D.N.I. o de otro documento que lo identifique o por correo electrónico a la dirección: privacy@arborknot.io. Las causas legitimadoras de los tratamientos descritos son: (i) la ejecución y control de la relación contractual con usted (ii) el cumplimiento de obligaciones legales a las que está sujeta el responsable del tratamiento y el (iii) interés legítimo de Arbor Knot. Los datos personales serán tratados después de la cancelación del derecho de crédito que el responsable del tratamiento ostenta frente a usted en tanto pudieran derivarse responsabilidades de su relación con aquél y con la sola finalidad de dar cumplimiento a cualquier ley aplicable o para ofrecerle productos o servicios que mejoren su capacidad financiera, siempre y cuando haya consentido dicho tratamiento. En cualquier caso, le informamos que puede presentar una reclamación ante la Agencia Española de Protección de Datos (");

                            string link3 = "www.aepd.es";
                            PdfLinkAnnotation linkAnnotation3 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                            linkAnnotation3.SetAction(PdfAction.CreateURI(link3));
                            linkAnnotation3.SetBorder(new PdfArray());
                            Link linkHtml2 = new Link(link3, linkAnnotation3);
                            prfCuarto.Add(linkHtml2.SetFontColor(colorAzul).SetUnderline());

                            prfCuarto.Add(").\r\n\r\nPor último, Arbor Knot le remite a la siguiente dirección web ");
                            prfCuarto.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());
                            prfCuarto.Add(" para cualquier consulta sobre nuestra política de privacidad y nuestros colaboradores que podrán acceder a sus datos personales cuando nos presten sus servicios.");
                            doc.Add(prfCuarto);

                            //-------------------------------------------------------------

                            string rutaTabla = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Proteccion_Datos.png"));
                            ImageData imagenDataTabla = ImageDataFactory.Create(rutaTabla);
                            var Tabla = new iText.Layout.Element.Image(imagenDataTabla)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, -25, 0, 0)
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

    //PAGANTIS
    /*public partial class Form1 : Form
    {
        ConexionDB conn = new ConexionDB();
        MCCommand mcComm = new MCCommand();
        Comp comp = new Comp();

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
                    string provincia = comp.CodigoPostal(cp);

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
                            doc.SetMargins(65, 70, 90, 70); //Margenes PDF

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

