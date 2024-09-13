using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Action;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography;
using iText.Svg.Renderers.Path.Impl;
using System.Drawing;
using System.Diagnostics.Contracts;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/*
 * Aqui se recogen todas las HelloLetters
 * Con cada nuevo envio, duplicamos la ultiuma clase y modificamos los nuevos campos
 * Analizar los encabezados del Excel con la informacion
 * La ruta principal en el ordenador es 'Mis documentos'
 * Para el codigo postal nos fijaremos mejor en la funcion creada que en el valor proporcionado en el Excel.
 * El recuadro ira apoyado en el primer parrafo y 'prfRecuadroInteriorIzq' y 'prfRecuadroInteriorDch' se tendra que ajustar a mano
 */
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

namespace Creacion_PDF_HelloLetter
{
    //CRISALIDA 4/CRISALIDA 5/CRISALIDA 6 - AXACTOR LUXEMBURGO/AXACTOR ESPAÑA/AXACTOR INVEST - SANTANDER
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
            string idCliente = "95"; //"93"  "95"
            string carpeta = "Crisalida 6";

            OleDbConnection oleConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim() + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter("SELECT * FROM [" + hoja + "$]", oleConnection);

            DataSet ds = new DataSet();
            oleAdapter.Fill(ds);
            oleConnection.Close();
            DataTable dt = ds.Tables[0];

            foreach (DataRow fila in dt.Rows)
            {
                string origen = fila["ORIGEN"].ToString();
                string refMC = fila["REF_MC"].ToString();
                string origenBanco = fila["Origen Bancario"].ToString();
                string contrato = fila["CONTRATO"].ToString();
                string refEnvio = fila["REF_ENVIO"].ToString();
                string importe = fila["IMPORTE"].ToString();
                string nombre = fila["NOMBRE"].ToString();
                object municipio = fila["MUNICIPIO"] != DBNull.Value ? fila["MUNICIPIO"].ToString() : "";
                object direccion = fila["DIRECCION_1"] != DBNull.Value ? fila["DIRECCION_1"].ToString() : "";
                object direccion2 = fila["DIRECCION_2"] != DBNull.Value ? fila["DIRECCION_2"].ToString() : "";
                direccion = direccion + " " + direccion2;
                object cp = fila["CP"] != DBNull.Value ? (object)fila["CP"] : 0;
                string tipo = fila["TIPO"].ToString();
                string provincia = fila["PROVINCIA"].ToString();
                string fechaEnvio = "27 de mayo de 2024";
                string fechaComunicacion = "22 de mayo de 2024";
                string fechaCompra = "22 de mayo de 2024";
                //string paginasLux = "52 a 54"; //Crisalida 4
                //string paginasEsp = "53 a 54"; //Crisalida 4
                //string paginasInv = "53 a 56"; //Crisalida 4
                //string paginasLux = "52 a 54"; //Crisalida 5
                //string paginasEsp = "53 a 54"; //Crisalida 5
                //string paginasInv = "53 a 56"; //Crisalida 5
                string paginasLux = "53 a 55"; //Crisalida 6
                string paginasEsp = "53"; //Crisalida 5
                string paginasInv = "53 a 55"; //Crisalida 5
                //string costeCartera = "sesenta y cinco mil novecientos setenta y dos euros con cuarenta y seis céntimos"; //Crisalida 4
                //string costeCartera = "cincuenta mil ochocientos veinticuatro euros con quince céntimos"; //Crisalida 5
                string costeCartera = "ciento setenta mil setenta y dos euros con ochenta y dos céntimos"; //Crisalida 6

                string[] nombreMinusculas = nombre.ToLower().Split(' ');
                for (int i = 0; i < nombreMinusculas.Length; i++) if (nombreMinusculas[i].Length > 2) nombreMinusculas[i] = char.ToUpper(nombreMinusculas[i][0]) + nombreMinusculas[i].Substring(1);
                string nombreFormateado = string.Join(" ", nombreMinusculas);
                if (nombreFormateado.Length > 59) nombreFormateado = nombreFormateado.Substring(0, 59);

                string[] palabrasLocalidad = municipio.ToString().ToLower().Split(' ');
                for (int i = 0; i < palabrasLocalidad.Length; i++) if (palabrasLocalidad[i].Length > 2) palabrasLocalidad[i] = char.ToUpper(palabrasLocalidad[i][0]) + palabrasLocalidad[i].Substring(1);
                string localidadFormateada = string.Join(" ", palabrasLocalidad);
                if (localidadFormateada.Length > 59) localidadFormateada = localidadFormateada.Substring(0, 59);

                //string[] palabrasDireccion = direccion.ToLower().Split(' ');
                //for (int i = 0; i < palabrasDireccion.Length; i++) if (palabrasDireccion[i].Length > 2) palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);

                string[] palabrasDireccion = direccion.ToString().ToLower().Split(' ');
                for (int i = 0; i < palabrasDireccion.Length; i++){
                    if (palabrasDireccion[i].Length > 2) {
                        palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                        if (palabrasDireccion[i] == "null" || palabrasDireccion[i] == "(null)") palabrasDireccion[i] = string.Empty;
                    }
                }
                string direccionFormateada = string.Join(" ", palabrasDireccion);
                if (direccionFormateada.Length > 59) direccionFormateada = direccionFormateada.Substring(0, 59);

                int iCP;
                if (int.TryParse(cp.ToString(), out iCP)) // Converimos cp a entero iCP

                if (provincia == string.Empty) if (int.TryParse(cp.ToString(), out iCP)) provincia = comp.CodigoPostal(iCP); // Si provincia es vacio y cp es entero, convertimos cp a provincia

                string sCP = cp.ToString() == string.Empty || cp.ToString() == "null" || cp.ToString() == "0" ? "" : cp.ToString(); // Si cp es vacio o 'null', ponemos vacio 

                var exportarPDF = "";
                if (origen == "Axactor Capital Luxemburgo S.A.R.L.") //"Axactor Capital Luxemburgo S.A.R.L."
                {
                    string rutaArchivosL = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor Luxemburgo");
                    string rutaArchivosGeneralL = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor Luxemburgo/General");
                    string rutaArchivosNavarraL = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor Luxemburgo/Navarra");
                    string rutaArchivosValenciaL = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor Luxemburgo/Valencia");
                    if (!Directory.Exists(rutaArchivosGeneralL)) Directory.CreateDirectory(rutaArchivosGeneralL);
                    if (!Directory.Exists(rutaArchivosNavarraL)) Directory.CreateDirectory(rutaArchivosNavarraL);
                    if (!Directory.Exists(rutaArchivosValenciaL)) Directory.CreateDirectory(rutaArchivosValenciaL);

                    if (iCP >= 31000 && iCP < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraL, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    else if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaL, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralL, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    string rutaArchivos = rutaArchivosL;
                    string rutaArchivosNavarra = rutaArchivosNavarraL;
                    string rutaArchivosValencia = rutaArchivosValenciaL;

                    using (var writter = new PdfWriter(exportarPDF)) {
                        using (var pdf = new PdfDocument(writter)) {
                            var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                            doc.SetMargins(50, 70, 50, 70);

                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            // LOGOS - En la izquierda se apoya sobre los margenes del formato A4 y en la derecha ajustamos su posicion de manera libre

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)
                                .SetRelativePosition(-2, 0, 0, 0)
                                .SetMaxWidth(70)
                                .SetMarginBottom(48);
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);

                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)
                                .SetFixedPosition(1, 385, 750)
                                .SetMaxWidth(140);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            // ENCABEZADO - En la izq. los datos se mueven por libre

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                //.SetRelativePosition(0, 0, 0, 0)
                                .SetFixedPosition(70, 620, 750)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MCdos LEGAL S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 910883105\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -24, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + sCP + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, " + fechaEnvio);
                            doc.Add(prfDatosCliente);

                            // PRIMER PARRAFO -------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000) {// ALICANTE // CASTELLON // VALENCIA                                 prfPrimero.Add(new Text("Referencia del crédito: " + refMC).SetFont(boldFont));
                                prfPrimero.Add(new Text("Referencia del crédito: " + refMC).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + contrato).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                            } else {
                                prfPrimero.Add("Referencia del crédito: ");
                                prfPrimero.Add(new Text(refMC).SetFont(boldFont));
                            }
                            prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha " + fechaComunicacion + ", ");
                            prfPrimero.Add(new Text("Axactor Capital Luxembourg S.á.r.l").SetFont(boldFont));
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
                            prfPrimero.Add("“), que ostenta frente a usted en su calidad de ");
                            prfPrimero.Add(new Text(tipo).SetFont(boldFont));
                            prfPrimero.Add(", con un saldo pendiente a fecha de " + fechaCompra + " de ");
                            prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text(origenBanco).SetFont(boldFont));

                            prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MCdos LEGAL S.L.").SetFont(boldFont));
                            prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                            doc.Add(prfPrimero);

                            // RECUADRO - En el caso de C.Valenciana hay que bajarlo puntos mas

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, 5, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000) {// ALICANTE // CASTELLON // VALENCIA
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 249, 250)
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
                                .SetFixedPosition(330, 249, 250)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                                prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                                prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                                doc.Add(prfRecuadroInteriorDch);
                            } 
                            else {
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 273, 250)
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
                                    .SetFixedPosition(330, 273, 250)
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
                            prfSegundo.Add("\r\n\r\nTambién le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a ");
                            prfSegundo.Add(new Text("MCDOS LEGAL S.L.").SetFont(boldFont));
                            prfSegundo.Add(" en el teléfono ");
                            prfSegundo.Add(new Text("91 088 31 05").SetFont(boldFont));
                            prfSegundo.Add(" o en la dirección de correo electrónico ");

                            string link1 = "contencioso@fondos.mc2legal.es ";
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

                            prfTercero.Add(new Text("\r\n\r\n\r\nAxactor Capital Luxembourg S.á.r.l\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t  AKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                            ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                            var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 1, 1)
                                .SetMaxWidth(110);
                            Paragraph firmaPrfIzq = new Paragraph("");
                            firmaPrfIzq.Add(firmaIzq);
                            doc.Add(firmaPrfIzq);

                            string rutaFirmaAKCP = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                            ImageData imagenDataFirmaAKCP = ImageDataFactory.Create(rutaFirmaAKCP);
                            var firmaAKCP = new iText.Layout.Element.Image(imagenDataFirmaAKCP)
                                .SetPageNumber(2)
                                .SetRelativePosition(270, -60, 1, 1)
                                .SetMaxWidth(180);
                            Paragraph firmaDch = new Paragraph("");
                            firmaDch.Add(firmaAKCP);
                            doc.Add(firmaDch);

                            //-------------------------------------------------------------

                            if (iCP >= 31000 && iCP < 32000) {//NAVARRA
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
                                prfNavarra.Add(new Text("Axactor Capital Luxembourg S.á.r.l").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de ");
                                prfNavarra.Add(costeCartera + "."); 
                                prfNavarra.Add(" Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);
                            }
                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000) {// ALICANTE // CASTELLON // VALENCIA
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
                                prfValencia.Add(new Text("Axactor Capital Luxembourg S.á.r.l").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("Axactor Capital Luxembourg S.á.r.l").SetFont(boldFont));
                                prfValencia.Add(" domiciliada en 6, rue Eugène Ruppert, L-2453, Luxemburgo, inscrita en el Registro Mercantil de Luxemburgo, sección B con número 217.699. ii) el Crédito se encuentra identificado en la página " + paginasLux + " del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de ");
                                prfValencia.Add(costeCartera + ".");
                                prfValencia.Add(" Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -20, 0, 0)
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
                                .SetRelativePosition(0, -20, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfTabla = new Paragraph();
                            prfTabla.Add(Tabla);
                            doc.Add(prfTabla);
                            count++;
                        }
                    }
                }
                else if (origen == "Axactor España, S.L.U.") //"Axactor España, S.L.U."
                {
                    string rutaArchivosE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor España");
                    string rutaArchivosGeneralE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor España/General");
                    string rutaArchivosNavarraE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor España/Navarra");
                    string rutaArchivosValenciaE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor España/Valencia");
                    if (!Directory.Exists(rutaArchivosGeneralE)) Directory.CreateDirectory(rutaArchivosGeneralE);
                    if (!Directory.Exists(rutaArchivosNavarraE)) Directory.CreateDirectory(rutaArchivosNavarraE);
                    if (!Directory.Exists(rutaArchivosValenciaE)) Directory.CreateDirectory(rutaArchivosValenciaE);

                    if (iCP >= 31000 && iCP < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraE, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    else if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaE, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralE, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    string rutaArchivos = rutaArchivosE;
                    string rutaArchivosNavarra = rutaArchivosNavarraE;
                    string rutaArchivosValencia = rutaArchivosValenciaE;

                    using (var writter = new PdfWriter(exportarPDF))
                    {
                        using (var pdf = new PdfDocument(writter))
                        {
                            var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                            doc.SetMargins(50, 70, 50, 70);

                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            // LOGOS - En la izquierda se apoya sobre los margenes del formato A4 y en la derecha ajustamos su posicion de manera libre

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)
                                .SetRelativePosition(-2, 0, 0, 0)
                                .SetMaxWidth(70)
                                .SetMarginBottom(48);
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);

                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)
                                .SetFixedPosition(1, 385, 750)
                                .SetMaxWidth(140);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            // ENCABEZADO - En la izq. los datos se mueven por libre

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                //.SetRelativePosition(0, 0, 0, 0)
                                .SetFixedPosition(70, 620, 750)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MCdos LEGAL S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 910883105\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -24, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + sCP + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, " + fechaEnvio);
                            doc.Add(prfDatosCliente);

                            // PRIMER PARRAFO -------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000)
                            {// ALICANTE // CASTELLON // VALENCIA                                
                                prfPrimero.Add(new Text("Referencia del crédito: " + refMC).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + contrato).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                            }
                            else
                            {
                                prfPrimero.Add("Referencia del crédito: ");
                                prfPrimero.Add(new Text(refMC).SetFont(boldFont));
                            }
                            prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha " + fechaCompra + ", ");
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
                            prfPrimero.Add("“), que ostenta frente a usted en su calidad de ");
                            prfPrimero.Add(new Text(tipo).SetFont(boldFont));
                            prfPrimero.Add(", con un saldo pendiente a fecha de " + fechaCompra + " de ");
                            prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text(origenBanco).SetFont(boldFont));

                            prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MCdos LEGAL S.L.").SetFont(boldFont));
                            prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                            doc.Add(prfPrimero);

                            // RECUADRO - En el caso de C.Valenciana hay que bajarlo puntos mas

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, 5, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000)
                            {// ALICANTE // CASTELLON // VALENCIA
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 249, 250)
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
                                .SetFixedPosition(330, 249, 250)
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
                                .SetFixedPosition(145, 273, 250)
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
                                    .SetFixedPosition(330, 273, 250)
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
                            prfSegundo.Add("\r\n\r\nTambién le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a ");
                            prfSegundo.Add(new Text("MCDOS LEGAL S.L.").SetFont(boldFont));
                            prfSegundo.Add(" en el teléfono ");
                            prfSegundo.Add(new Text("91 088 31 05").SetFont(boldFont));
                            prfSegundo.Add(" o en la dirección de correo electrónico ");

                            string link1 = "contencioso@fondos.mc2legal.es ";
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

                            prfTercero.Add(new Text("\r\n\r\n\r\nAXACTOR ESPAÑA, S.L.U.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                            ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                            var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 1, 1)
                                .SetMaxWidth(110);
                            Paragraph firmaPrfIzq = new Paragraph("");
                            firmaPrfIzq.Add(firmaIzq);
                            doc.Add(firmaPrfIzq);

                            string rutaFirmaAKCP = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                            ImageData imagenDataFirmaAKCP = ImageDataFactory.Create(rutaFirmaAKCP);
                            var firmaAKCP = new iText.Layout.Element.Image(imagenDataFirmaAKCP)
                                .SetPageNumber(2)
                                .SetRelativePosition(270, -60, 1, 1)
                                .SetMaxWidth(180);
                            Paragraph firmaDch = new Paragraph("");
                            firmaDch.Add(firmaAKCP);
                            doc.Add(firmaDch);

                            //-------------------------------------------------------------

                            if (iCP >= 31000 && iCP < 32000)
                            {//NAVARRA
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
                                prfNavarra.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de ");
                                prfNavarra.Add(costeCartera + ".");
                                prfNavarra.Add(" Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);
                            }
                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000)
                            {// ALICANTE // CASTELLON // VALENCIA
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
                                prfValencia.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                                prfValencia.Add(" está domiciliada en Madrid (28007) Calle Doctor Esquerdo 136, 4ª planta, constituido el 09 de Julio de 2015 e inscrita en el Registro Mercantil de Madrid, al Tomo 33.781, Folio 113, Sección 8ª, Hoja M-607982, ii) el Crédito se encuentra identificado en la página " + paginasEsp + " del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de ");
                                prfValencia.Add(costeCartera + ".");
                                prfValencia.Add(" Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -20, 0, 0)
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
                                .SetRelativePosition(0, -20, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfTabla = new Paragraph();
                            prfTabla.Add(Tabla);
                            doc.Add(prfTabla);
                            count++;
                        }
                    }
                }
                else if (origen == "Axactor Invest 1 SARL") //"Axactor Invest 1 SARL"
                {
                    string rutaArchivosI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor Invest");
                    string rutaArchivosGeneralI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor Invest/General");
                    string rutaArchivosNavarraI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor Invest/Navarra");
                    string rutaArchivosValenciaI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/" + carpeta + "/Axactor Invest/Valencia");
                    if (!Directory.Exists(rutaArchivosGeneralI)) Directory.CreateDirectory(rutaArchivosGeneralI);
                    if (!Directory.Exists(rutaArchivosNavarraI)) Directory.CreateDirectory(rutaArchivosNavarraI);
                    if (!Directory.Exists(rutaArchivosValenciaI)) Directory.CreateDirectory(rutaArchivosValenciaI);

                    if (iCP >= 31000 && iCP < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraI, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    else if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaI, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralI, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                    string rutaArchivos = rutaArchivosI;
                    string rutaArchivosNavarra = rutaArchivosNavarraI;
                    string rutaArchivosValencia = rutaArchivosValenciaI;

                    using (var writter = new PdfWriter(exportarPDF))
                    {
                        using (var pdf = new PdfDocument(writter))
                        {
                            var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                            doc.SetMargins(50, 70, 50, 70);

                            string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                            string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                            PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                            PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                            iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                            iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                            // LOGOS - En la izquierda se apoya sobre los margenes del formato A4 y en la derecha ajustamos su posicion de manera libre

                            string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                            ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);
                            var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)
                                .SetRelativePosition(-2, 0, 0, 0)
                                .SetMaxWidth(70)
                                .SetMarginBottom(48);
                            Paragraph encabezadoIzq = new Paragraph("");
                            encabezadoIzq.Add(logoIzq);
                            doc.Add(encabezadoIzq);

                            string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));
                            ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                            var logoDch = new iText.Layout.Element.Image(imagenDataDch)
                                .SetFixedPosition(1, 385, 750)
                                .SetMaxWidth(140);
                            Paragraph encabezadoDch = new Paragraph("");
                            encabezadoDch.Add(logoDch);
                            doc.Add(encabezadoDch);

                            // ENCABEZADO - En la izq. los datos se mueven por libre

                            Paragraph prfMC2 = new Paragraph()
                                .SetPageNumber(1)
                                //.SetRelativePosition(0, 0, 0, 0)
                                .SetFixedPosition(70, 620, 750)
                                .SetFontColor(colorNegro)
                                .SetFont(boldFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfMC2.Add("MCdos LEGAL S.L.");
                            prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 910883105\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                            doc.Add(prfMC2);

                            Paragraph prfDatosCliente = new Paragraph()
                                .SetPageNumber(1)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetRelativePosition(0, -24, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);
                            prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + sCP + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, " + fechaEnvio);
                            doc.Add(prfDatosCliente);

                            // PRIMER PARRAFO -------------------------------------------------------------

                            Paragraph prfPrimero = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, 0, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(11)
                                .SetFixedLeading(12);

                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000)
                            {// ALICANTE // CASTELLON // VALENCIA                             
                                prfPrimero.Add(new Text("Referencia del crédito: " + refMC).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nCódigo de identificación del Contrato nº " + contrato).SetFont(boldFont));
                                prfPrimero.Add(new Text("\r\nContrato incluido en el Anexo V del Contrato de Cesión de la Cartera de Créditos").SetFont(boldFont));
                            }
                            else
                            {
                                prfPrimero.Add("Referencia del crédito: ");
                                prfPrimero.Add(new Text(refMC).SetFont(boldFont));
                            }
                            prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha " + fechaCompra + ", ");
                            prfPrimero.Add(new Text("Axactor Invest 1, S.á.r.l.").SetFont(boldFont));
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
                            prfPrimero.Add("“), que ostenta frente a usted en su calidad de ");
                            prfPrimero.Add(new Text(tipo).SetFont(boldFont));
                            prfPrimero.Add(", con un saldo pendiente a fecha de " + fechaCompra + " de ");
                            prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                            prfPrimero.Add(", cuyo origen es ");
                            prfPrimero.Add(new Text(origenBanco).SetFont(boldFont));

                            prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                            prfPrimero.Add(new Text("MCdos LEGAL S.L.").SetFont(boldFont));
                            prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                            prfPrimero.Add("Por la presente, le requerimos para que - ");
                            prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                            prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                            prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                            doc.Add(prfPrimero);

                            // RECUADRO - En el caso de C.Valenciana hay que bajarlo puntos mas

                            string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                            ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                            var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                                .SetRelativePosition(0, 5, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfRecuadro = new Paragraph();
                            prfRecuadro.Add(Recuadro);
                            doc.Add(prfRecuadro);

                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000)
                            {// ALICANTE // CASTELLON // VALENCIA
                                Paragraph prfRecuadroInteriorIzq = new Paragraph()
                                .SetPageNumber(1)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetFixedPosition(145, 249, 250)
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
                                .SetFixedPosition(330, 249, 250)
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
                                .SetFixedPosition(145, 273, 250)
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
                                    .SetFixedPosition(330, 273, 250)
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
                            prfSegundo.Add("\r\n\r\nTambién le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a ");
                            prfSegundo.Add(new Text("MCDOS LEGAL S.L.").SetFont(boldFont));
                            prfSegundo.Add(" en el teléfono ");
                            prfSegundo.Add(new Text("91 088 31 05").SetFont(boldFont));
                            prfSegundo.Add(" o en la dirección de correo electrónico ");

                            string link1 = "contencioso@fondos.mc2legal.es ";
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

                            prfTercero.Add(new Text("\r\n\r\n\r\nAxactor Invest 1, S.á.r.l.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                            doc.Add(prfTercero);

                            //-------------------------------------------------------------

                            string rutaFirmaIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Izq.jpg"));
                            ImageData imagenDataFirmaIzq = ImageDataFactory.Create(rutaFirmaIzq);
                            var firmaIzq = new iText.Layout.Element.Image(imagenDataFirmaIzq)
                                .SetPageNumber(2)
                                .SetRelativePosition(0, 0, 1, 1)
                                .SetMaxWidth(110);
                            Paragraph firmaPrfIzq = new Paragraph("");
                            firmaPrfIzq.Add(firmaIzq);
                            doc.Add(firmaPrfIzq);

                            string rutaFirmaAKCP = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firma_Dch.jpg"));
                            ImageData imagenDataFirmaAKCP = ImageDataFactory.Create(rutaFirmaAKCP);
                            var firmaAKCP = new iText.Layout.Element.Image(imagenDataFirmaAKCP)
                                .SetPageNumber(2)
                                .SetRelativePosition(270, -60, 1, 1)
                                .SetMaxWidth(180);
                            Paragraph firmaDch = new Paragraph("");
                            firmaDch.Add(firmaAKCP);
                            doc.Add(firmaDch);

                            //-------------------------------------------------------------

                            if (iCP >= 31000 && iCP < 32000)
                            {//NAVARRA
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
                                prfNavarra.Add(new Text("Axactor Invest 1, S.á.r.l.").SetFont(boldFont));
                                prfNavarra.Add(" y ");
                                prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de ");
                                prfNavarra.Add(costeCartera + "."); 
                                prfNavarra.Add(" Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfNavarra);
                            }
                            if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000) // ALICANTE // CASTELLON // VALENCIA
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
                                prfValencia.Add(new Text("Axactor Invest 1, S.á.r.l.").SetFont(boldFont));
                                prfValencia.Add(" cedio los creditos a ");
                                prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                                prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                                prfValencia.Add(new Text("Axactor Invest 1, S.á.r.l.").SetFont(boldFont));
                                prfValencia.Add(" domiciliada en 1B, rue Jean Piret, L-2350 Luxemburgo, Gran Ducado de Luxemburgoy Número de Identificación Fiscal con el número 256.458. ii) el Crédito se encuentra identificado en la página " + paginasInv + " del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de ");
                                prfValencia.Add(costeCartera + ".");
                                prfValencia.Add(" Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                                doc.Add(prfValencia);
                            }

                            //-------------------------------------------------------------

                            Paragraph prfCuarto = new Paragraph()
                            .SetPageNumber(2)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, -20, 0, 0)
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
                                .SetRelativePosition(0, -20, 0, 0)
                                .SetMaxWidth(455);
                            Paragraph prfTabla = new Paragraph();
                            prfTabla.Add(Tabla);
                            doc.Add(prfTabla);
                            count++;
                        }
                    }
                }                
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

    //"" - Quartz Capital Fund II - ORANGE
    /*public partial class Form1 : Form
    {
        ConexionDB conn = new ConexionDB();
        MCCommand mcComm = new MCCommand();
        Comp comp = new Comp();
        private OpenFileDialog openFileDialog;
        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        int count = 0;
        string idCliente = "91";

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

            OleDbConnection oleConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim() + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter("SELECT * FROM [" + hoja + "$]", oleConnection);

            DataSet ds = new DataSet();
            oleAdapter.Fill(ds);
            oleConnection.Close();
            DataTable dt = ds.Tables[0];

            foreach (DataRow fila in dt.Rows)
            {
                string expediente = fila["REF_PAGO"].ToString();
                string contrato = fila["REF_MC"].ToString();
                string refEnvio = fila["REF_ENVIO"].ToString();
                string importe = fila["IMPORTE"].ToString();
                string nombre = fila["NOMBRE"].ToString();
                string municipio = fila["MUNICIPIO"] != DBNull.Value ? fila["MUNICIPIO"].ToString() : "";
                string direccion = fila["DIRECCION_1"].ToString();
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

                string rutaArchivos = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Orange");
                string rutaArchivosGeneral = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Orange/General");
                string rutaArchivosNavarra = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Orange/Navarra");
                string rutaArchivosValencia = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Orange/Valencia");
                if (!Directory.Exists(rutaArchivos)) Directory.CreateDirectory(rutaArchivos);
                if (!Directory.Exists(rutaArchivosGeneral)) Directory.CreateDirectory(rutaArchivosGeneral);
                if (!Directory.Exists(rutaArchivosNavarra)) Directory.CreateDirectory(rutaArchivosNavarra);
                if (!Directory.Exists(rutaArchivosValencia)) Directory.CreateDirectory(rutaArchivosValencia);

                var exportarPDF = "";
                if (cp >= 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarra, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                else if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValencia, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");
                else exportarPDF = Path.Combine(rutaArchivosGeneral, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato + ".pdf");

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
                        prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 911088904\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
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
                            .SetFontSize(10.5f)
                            .SetFixedLeading(12);

                        prfPrimero.Add("\r\nReferencia del crédito: ");
                        prfPrimero.Add(new Text(contrato).SetFont(boldFont));                        

                        prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha 21 de febrero de 2024, ");
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
                        prfPrimero.Add("“), que ostenta frente a usted en su calidad de Titular, con un saldo pendiente a fecha de 21 de febrero de 2024 de ");
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
                            .SetFixedPosition(135, 278, 250)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(10.5f)
                            .SetFixedLeading(12);
                        prfRecuadroInteriorIzq.Add("REFERENCIA DEL PAGO:\r\n");
                        prfRecuadroInteriorIzq.Add("IMPORTE:\r\n");
                        prfRecuadroInteriorIzq.Add("CUENTA DE PAGO:");
                        doc.Add(prfRecuadroInteriorIzq);

                        Paragraph prfRecuadroInteriorDch = new Paragraph()
                            .SetPageNumber(1)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetFixedPosition(330, 278, 250)
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
                            .SetFontSize(10.5f)
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
                            .SetRelativePosition(0, -10, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(10.5f)
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

                        if (cp >= 31000 && cp < 32000) //NAVARRA
                        {
                            Paragraph prfNavarra = new Paragraph()
                                .SetPageNumber(2)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -20, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(10.5f)
                                .SetFixedLeading(12);

                            prfNavarra.Add("Asimismo, ");
                            prfNavarra.Add(new Text("QUARTZ CAPITAL FUND II").SetFont(boldFont));
                            prfNavarra.Add(" y ");
                            prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfNavarra.Add(", de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de trescientos mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                            doc.Add(prfNavarra);
                        }
                        if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
                        {
                            Paragraph prfValencia = new Paragraph()
                                .SetPageNumber(2)
                                .SetVerticalAlignment(VerticalAlignment.TOP)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetRelativePosition(0, -20, 0, 0)
                                .SetFontColor(colorNegro)
                                .SetFont(regularFont)
                                .SetFontSize(10.5f)
                                .SetFixedLeading(12);

                            prfValencia.Add("Asimismo, ");
                            prfValencia.Add(new Text("QUARTZ CAPITAL FUND II").SetFont(boldFont));
                            prfValencia.Add(" cedio los creditos a ");
                            prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) QUARTZ CAPITAL FUND S.C.A. – QUARTZ CAPITAL II, una Sociedad de Inversión de Capital Variable, Fondo de Inversión Especializado organizado bajo las leyes del Gran Ducado de Luxemburgo en forma de una sociedad en comandita por acciones, con sede social en 6A Rue Gabriel Lippman, L-5365 Schuttrange-Munsbach, Gran Ducado de Luxemburgo, registrada en el Registro de Comercio y Sociedades de Luxemburgo (Registre de Commerce et des Sociétés o RCS) con el número 167191, representada por su Socio General QUARTZ MANAGEMENT GP S.A.R.L, una sociedad de responsabilidad limitada privada de Luxemburgo, con domicilio social en 16 Rue d’Epernay L-1616 Luxemburgo, Gran Ducado de Luxemburgo, registrada en el Registro de Comercio y Sociedades de Luxemburgo con el número B 211.727 y titular del número de identificación fiscal español N0076022C ii) el Crédito se encuentra identificado en la página 21 a 41 del Contrato de Cesión de la Cartera de Créditos (Anexo III) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de trescientos mil euros. Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");                            
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
            string idCliente = "90";

            OleDbConnection oleConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim() + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter("SELECT * FROM [" + hoja + "$]", oleConnection);

            DataSet ds = new DataSet();
            oleAdapter.Fill(ds);
            oleConnection.Close();

            DataTable dt = ds.Tables[0];

            foreach (DataRow fila in dt.Rows)
            {
                string refMC = fila["REF_MC"].ToString();
                string contrato = fila["CONTRATO"].ToString();
                string refEnvio = fila["REF_ENVIO"].ToString();
                string origen = fila["ORIGEN"].ToString();
                string importe = fila["IMPORTE"].ToString();
                string nombre = fila["NOMBRE"].ToString();
                string municipio = fila["MUNICIPIO"] != DBNull.Value ? fila["MUNICIPIO"].ToString() : "";
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

                string rutaArchivosA = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Alerin");
                string rutaArchivosGeneralA = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Alerin/General");
                string rutaArchivosNavarraA = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Alerin/Navarra");
                string rutaArchivosValenciaA = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Alerin/Valencia");
                if (!Directory.Exists(rutaArchivosA)) Directory.CreateDirectory(rutaArchivosA);
                if (!Directory.Exists(rutaArchivosGeneralA)) Directory.CreateDirectory(rutaArchivosGeneralA);
                if (!Directory.Exists(rutaArchivosNavarraA)) Directory.CreateDirectory(rutaArchivosNavarraA);
                if (!Directory.Exists(rutaArchivosValenciaA)) Directory.CreateDirectory(rutaArchivosValenciaA);

                string rutaArchivosE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Axactor España");
                string rutaArchivosGeneralE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Axactor España/General");
                string rutaArchivosNavarraE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Axactor España/Navarra");
                string rutaArchivosValenciaE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Axactor España/Valencia");
                if (!Directory.Exists(rutaArchivosE)) Directory.CreateDirectory(rutaArchivosE);
                if (!Directory.Exists(rutaArchivosGeneralE)) Directory.CreateDirectory(rutaArchivosGeneralE);
                if (!Directory.Exists(rutaArchivosNavarraE)) Directory.CreateDirectory(rutaArchivosNavarraE);
                if (!Directory.Exists(rutaArchivosValenciaE)) Directory.CreateDirectory(rutaArchivosValenciaE);

                string rutaArchivosI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Axactor Invest");
                string rutaArchivosGeneralI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Axactor Invest/General");
                string rutaArchivosNavarraI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Axactor Invest/Navarra");
                string rutaArchivosValenciaI = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 3/Axactor Invest/Valencia");
                if (!Directory.Exists(rutaArchivosI)) Directory.CreateDirectory(rutaArchivosA);
                if (!Directory.Exists(rutaArchivosGeneralI)) Directory.CreateDirectory(rutaArchivosGeneralI);
                if (!Directory.Exists(rutaArchivosNavarraI)) Directory.CreateDirectory(rutaArchivosNavarraI);
                if (!Directory.Exists(rutaArchivosValenciaI)) Directory.CreateDirectory(rutaArchivosValenciaI);

                var exportarPDF = "";
                if (origen == "Alerin")
                {
                    if (cp >= 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraA, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                    else if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaA, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralA, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
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

                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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

                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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

                            if (cp >= 31000 && cp < 32000) //NAVARRA
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
                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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
                    if (cp >= 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraE, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                    else if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaE, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralE, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
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

                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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

                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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

                            if (cp >= 31000 && cp < 32000) //NAVARRA
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
                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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
                    if (cp >= 31000 && cp < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraI, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                    else if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaI, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                    else exportarPDF = Path.Combine(rutaArchivosGeneralI, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
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

                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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

                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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

                            if (cp >= 31000 && cp < 32000) //NAVARRA
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
                            if (cp >= 3000 && cp < 4000 || cp >= 12000 && cp < 13000 || cp >= 46000 && cp < 47000)// ALICANTE // CASTELLON // VALENCIA
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

    //PAGANTIS
    /*public partial class Form1 : Form
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
            string idCliente = "89";

            OleDbConnection oleConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim() + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter("SELECT * FROM [" + hoja + "$]", oleConnection);

            DataSet ds = new DataSet();
            oleAdapter.Fill(ds);
            oleConnection.Close();
            DataTable dt = ds.Tables[0];

            foreach (DataRow fila in dt.Rows)
            {
                string refEnvio = fila["REF_ENVIO"].ToString();
                string refMC = fila["CONTRATO"].ToString();
                string origenBanco = "Pagantis";
                string importe = fila["IMPORTE"].ToString();
                string nombre = fila["NOMBRE"].ToString();
                object municipio = fila["MUNICIPIO"] != DBNull.Value ? fila["MUNICIPIO"].ToString() : "";
                object direccion = fila["DIRECCION_1"] != DBNull.Value ? fila["DIRECCION_1"].ToString() : "";
                object cp = fila["CP"] != DBNull.Value ? (object)fila["CP"] : 0;
                string tipo = "Titular";
                string provincia = string.Empty; //fila["PROVINCIA"].ToString();
                string fechaEnvio = "29 de Noviembre de 2023";
                string fechaCompra = "20 de noviembre de 2023";
                //string paginasEsp = "179 a 197";
                //string costeCartera = "doscientos nueve mil novecientos noventa y seis con veinticinco euros";

                string[] nombreMinusculas = nombre.ToLower().Split(' ');
                for (int i = 0; i < nombreMinusculas.Length; i++) if (nombreMinusculas[i].Length > 2) nombreMinusculas[i] = char.ToUpper(nombreMinusculas[i][0]) + nombreMinusculas[i].Substring(1);
                string nombreFormateado = string.Join(" ", nombreMinusculas);
                if (nombreFormateado.Length > 59) nombreFormateado = nombreFormateado.Substring(0, 59);

                string[] palabrasLocalidad = municipio.ToString().ToLower().Split(' ');
                for (int i = 0; i < palabrasLocalidad.Length; i++) if (palabrasLocalidad[i].Length > 2) palabrasLocalidad[i] = char.ToUpper(palabrasLocalidad[i][0]) + palabrasLocalidad[i].Substring(1);
                string localidadFormateada = string.Join(" ", palabrasLocalidad);
                if (localidadFormateada.Length > 59) localidadFormateada = localidadFormateada.Substring(0, 59);

                //string[] palabrasDireccion = direccion.ToLower().Split(' ');
                //for (int i = 0; i < palabrasDireccion.Length; i++) if (palabrasDireccion[i].Length > 2) palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);

                string[] palabrasDireccion = direccion.ToString().ToLower().Split(' ');
                for (int i = 0; i < palabrasDireccion.Length; i++)
                {
                    if (palabrasDireccion[i].Length > 2)
                    {
                        palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                        if (palabrasDireccion[i] == "null" || palabrasDireccion[i] == "(null)") palabrasDireccion[i] = string.Empty;
                    }
                }
                string direccionFormateada = string.Join(" ", palabrasDireccion);
                if (direccionFormateada.Length > 59) direccionFormateada = direccionFormateada.Substring(0, 59);

                int iCP;
                if (int.TryParse(cp.ToString(), out iCP)) // Converimos cp a entero iCP

                if (provincia == string.Empty) if (int.TryParse(cp.ToString(), out iCP)) provincia = comp.CodigoPostal(iCP); // Si provincia es vacio y cp es entero, convertimos cp a provincia

                string sCP = cp.ToString() == string.Empty || cp.ToString() == "null" || cp.ToString() == "0" ? "" : cp.ToString(); // Si cp es vacio o 'null', ponemos vacio 

                var exportarPDF = "";
                string rutaArchivos = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Pagantis");

                exportarPDF = Path.Combine(rutaArchivos + "/General", "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");

                using (var writter = new PdfWriter(exportarPDF))
                {
                    using (var pdf = new PdfDocument(writter))
                    {
                        var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                        doc.SetMargins(50, 70, 50, 70);

                        string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                        string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                        PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                        PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                        iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                        iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                        // LOGOS - En la izquierda se apoya sobre los margenes del formato A4 y en la derecha ajustamos su posicion de manera libre

                        string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                        ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);
                        var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)
                            .SetRelativePosition(-2, 0, 0, 0)
                            .SetMaxWidth(70)
                            .SetMarginBottom(48);
                        Paragraph encabezadoIzq = new Paragraph("");
                        encabezadoIzq.Add(logoIzq);
                        doc.Add(encabezadoIzq);

                        string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Pagantis.jpg"));
                        ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                        var logoDch = new iText.Layout.Element.Image(imagenDataDch)
                            .SetFixedPosition(1, 385, 750)
                            .SetMaxWidth(140);
                        Paragraph encabezadoDch = new Paragraph("");
                        encabezadoDch.Add(logoDch);
                        doc.Add(encabezadoDch);

                        // ENCABEZADO - En la izq. los datos se mueven por libre

                        Paragraph prfMC2 = new Paragraph()
                            .SetPageNumber(1)
                            //.SetRelativePosition(0, 0, 0, 0)
                            .SetFixedPosition(70, 620, 750)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfMC2.Add("MCdos LEGAL S.L.");
                        prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 910883105\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                        doc.Add(prfMC2);

                        Paragraph prfDatosCliente = new Paragraph()
                            .SetPageNumber(1)
                            .SetTextAlignment(TextAlignment.RIGHT)
                            .SetRelativePosition(0, -24, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + sCP + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, " + fechaEnvio);
                        doc.Add(prfDatosCliente);

                        // PRIMER PARRAFO -------------------------------------------------------------

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
                        prfPrimero.Add(new Text(refMC).SetFont(boldFont));

                        prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha " + fechaCompra + ", ");
                        prfPrimero.Add(new Text("Pagamastarde S.L.").SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Cedente").SetFont(boldFont));
                        prfPrimero.Add("“) cedió a ");
                        prfPrimero.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Cesionario").SetFont(boldFont));
                        prfPrimero.Add("”) una cartera de créditos y, entre ellos, el crédito de referencia ");
                        prfPrimero.Add(new Text(refMC).SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Crédito").SetFont(boldFont));
                        prfPrimero.Add("“), que ostenta frente a usted en su calidad de ");
                        prfPrimero.Add(new Text(tipo).SetFont(boldFont));
                        prfPrimero.Add(", con un saldo pendiente a fecha de ");
                        prfPrimero.Add(new Text(fechaCompra).SetFont(boldFont));
                        prfPrimero.Add(" de ");
                        prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                        prfPrimero.Add(", cuyo origen es ");
                        prfPrimero.Add(new Text(origenBanco).SetFont(boldFont));

                        prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                        prfPrimero.Add(new Text("MC2 LEGAL S.L.").SetFont(boldFont));
                        prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                        prfPrimero.Add("Por la presente, le requerimos para que - ");
                        prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                        prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n");
                        doc.Add(prfPrimero);

                        string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                        ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                        var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                            .SetRelativePosition(0, 5, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfRecuadro = new Paragraph();
                        prfRecuadro.Add(Recuadro);
                        doc.Add(prfRecuadro);

                        Paragraph prfRecuadroInteriorIzq = new Paragraph()
                            .SetPageNumber(1)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetFixedPosition(145, 285, 250)
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
                            .SetFixedPosition(330, 285, 250)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                        prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                        prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                        doc.Add(prfRecuadroInteriorDch);

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
                        prfSegundo.Add("\r\n\r\nTambién le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a ");
                        prfSegundo.Add(new Text("MC2 LEGAL S.L.").SetFont(boldFont));
                        prfSegundo.Add(" en el teléfono ");
                        prfSegundo.Add(new Text("91 117 54 38").SetFont(boldFont));
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

                        prfTercero.Add(new Text("\r\n\r\n\r\nPagamastarde S.L.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                        doc.Add(prfTercero);

                        //-------------------------------------------------------------

                        string rutaFirmas = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firmas.jpg"));
                        ImageData imagenDataFirmas = ImageDataFactory.Create(rutaFirmas);
                        var Firmas = new iText.Layout.Element.Image(imagenDataFirmas)
                            .SetRelativePosition(0, -10, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfFirmas = new Paragraph();
                        prfFirmas.Add(Firmas);
                        doc.Add(prfFirmas);

                        //-------------------------------------------------------------                        

                        Paragraph prfCuarto = new Paragraph()
                        .SetPageNumber(2)
                        .SetVerticalAlignment(VerticalAlignment.TOP)
                        .SetTextAlignment(TextAlignment.JUSTIFIED)
                        .SetRelativePosition(0, -20, 0, 0)
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
                            .SetRelativePosition(0, -20, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfTabla = new Paragraph();
                        prfTabla.Add(Tabla);
                        doc.Add(prfTabla);
                        count++;
                    }
                }
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
    }*/

    //CRISALIDA 1 - AXACTOR LUXEMBURGO/AXACTOR ESPAÑA/AXACTOR INVEST - SANTANDER
    /*public partial class Form1 : Form
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
            string idCliente = "87";

            OleDbConnection oleConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim() + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter("SELECT * FROM [" + hoja + "$]", oleConnection);

            DataSet ds = new DataSet();
            oleAdapter.Fill(ds);
            oleConnection.Close();
            DataTable dt = ds.Tables[0];

            foreach (DataRow fila in dt.Rows)
            {
                string refEnvio = fila["REF"].ToString();
                string refMC = fila["CONTRATO"].ToString();
                string origenBanco = "Banco Santander";
                string importe = fila["IMPORTE"].ToString();
                string nombre = fila["NOMBRE"].ToString();
                object municipio = fila["MUNICIPIO"] != DBNull.Value ? fila["MUNICIPIO"].ToString() : "";
                object direccion = fila["DIRECCION_1"] != DBNull.Value ? fila["DIRECCION_1"].ToString() : "";
                object direccion2 = fila["DIRECCION_2"] != DBNull.Value ? fila["DIRECCION_2"].ToString() : "";
                direccion = direccion + " " + direccion2;
                object cp = fila["CP"] != DBNull.Value ? (object)fila["CP"] : 0;
                string tipo = fila["TIPO"].ToString();
                string provincia = fila["PROVINCIA"].ToString();
                string fechaEnvio = "31 de Julio de 2023";
                string fechaCompra = "20 de Julio de 2023";
                string paginasEsp = "179 a 197";
                string costeCartera = "doscientos nueve mil novecientos noventa y seis con veinticinco euros";

                string[] nombreMinusculas = nombre.ToLower().Split(' ');
                for (int i = 0; i < nombreMinusculas.Length; i++) if (nombreMinusculas[i].Length > 2) nombreMinusculas[i] = char.ToUpper(nombreMinusculas[i][0]) + nombreMinusculas[i].Substring(1);
                string nombreFormateado = string.Join(" ", nombreMinusculas);
                if (nombreFormateado.Length > 59) nombreFormateado = nombreFormateado.Substring(0, 59);

                string[] palabrasLocalidad = municipio.ToString().ToLower().Split(' ');
                for (int i = 0; i < palabrasLocalidad.Length; i++) if (palabrasLocalidad[i].Length > 2) palabrasLocalidad[i] = char.ToUpper(palabrasLocalidad[i][0]) + palabrasLocalidad[i].Substring(1);
                string localidadFormateada = string.Join(" ", palabrasLocalidad);
                if (localidadFormateada.Length > 59) localidadFormateada = localidadFormateada.Substring(0, 59);

                //string[] palabrasDireccion = direccion.ToLower().Split(' ');
                //for (int i = 0; i < palabrasDireccion.Length; i++) if (palabrasDireccion[i].Length > 2) palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);

                string[] palabrasDireccion = direccion.ToString().ToLower().Split(' ');
                for (int i = 0; i < palabrasDireccion.Length; i++){
                    if (palabrasDireccion[i].Length > 2) {
                        palabrasDireccion[i] = char.ToUpper(palabrasDireccion[i][0]) + palabrasDireccion[i].Substring(1);
                        if (palabrasDireccion[i] == "null" || palabrasDireccion[i] == "(null)") palabrasDireccion[i] = string.Empty;
                    }
                }
                string direccionFormateada = string.Join(" ", palabrasDireccion);
                if (direccionFormateada.Length > 59) direccionFormateada = direccionFormateada.Substring(0, 59);

                int iCP;
                if (int.TryParse(cp.ToString(), out iCP)) // Converimos cp a entero iCP

                if (provincia == string.Empty) if (int.TryParse(cp.ToString(), out iCP)) provincia = comp.CodigoPostal(iCP); // Si provincia es vacio y cp es entero, convertimos cp a provincia

                string sCP = cp.ToString() == string.Empty || cp.ToString() == "null" || cp.ToString() == "0" ? "" : cp.ToString(); // Si cp es vacio o 'null', ponemos vacio 

                var exportarPDF = "";
                string rutaArchivosE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 1");
                string rutaArchivosGeneralE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 1/General");
                string rutaArchivosNavarraE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 1/Navarra");
                string rutaArchivosValenciaE = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/Crisalida 1/Valencia");
                if (!Directory.Exists(rutaArchivosGeneralE)) Directory.CreateDirectory(rutaArchivosGeneralE);
                if (!Directory.Exists(rutaArchivosNavarraE)) Directory.CreateDirectory(rutaArchivosNavarraE);
                if (!Directory.Exists(rutaArchivosValenciaE)) Directory.CreateDirectory(rutaArchivosValenciaE);

                if (iCP >= 31000 && iCP < 32000) exportarPDF = Path.Combine(rutaArchivosNavarraE, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                else if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000) exportarPDF = Path.Combine(rutaArchivosValenciaE, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                else exportarPDF = Path.Combine(rutaArchivosGeneralE, "Hello_" + refEnvio + "_" + idCliente + "_" + refMC + ".pdf");
                string rutaArchivos = rutaArchivosE;
                string rutaArchivosNavarra = rutaArchivosNavarraE;
                string rutaArchivosValencia = rutaArchivosValenciaE;

                using (var writter = new PdfWriter(exportarPDF))
                {
                    using (var pdf = new PdfDocument(writter))
                    {
                        var doc = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);//Damos el formato de A4
                        doc.SetMargins(50, 70, 50, 70);

                        string rutaCalibri = @"C:\Windows\Fonts\calibri.ttf";
                        string rutaCalibriBold = @"C:\Windows\Fonts\calibrib.ttf";
                        PdfFont regularFont = PdfFontFactory.CreateFont(rutaCalibri, PdfEncodings.IDENTITY_H);
                        PdfFont boldFont = PdfFontFactory.CreateFont(rutaCalibriBold, PdfEncodings.IDENTITY_H);
                        iText.Kernel.Colors.Color colorNegro = new DeviceRgb(0, 0, 0);
                        iText.Kernel.Colors.Color colorAzul = new DeviceRgb(0, 0, 255);

                        // LOGOS - En la izquierda se apoya sobre los margenes del formato A4 y en la derecha ajustamos su posicion de manera libre

                        string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                        ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);
                        var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)
                            .SetRelativePosition(-2, 0, 0, 0)
                            .SetMaxWidth(70)
                            .SetMarginBottom(48);
                        Paragraph encabezadoIzq = new Paragraph("");
                        encabezadoIzq.Add(logoIzq);
                        doc.Add(encabezadoIzq);

                        string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Axactor.jpg"));
                        ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                        var logoDch = new iText.Layout.Element.Image(imagenDataDch)
                            .SetFixedPosition(1, 385, 750)
                            .SetMaxWidth(140);
                        Paragraph encabezadoDch = new Paragraph("");
                        encabezadoDch.Add(logoDch);
                        doc.Add(encabezadoDch);

                        // ENCABEZADO - En la izq. los datos se mueven por libre

                        Paragraph prfMC2 = new Paragraph()
                            .SetPageNumber(1)
                            //.SetRelativePosition(0, 0, 0, 0)
                            .SetFixedPosition(70, 620, 750)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfMC2.Add("MCdos LEGAL S.L.");
                        prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 910883105\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                        doc.Add(prfMC2);

                        Paragraph prfDatosCliente = new Paragraph()
                            .SetPageNumber(1)
                            .SetTextAlignment(TextAlignment.RIGHT)
                            .SetRelativePosition(0, -24, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfDatosCliente.Add(nombreFormateado + "\r\n" + direccionFormateada + "\r\n" + sCP + " " + localidadFormateada + "\r\n" + provincia + "\r\n\r\nMadrid, " + fechaEnvio);
                        doc.Add(prfDatosCliente);

                        // PRIMER PARRAFO -------------------------------------------------------------

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
                        prfPrimero.Add(new Text(refMC).SetFont(boldFont));

                        prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a:\r\n\r\nPor la presente le comunicamos que con fecha " + fechaCompra + ", ");
                        prfPrimero.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Cedente").SetFont(boldFont));
                        prfPrimero.Add("“) cedió a ");
                        prfPrimero.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Cesionario").SetFont(boldFont));
                        prfPrimero.Add("”) una cartera de créditos y, entre ellos, el crédito de referencia ");
                        prfPrimero.Add(new Text(refMC).SetFont(boldFont));
                        prfPrimero.Add(" (el “");
                        prfPrimero.Add(new Text("Crédito").SetFont(boldFont));
                        prfPrimero.Add("“), que ostenta frente a usted en su calidad de ");
                        prfPrimero.Add(new Text(tipo).SetFont(boldFont));
                        prfPrimero.Add(", con un saldo pendiente a fecha de 24 de julio de 2023 de ");
                        prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                        prfPrimero.Add(", cuyo origen es ");
                        prfPrimero.Add(new Text(origenBanco).SetFont(boldFont));

                        prfPrimero.Add("\r\n\r\nEl cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                        prfPrimero.Add(new Text("MCdos LEGAL S.L.").SetFont(boldFont));
                        prfPrimero.Add(" Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                        prfPrimero.Add("Por la presente, le requerimos para que - ");
                        prfPrimero.Add(new Text("en el plazo de 30 días naturales").SetFont(boldFont).SetUnderline());
                        prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                        prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont).SetUnderline());
                        doc.Add(prfPrimero);

                        string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                        ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                        var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                            .SetRelativePosition(0, 5, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfRecuadro = new Paragraph();
                        prfRecuadro.Add(Recuadro);
                        doc.Add(prfRecuadro);

                        Paragraph prfRecuadroInteriorIzq = new Paragraph()
                            .SetPageNumber(1)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetFixedPosition(145, 285, 250)
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
                            .SetFixedPosition(330, 285, 250)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfRecuadroInteriorDch.Add(new Text(refMC + "\r\n").SetFont(boldFont));
                        prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                        prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                        doc.Add(prfRecuadroInteriorDch);

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
                        prfSegundo.Add("\r\n\r\nTambién le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a ");
                        prfSegundo.Add(new Text("MCDOS LEGAL S.L.").SetFont(boldFont));
                        prfSegundo.Add(" en el teléfono ");
                        prfSegundo.Add(new Text("91 117 54 38").SetFont(boldFont));
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

                        prfTercero.Add(new Text("\r\n\r\n\r\nAXACTOR ESPAÑA, S.L.U.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tAKCP EUROPE SCSP").SetFont(boldFont));
                        doc.Add(prfTercero);

                        //-------------------------------------------------------------

                        string rutaFirmas = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firmas.jpg"));
                        ImageData imagenDataFirmas = ImageDataFactory.Create(rutaFirmas);
                        var Firmas = new iText.Layout.Element.Image(imagenDataFirmas)
                            .SetRelativePosition(0, -10, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfFirmas = new Paragraph();
                        prfFirmas.Add(Firmas);
                        doc.Add(prfFirmas);

                        //-------------------------------------------------------------

                        if (iCP >= 31000 && iCP < 32000)
                        {//NAVARRA
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
                            prfNavarra.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                            prfNavarra.Add(" y ");
                            prfNavarra.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfNavarra.Add(" de conformidad con la Ley 511, de la Ley 21/2019 de 4 de abril de modificación y actualización de la Compilación del Derecho Civil Foral de Navarra, le informan, en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de ");
                            prfNavarra.Add(costeCartera + ".");
                            prfNavarra.Add(" Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                            doc.Add(prfNavarra);
                        }
                        if (iCP >= 3000 && iCP < 4000 || iCP >= 12000 && iCP < 13000 || iCP >= 46000 && iCP < 47000)
                        {// ALICANTE // CASTELLON // VALENCIA
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
                            prfValencia.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                            prfValencia.Add(" cedio los creditos a ");
                            prfValencia.Add(new Text("AKCP Europe SCSp").SetFont(boldFont));
                            prfValencia.Add(", de conformidad con el Decreto Legislativo 1/2019, de 13 de diciembre, del Consell, de aprobación del texto refundido de la Ley del Estatuto de las personas consumidoras y usuarias de la Comunitat Valencia, le informan de los siguientes extremos: i ) que ");
                            prfValencia.Add(new Text("AXACTOR ESPAÑA, S.L.U.").SetFont(boldFont));
                            prfValencia.Add(" está domiciliada en Madrid (28007) Calle Doctor Esquerdo 136, 4ª planta, constituido el 09 de Julio de 2015 e inscrita en el Registro Mercantil de Madrid, al Tomo 33.781, Folio 113, Sección 8ª, Hoja M-607982, ii) el Crédito se encuentra identificado en las páginas " + paginasEsp + " del Contrato de Cesión de la Cartera de Créditos (Anexo V) con el código de identificación arriba referenciado y iii) en cuanto a la cesión de la cartera de créditos antes indicada, que los Créditos se cedieron, junto con otros créditos de características similares y que forman parte de la misma cartera, por un importe alzado de ");
                            prfValencia.Add(costeCartera + ".");
                            prfValencia.Add(" Dado que la operación consiste en una transmisión global de créditos, el precio de la operación es fijo, conjunto, global y único, sin que sea posible realizar una individualización de dicho precio.");
                            doc.Add(prfValencia);
                        }

                        //-------------------------------------------------------------

                        Paragraph prfCuarto = new Paragraph()
                        .SetPageNumber(2)
                        .SetVerticalAlignment(VerticalAlignment.TOP)
                        .SetTextAlignment(TextAlignment.JUSTIFIED)
                        .SetRelativePosition(0, -20, 0, 0)
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
                            .SetRelativePosition(0, -20, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfTabla = new Paragraph();
                        prfTabla.Add(Tabla);
                        doc.Add(prfTabla);
                        count++;
                    }                
                }
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
    }*/

    //BENKI
    /*public partial class Form1 : Form
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
            string idCliente = "84"; //BENKI

            OleDbConnection oleConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtFichero.Text.Trim() + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';");
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter("SELECT * FROM [" + hoja + "$]", oleConnection);

            DataSet ds = new DataSet();
            oleAdapter.Fill(ds);
            oleConnection.Close();
            DataTable dt = ds.Tables[0];

            foreach (DataRow fila in dt.Rows)
            {
                string refEnvio = fila["REF_ENVIO"].ToString();
                string contrato = fila["CONTRATO"].ToString();
                string importe = fila["IMPORTE"].ToString();
                string nombre = fila["NOMBRE"].ToString();
                string fechaEnvio = "27 de Marzo de 2023";
                string fechaCompra = "13 de abril de 2023";

                string[] nombreMinusculas = nombre.ToLower().Split(' ');
                for (int i = 0; i < nombreMinusculas.Length; i++) if (nombreMinusculas[i].Length > 2) nombreMinusculas[i] = char.ToUpper(nombreMinusculas[i][0]) + nombreMinusculas[i].Substring(1);
                string nombreFormateado = string.Join(" ", nombreMinusculas);
                if (nombreFormateado.Length > 59) nombreFormateado = nombreFormateado.Substring(0, 59);                

                var exportarPDF = "";
                string rutaArchivos = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/2023-04-13 Benki");
                string rutaArchivosGeneral = Path.Combine(ruta, "OneDrive - MC PROINVEST, S.L/Proyectos/HelloLetters/ArborKnot/2023-04-13 Benki/General");
                if (!Directory.Exists(rutaArchivosGeneral)) Directory.CreateDirectory(rutaArchivosGeneral);

                else exportarPDF = Path.Combine(rutaArchivosGeneral, "Hello_" + refEnvio + "_" + idCliente + "_" + contrato +".pdf");

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

                        // LOGOS - En la izquierda se apoya sobre los margenes del formato A4 y en la derecha ajustamos su posicion de manera libre

                        string rutaLogoIzq = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Arborknot.jpg"));
                        ImageData imagenDataIzq = ImageDataFactory.Create(rutaLogoIzq);
                        var logoIzq = new iText.Layout.Element.Image(imagenDataIzq)
                            .SetRelativePosition(-4, 0, 0, 0)
                            .SetMaxWidth(70)
                            .SetMarginBottom(48);
                        Paragraph encabezadoIzq = new Paragraph("");
                        encabezadoIzq.Add(logoIzq);
                        doc.Add(encabezadoIzq);

                        string rutaLogoDch = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Logo_Benki.jpg"));
                        ImageData imagenDataDch = ImageDataFactory.Create(rutaLogoDch);
                        var logoDch = new iText.Layout.Element.Image(imagenDataDch)
                            .SetFixedPosition(1, 370, 740)
                            .SetMaxWidth(160);
                        Paragraph encabezadoDch = new Paragraph("");
                        encabezadoDch.Add(logoDch);
                        doc.Add(encabezadoDch);

                        // ENCABEZADO - En la izq. los datos se mueven por libre

                        Paragraph prfMC2 = new Paragraph()
                            .SetPageNumber(1)
                            .SetRelativePosition(0, 0, 0, 0)
                            //.SetFixedPosition(70, 620, 750)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfMC2.Add("Prejudicial S.L.");
                        prfMC2.Add(new Text("\r\nC/ Goya 115 Pl 5 5º izq MADRID\r\nTELÉFONO DE CONTACTO: 910883105\r\nHORARIO DE ATENCIÓN AL CLIENTE:\r\nL- J: 09:00h a 18:00h\r\nV: 09:00h a 15:00h\r\n").SetFont(regularFont));
                        doc.Add(prfMC2);                        

                        // PRIMER PARRAFO -------------------------------------------------------------

                        Paragraph prfPrimero = new Paragraph()
                            .SetPageNumber(1)
                            .SetVerticalAlignment(VerticalAlignment.TOP)
                            .SetTextAlignment(TextAlignment.JUSTIFIED)
                            .SetRelativePosition(0, 0, 0, 0)
                            .SetFontColor(colorNegro)
                            .SetFont(regularFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);

                        prfPrimero.Add("\r\n\r\n\r\n\r\nReferencia del crédito: ");
                        prfPrimero.Add(new Text(contrato).SetFont(boldFont));

                        prfPrimero.Add("\r\n\r\nMuy Sr./a. nuestro/a: ");
                        prfPrimero.Add(new Text(nombre).SetFont(boldFont));
                        prfPrimero.Add("\r\n\r\nPor la presente le comunicamos que con fecha " + fechaEnvio + ", ");
                        prfPrimero.Add(new Text("BENKI DIGITAL LENDING S.L.U.").SetFont(boldFont));
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
                        prfPrimero.Add(new Text(fechaCompra).SetFont(boldFont));
                        prfPrimero.Add(" de ");
                        prfPrimero.Add(new Text(importe + " €").SetFont(boldFont));
                        prfPrimero.Add(", cuyo origen es ");
                        prfPrimero.Add(new Text("MILINEA.").SetFont(boldFont));

                        prfPrimero.Add("\r\n\r\nEl Cesionario, a su vez, ha encomendado la gestión de cobro del Crédito a ");
                        prfPrimero.Add(new Text("Prejudicial S.L.").SetFont(boldFont));
                        prfPrimero.Add(" (la “");
                        prfPrimero.Add(new Text("Agencia de Cobro").SetFont(boldFont));
                        prfPrimero.Add("“). Como consecuencia de lo anterior, y desde la fecha de esta comunicación, cualquier pago del Crédito deberá realizarlo al Cesionario en la cuenta bancaria, reflejando la referencia, que le indicamos abajo, careciendo de efectos liberatorios cualquier pago que realice en adelante a favor del Cedente conforme al artículo 1.527 del Código Civil.\r\n\r\n");

                        prfPrimero.Add("Por la presente, le requerimos para que - ");
                        prfPrimero.Add(new Text("en el plazo de 10 días naturales").SetFont(boldFont).SetUnderline());
                        prfPrimero.Add(" - proceda al pago de las cantidades que Ud. nos adeuda y que indicamos a continuación. ");
                        prfPrimero.Add(new Text("Asimismo, le facilitamos los datos bancarios donde a partir de la fecha de la presente notificación deberá realizar el ingreso a favor del Cesionario, indicando la siguiente REFERENCIA DEL PAGO:\r\n\r\n").SetFont(boldFont));
                        doc.Add(prfPrimero);

                        string rutaRecuadro = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Recuadro.png"));
                        ImageData imagenDataRecuadro = ImageDataFactory.Create(rutaRecuadro);
                        var Recuadro = new iText.Layout.Element.Image(imagenDataRecuadro)
                            .SetRelativePosition(0, 5, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfRecuadro = new Paragraph();
                        prfRecuadro.Add(Recuadro);
                        doc.Add(prfRecuadro);

                        Paragraph prfRecuadroInteriorIzq = new Paragraph()
                        .SetPageNumber(1)
                        .SetVerticalAlignment(VerticalAlignment.TOP)
                        .SetFixedPosition(145, 237, 250)
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
                            .SetFixedPosition(330, 237, 250)
                            .SetFontColor(colorNegro)
                            .SetFont(boldFont)
                            .SetFontSize(11)
                            .SetFixedLeading(12);
                        prfRecuadroInteriorDch.Add(new Text(contrato + "\r\n").SetFont(boldFont));
                        prfRecuadroInteriorDch.Add(new Text(importe + " €\r\n").SetFont(boldFont));
                        prfRecuadroInteriorDch.Add("ES6021008641630200187714");
                        doc.Add(prfRecuadroInteriorDch);                        

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
                        prfSegundo.Add("\r\n\r\nTambién le ofrecemos la posibilidad de adaptar el pago a sus circunstancias particulares gracias a los planes de pago personalizados que le ofrecerán nuestros gestores telefónicos.");
                        prfSegundo.Add("\r\n\r\nPara cualquier comunicación relativa al Crédito deberá dirigirse a la Agencia de Cobro en el teléfono ");
                        prfSegundo.Add(new Text("91 108 89 04").SetFont(boldFont));
                        prfSegundo.Add(" o en la dirección de correo electrónico ");

                        string link1 = "contencioso@arborknot.prejudicial.es.";
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

                        prfTercero.Add("Por último, AKCP Europe SCSp ");
                        prfTercero.Add(new Text("(“Arbor Knot”)").SetFont(boldFont));
                        prfTercero.Add(" le informa de que sus datos de carácter personal estarán sujetos a su política de protección de datos que puede consultar en la siguiente dirección web: ");

                        string link2 = "https://arborknot.com/privacy-policy/";
                        PdfLinkAnnotation linkAnnotation2 = new PdfLinkAnnotation(new iText.Kernel.Geom.Rectangle(0, 0, 0, 0));
                        linkAnnotation2.SetAction(PdfAction.CreateURI(link2));
                        linkAnnotation2.SetBorder(new PdfArray());
                        Link linkHtml1 = new Link(link2, linkAnnotation2);
                        prfTercero.Add(linkHtml1.SetFontColor(colorAzul).SetUnderline());

                        prfTercero.Add(". Puede encontrar un breve resumen sobre la cesión de sus datos, así como su tratamiento por Arbor Knot en el pie de la presente comunicación. Tenga en cuenta que algunos de los tratamientos descritos en la citada política pueden estar sujetos a su previa conformidad en el momento en que usted facilite voluntariamente datos de carácter personal adicionales para valorar alternativas que le permitan mejorar su capacidad financiera.");

                        prfTercero.Add("\r\n\r\nSin otro particular, aprovechamos la ocasión para saludarle atentamente.");

                        prfTercero.Add(new Text("\r\n\r\n\r\nBENKI DIGITAL LENDING, S.L.U.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t  AKCP EUROPE SCSP").SetFont(boldFont));
                        doc.Add(prfTercero);

                        //-------------------------------------------------------------

                        string rutaFirmas = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), rutaArchivos, "Firmas.jpg"));
                        ImageData imagenDataFirmas = ImageDataFactory.Create(rutaFirmas);
                        var Firmas = new iText.Layout.Element.Image(imagenDataFirmas)
                            .SetRelativePosition(0, 5, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfFirmas = new Paragraph();
                        prfFirmas.Add(Firmas);
                        doc.Add(prfFirmas);

                        //-------------------------------------------------------------

                        Paragraph prfCuarto = new Paragraph()
                        .SetPageNumber(2)
                        .SetVerticalAlignment(VerticalAlignment.TOP)
                        .SetTextAlignment(TextAlignment.JUSTIFIED)
                        .SetRelativePosition(0, 20, 0, 0)
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
                            .SetRelativePosition(0, 25, 0, 0)
                            .SetMaxWidth(455);
                        Paragraph prfTabla = new Paragraph();
                        prfTabla.Add(Tabla);
                        doc.Add(prfTabla);
                        count++;
                    }
                }
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
    */
}

