using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creacion_PDF_HelloLetter
{
    public class Comp
    {
        public string CodigoPostal(int cp)
        {
            string provincia = "";
            if (cp <= 1000) return provincia = "";
            else if (cp > 1000 && cp < 2999) return provincia = "ÁLAVA";
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
    }
}
