using Microsoft.Office.Interop.Excel;

namespace ConvertListToExcel.Configuracoes
{
    public class EstiloFonteConfig
    {
        public int? TamanhoFonte = 12;
        public string? FonteStyle = "Arial";
        public string CorFonte = "Black";
        public bool FonteNegrito = false;
        public bool FonteItalica = false;
        public bool FonteSublinhada = false;
        public bool Bordas = true;
        public string CorFundo = "White";
        public XlHAlign AlinhamentoHorizontal = XlHAlign.xlHAlignLeft;
        public XlVAlign AlinhamentoVertical = XlVAlign.xlVAlignCenter;
    }
}
