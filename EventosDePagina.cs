using iTextSharp.text;
using iTextSharp.text.pdf;

namespace GeradorDeRelatoriosEmPDF
{
    public class EventosDePagina : PdfPageEventHelper
    {
        private PdfContentByte wdc;
        private BaseFont fonteBaseRodape { get; set; }
        private Font fonteRodape { get; set; }

        public int totalPaginas { get; set; }

        public EventosDePagina(int totalPaginas)
        {
            //Inicializando as fontes do Rodapé da pagina
            fonteBaseRodape = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            fonteRodape = new Font(fonteBaseRodape, 8f, Font.NORMAL, BaseColor.Black);
            this.totalPaginas = totalPaginas;
        }

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            base.OnOpenDocument(writer, document);
            wdc = writer.DirectContent;
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            AdicionarMomentoGeracaoDeRelatorio(writer, document);
            AdicionarNumeroDePaginas(writer, document);
        }

        private void AdicionarMomentoGeracaoDeRelatorio(PdfWriter writer, Document document)
        {
            var textoMomentoGeracao = $"Gerado em {DateTime.Now.ToShortDateString()} Às {DateTime.Now.ToShortTimeString()}";

            wdc.BeginText();
            wdc.SetFontAndSize(fonteRodape.BaseFont, fonteRodape.Size);
            wdc.SetTextMatrix(document.LeftMargin, document.BottomMargin * 0.75f);
            wdc.ShowText(textoMomentoGeracao);
            wdc.EndText();
        }

        private void AdicionarNumeroDePaginas(PdfWriter writer, Document document)
        {
            int paginaAtual = writer.PageNumber;
            var textoPaginacao = $"Página {paginaAtual} de {totalPaginas}";

            float larguraTextoPaginacao = fonteBaseRodape.GetWidthPoint(textoPaginacao, fonteRodape.Size);
            var tamanhoPagina = document.PageSize;

            wdc.BeginText();
            wdc.SetFontAndSize(fonteRodape.BaseFont, fonteRodape.Size);
            wdc.SetTextMatrix(tamanhoPagina.Width - document.RightMargin - larguraTextoPaginacao, document.BottomMargin * 0.75f);
            wdc.ShowText(textoPaginacao);
            wdc.EndText();

        }

    }
}
