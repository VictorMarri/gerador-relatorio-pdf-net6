﻿using GeradorDeRelatoriosEmPDF;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Diagnostics;
using System.Text.Json;
class Program
{
    static List<Pessoa> pessoas = new List<Pessoa>();
    static BaseFont fonteBaseTextosRelatorio = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);

    static void Main(string[] args)
    {
        DeserializarPessoas();
        GerarRelatorioEmPDF(100);
    }
    static void DeserializarPessoas()
    {
        if (File.Exists("C:\\Users\\victor.marri\\source\\repos\\ProjetosPessoais\\GeradorDeRelatoriosEmPDF\\pessoas.json"))
        {
            using (var sr = new StreamReader("C:\\Users\\victor.marri\\source\\repos\\ProjetosPessoais\\GeradorDeRelatoriosEmPDF\\pessoas.json"))
            {
                var dados = sr.ReadToEnd();
                pessoas = JsonSerializer.Deserialize(dados, typeof(List<Pessoa>)) as List<Pessoa>;
            }
        }
    }

    static void GerarRelatorioEmPDF(int quantidadePessoas)
    {
        var pessoasSelecionadas = pessoas.Take(quantidadePessoas).ToList(); //Pegando a quantidade indicada de pessoas de dentro dessa lista e convertendo pra ToList()

        if (pessoasSelecionadas.Any())
        {
            //Configurando o documento PDF
            var pixelsPorMilimetro = 72 / 25.2F; //Ajustando a resolução do PDF. Sabemos que é 72dpi por padrão 
            var tamanhoPagina = PageSize.A4;
            var margemEsquerda = 15 * pixelsPorMilimetro;
            var margemDireita = 15 * pixelsPorMilimetro;
            var margemSuperior = 15 * pixelsPorMilimetro;
            var margemInferior = 20 * pixelsPorMilimetro;
            var pdf = new Document(tamanhoPagina, margemEsquerda, margemDireita, margemSuperior, margemInferior);

            var nomeArquivo = $"pessoas.{DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss")}.pdf";

            var arquivo = new FileStream(nomeArquivo, FileMode.Create); //Criando um novo arquivo
            var writer = PdfWriter.GetInstance(pdf, arquivo); //Associamos o documento pdf que estamos criando, ao arquivo que criamos. Portantom tudo que fizemos no documento PDF será salvo nesse arquivo
            pdf.Open(); //Inicializa o objeto pra ele começar a receber conteudo no PDF


            //Adicionando o titulo
            var fonteParagrafo = new iTextSharp.text.Font(fonteBaseTextosRelatorio, 32, iTextSharp.text.Font.NORMAL, BaseColor.Black);
            var titulo = new Paragraph("Relatório de Pessoas\n\n", fonteParagrafo);
            titulo.Alignment = Element.ALIGN_LEFT; //Ajustando o titulo à esquerda do docimento
            pdf.Add(titulo);

            //Adicionando imagem
            var caminhoImagem = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "C:\\Users\\victor.marri\\source\\repos\\ProjetosPessoais\\GeradorDeRelatoriosEmPDF\\img\\youtube.png");

            if (File.Exists(caminhoImagem))
            {
                iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(caminhoImagem);

                //Redimensionando a imagem mantendo as proporções
                float razaoAlturaLargura = logo.Width / logo.Height;
                float alturaLogo = 32;
                float larguraLogo = alturaLogo * razaoAlturaLargura;
                logo.ScaleToFit(larguraLogo, alturaLogo);

                var margemEsquerdaLogo = pdf.PageSize.Width - pdf.RightMargin - larguraLogo; ;//Alinhando a margem esquerda da imagem alinhado com a margem direita do documento
                var margemTopoLogo = pdf.PageSize.Height - pdf.TopMargin - 54;
                logo.SetAbsolutePosition(margemEsquerdaLogo, margemTopoLogo);
                writer.DirectContent.AddImage(logo, false);
            }

            //Adicionando a tabela de dados
            var tabela = new PdfPTable(5); //Gerando uma tabela com 5 colunas
            tabela.DefaultCell.BorderWidth = 0; //Definindo que essa tabela não vai ter borda
            tabela.WidthPercentage = 100; //Essa tabela vai ocupar 100% da largura disponivel da pagina, respeitando as margens direita e equerda

            //Adicionando celulas de titulos das colunas
            CriarCelulaTexto(tabela, "Código", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Nome", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Profissão", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Salário", PdfPCell.ALIGN_CENTER, true);
            CriarCelulaTexto(tabela, "Empregado", PdfPCell.ALIGN_CENTER, true);

            pdf.Add(tabela);

            pdf.Close();
            arquivo.Close();

            //Abrindo o PDF no visualizador padrão do Sistema operacional
            var caminhoPDF = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nomeArquivo); //Criando o caminho completo do docmento
            if (File.Exists(caminhoPDF))
            {
                //Esse comando em linha de comando vai fazer o PDF abrir automaticamente no nosso sistema operacional, e sem mostrar a janela de command
                Process.Start(new ProcessStartInfo()
                {
                    Arguments = $"/c start {caminhoPDF}",
                    FileName = "cmd.exe",
                    CreateNoWindow = true,
                });
            }


        }
    }

    static void CriarCelulaTexto(PdfPTable tabela,
                                 string texto,
                                 int alinhamentoHoriz = PdfPCell.ALIGN_LEFT,
                                 bool negrito = false,
                                 bool italico = false,
                                 int tamanhoFonte = 12,
                                 int alturaCelula = 25)
    {
        int estilo = iTextSharp.text.Font.NORMAL;
        if(negrito && italico)
        {
            estilo = iTextSharp.text.Font.BOLDITALIC;
        }
        else if (negrito)
        {
            estilo = iTextSharp.text.Font.BOLD;
        }
        else if (italico)
        {
            estilo = iTextSharp.text.Font.ITALIC;
        }

        var fonteCelula = new iTextSharp.text.Font(fonteBaseTextosRelatorio, tamanhoFonte, estilo, BaseColor.Black);

        var celula = new PdfPCell(new Phrase(texto, fonteCelula));
        celula.HorizontalAlignment = alinhamentoHoriz;
        celula.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
        celula.Border = 0;
        celula.BorderWidthBottom = 1;
        celula.FixedHeight = alturaCelula;
        tabela.AddCell(celula);
    }
}