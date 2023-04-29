using GeradorRecibo.Model;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using OfficeOpenXml;
using System.Globalization;
using System.Text;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

int totalCols = 27;
int cabecalho = 3;
int conteudo = 4;

int sequencial = 1;
string endereco = "";

var dataAtual = DateTime.Today;
var cultureInfo = new CultureInfo("pt-BR");
var data = $@"{dataAtual.Day} de {dataAtual.ToString("MMMM", cultureInfo)} de {dataAtual.Year}";

var caminhoPdf = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
var path = @$"../../../PlanilhaPagamento-{DateTime.Today.Year}.xlsx";
List<MoradorModel> list = new List<MoradorModel>();

void Pausa()
{
    Console.WriteLine("APERTE QUALQUER TECLA PARA CONTINUAR...");
    //Console.ReadKey();
}

Main();

void Main()
{
    Console.WriteLine("CARREGANDO...");
    LerArquivo();
    //GerarRecibos();
}

void LerArquivo()
{
    Console.WriteLine("INICIANDO LEITURA DO ARQUIVO...");
    Pausa();

    using (MemoryStream ms = new MemoryStream())
    {
        using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
        {
            byte[] bytes = new byte[file.Length];
            file.Read(bytes, 0, (int)file.Length);
            ms.Write(bytes, 0, (int)bytes.Length);

            using(var package = new ExcelPackage(ms))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                //BASE ENDERECO
                endereco = worksheet.Cells[1, 2].Value.ToString();
                //SEQUENCIAL
                sequencial = int.Parse(worksheet.Cells[2, 2].Value.ToString());

                for (int row = conteudo; row < worksheet.Dimension.Rows; row++)
                {
                    MoradorModel moradorModel = new MoradorModel();
                    MesesModel mesModel = new MesesModel();
                    moradorModel.Meses = new List<MesesModel>();

                    for (int i = 1; i <= totalCols; i ++)
                    {
                        var value = worksheet.Cells[row, i].Value;

                        if(i == 1)
                        {
                            if (value == null)
                                break;
                            moradorModel.Id = int.Parse(value.ToString());
                        }
                        else if(i == 2)
                        {
                            moradorModel.Casa = value.ToString();
                        }
                        else if (i == 3)
                        {
                            moradorModel.Morador = value.ToString();
                        }
                        else if (i == 4)
                        {
                            moradorModel.Email = value == null ? "" : value.ToString();
                        }
                        else
                        {
                            if (i % 2 != 0)
                            {
                                var mes = worksheet.Cells[cabecalho, i].Value;
                                mesModel.Mes = mes.ToString();
                                mesModel.Pago = value == null ? null : value.ToString();
                            }
                            else
                            {
                                mesModel.Gerado = value == null ? false : true;
                                moradorModel.Meses.Add(mesModel);
                                mesModel = new MesesModel();
                            }
                        }
                    }

                    if(moradorModel.Id > 0)
                        list.Add(moradorModel);
                }

                Console.WriteLine("ARQUIVO CARREGADO!");                
            }
        }

        ms.Close();
    }
}

void GerarRecibos()
{
    Console.WriteLine("INICIANDO GERACAO DOS RECIBOS...");
    Pausa();

    foreach(var morador in list.OrderBy(x => x.Id))
    {
        //CRIACAO DA PASTA DA CASA E ANO REFERENTE
        var caminhoGravacaoBase = $@"{caminhoPdf}\MatupaRecibos\{morador.Casa}\{DateTime.Today.Year}";
        
        if(!Directory.Exists(caminhoGravacaoBase))
            Directory.CreateDirectory(caminhoGravacaoBase);

        Console.WriteLine($@"MORADOR: {morador.Morador}");
        Console.WriteLine($@"CASA: {morador.Casa}");

        var mesesPagos = morador.Meses.Where(x => x.Pago != null && !x.Gerado);
        var qtdMesesPagos = mesesPagos.Count();
        var enderecoGravar = endereco.Replace("X", morador.Casa);
        var texto = "";

        if (qtdMesesPagos > 1)
        {
            var valor = decimal.Parse(mesesPagos.LastOrDefault().Pago == "X" ? mesesPagos.FirstOrDefault().Pago : mesesPagos.LastOrDefault().Pago);
            if (qtdMesesPagos > 2)
            {
                texto = $@"manutenção de {mesesPagos.FirstOrDefault().Mes} até {mesesPagos.LastOrDefault().Mes}";
            }
            else
            {
                texto = $@"manutenção de {mesesPagos.FirstOrDefault().Mes} e {mesesPagos.LastOrDefault().Mes}";
            }

            var caminhoReferencia = $@"{caminhoGravacaoBase}\Recibo-{mesesPagos.FirstOrDefault().Mes}-{mesesPagos.LastOrDefault().Mes}.pdf";

            GeradorPDF(caminhoReferencia, sequencial, valor, morador.Morador, EscreverExtenso(valor), enderecoGravar, texto, data);
        }
        else if (qtdMesesPagos == 1)
        {
            var caminhoReferencia = $@"{caminhoGravacaoBase}\Recibo-{mesesPagos.FirstOrDefault().Mes}.pdf";

            texto = $@"manutenção de {mesesPagos.FirstOrDefault().Mes}";

            var valor = decimal.Parse(mesesPagos.FirstOrDefault().Pago);

            GeradorPDF(caminhoReferencia, sequencial, valor, morador.Morador, EscreverExtenso(valor), enderecoGravar, texto, data);
        }
    }
}

void GeradorPDF(string caminho, int numeroRecibo, decimal valor, string morador, string valorExtenso, string endereco, string texto, string data)
{
    #region CRIACAO DO ARQUIVO
    Document doc = new Document(PageSize.A6.Rotate());
    doc.SetMargins(10, 10, 10, 10);
    FileStream fs = new FileStream(caminho, FileMode.Create, FileAccess.Write);
    PdfWriter writer = PdfWriter.GetInstance(doc, fs);
    doc.Open();
    #endregion

    var css = "";
    css += ".corpo { padding: 10px; }";
    css += ".space { padding: 20px; }";
    css += ".campo { border: 2px solid rgb(54, 161, 223); background-color: rgb(110, 193, 241); color: white !important; padding: 10px; }";
    css += ".fonte { font-size: 14px; color: rgb(54, 161, 223); line-height: 40px; }";
    css += ".start { text-align: start; }";
    css += "td { padding: 15 0 15 0; }";
    //css += "tbody { padding: 10 }";
    //css += "table { border: 1px solid rgb(54, 161, 223) }";

    StringBuilder sb = new StringBuilder();
    sb.AppendLine("<div class=\"corpo\">");
    sb.AppendLine("<table style=\"width: 100%\">");
    sb.AppendLine("<tbody>");
    sb.AppendLine("<tr class=\"fonte\">");
    sb.AppendLine("<td class=\"campo\"> <span>Recibo: </span> <span>{-numero-}</span> </td>");
    sb.AppendLine("<td class=\"space\"></td>");
    sb.AppendLine("<td class=\"start campo\"> <span>Valor:</span> <span>{-valor-}</span> </td>");
    sb.AppendLine("</tr>");
    sb.AppendLine("<tr class=\"fonte\">");
    sb.AppendLine("<td colspan=\"3\"> <span>Recebi(emos) de:</span> <span>{-morador-}</span> </td>");
    sb.AppendLine("</tr>");
    sb.AppendLine("<tr class=\"fonte\">");
    sb.AppendLine("<td colspan=\"3\"> <span>Valor de:</span> <span>{-valorExtenso-}</span> </td>");
    sb.AppendLine("</tr>");
    sb.AppendLine("<tr class=\"fonte\">");
    sb.AppendLine("<td colspan=\"3\"> <span>Endereço: </span> <span>{-endereco-}</span> </td>");
    sb.AppendLine("</tr>");
    sb.AppendLine("<tr class=\"fonte\">");
    sb.AppendLine("<td colspan=\"3\"> <span>Correspondente a </span> <span>{-texto-}</span> <span>e para clareza firmo(amos) o presente.</span> </td>");
    sb.AppendLine("</tr>");
    sb.AppendLine("<tr class=\"fonte\">");
    sb.AppendLine("<td colspan=\"3\"> <span>{-data-}</span> </td>");
    sb.AppendLine("</tr>");
    sb.AppendLine("</tbody>");
    sb.AppendLine("</table>");
    sb.AppendLine("</div>");

    sb.Replace("{-numero-}", numeroRecibo.ToString());
    sb.Replace("{-valor-}", valor.ToString());
    sb.Replace("{-morador-}", morador);
    sb.Replace("{-valorExtenso-}", valorExtenso + " REAIS");
    sb.Replace("{-endereco-}", endereco);
    sb.Replace("{-texto-}", texto);
    sb.Replace("{-data-}", data);

    #region FINALIZACAO DO ARQUIVO
    XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString())), new MemoryStream(Encoding.UTF8.GetBytes(css.ToString())));
    doc.Close();
    #endregion
}

#region HELPERS
static string EscreverExtenso(decimal valor)
{
    if (valor <= 0 | valor >= 1000000000000000)
        return "Valor não suportado pelo sistema.";
    else
    {
        string strValor = valor.ToString("000000000000000.00");
        string valor_por_extenso = string.Empty;
        for (int i = 0; i <= 15; i += 3)
        {
            valor_por_extenso += Escrever_Valor_Extenso(Convert.ToDecimal(strValor.Substring(i, 3)));
            if (i == 0 & valor_por_extenso != string.Empty)
            {
                if (Convert.ToInt32(strValor.Substring(0, 3)) == 1)
                    valor_por_extenso += " TRILHÃO" + ((Convert.ToDecimal(strValor.Substring(3, 12)) > 0) ? " E " : string.Empty);
                else if (Convert.ToInt32(strValor.Substring(0, 3)) > 1)
                    valor_por_extenso += " TRILHÕES" + ((Convert.ToDecimal(strValor.Substring(3, 12)) > 0) ? " E " : string.Empty);
            }
            else if (i == 3 & valor_por_extenso != string.Empty)
            {
                if (Convert.ToInt32(strValor.Substring(3, 3)) == 1)
                    valor_por_extenso += " BILHÃO" + ((Convert.ToDecimal(strValor.Substring(6, 9)) > 0) ? " E " : string.Empty);
                else if (Convert.ToInt32(strValor.Substring(3, 3)) > 1)
                    valor_por_extenso += " BILHÕES" + ((Convert.ToDecimal(strValor.Substring(6, 9)) > 0) ? " E " : string.Empty);
            }
            else if (i == 6 & valor_por_extenso != string.Empty)
            {
                if (Convert.ToInt32(strValor.Substring(6, 3)) == 1)
                    valor_por_extenso += " MILHÃO" + ((Convert.ToDecimal(strValor.Substring(9, 6)) > 0) ? " E " : string.Empty);
                else if (Convert.ToInt32(strValor.Substring(6, 3)) > 1)
                    valor_por_extenso += " MILHÕES" + ((Convert.ToDecimal(strValor.Substring(9, 6)) > 0) ? " E " : string.Empty);
            }
            else if (i == 9 & valor_por_extenso != string.Empty)
                if (Convert.ToInt32(strValor.Substring(9, 3)) > 0)
                    valor_por_extenso += " MIL" + ((Convert.ToDecimal(strValor.Substring(12, 3)) > 0) ? " E " : string.Empty);
            if (i == 12)
            {
                if (valor_por_extenso.Length > 8)
                    if (valor_por_extenso.Substring(valor_por_extenso.Length - 6, 6) == "BILHÃO" | valor_por_extenso.Substring(valor_por_extenso.Length - 6, 6) == "MILHÃO")
                        valor_por_extenso += " DE";
                    else
                        if (valor_por_extenso.Substring(valor_por_extenso.Length - 7, 7) == "BILHÕES" | valor_por_extenso.Substring(valor_por_extenso.Length - 7, 7) == "MILHÕES"
| valor_por_extenso.Substring(valor_por_extenso.Length - 8, 7) == "TRILHÕES")
                        valor_por_extenso += " DE";
                    else
                            if (valor_por_extenso.Substring(valor_por_extenso.Length - 8, 8) == "TRILHÕES")
                        valor_por_extenso += " DE";
                if (Convert.ToInt64(strValor.Substring(0, 15)) == 1)
                    valor_por_extenso += " REAL";
                else if (Convert.ToInt64(strValor.Substring(0, 15)) > 1)
                    valor_por_extenso += " REAIS";
                if (Convert.ToInt32(strValor.Substring(16, 2)) > 0 && valor_por_extenso != string.Empty)
                    valor_por_extenso += " E ";
            }
            if (i == 15)
                if (Convert.ToInt32(strValor.Substring(16, 2)) == 1)
                    valor_por_extenso += " CENTAVO";
                else if (Convert.ToInt32(strValor.Substring(16, 2)) > 1)
                    valor_por_extenso += " CENTAVOS";
        }
        return valor_por_extenso;
    }
}
static string Escrever_Valor_Extenso(decimal valor)
{
    if (valor <= 0)
        return string.Empty;
    else
    {
        string montagem = string.Empty;
        if (valor > 0 & valor < 1)
        {
            valor *= 100;
        }
        string strValor = valor.ToString("000");
        int a = Convert.ToInt32(strValor.Substring(0, 1));
        int b = Convert.ToInt32(strValor.Substring(1, 1));
        int c = Convert.ToInt32(strValor.Substring(2, 1));
        if (a == 1) montagem += (b + c == 0) ? "CEM" : "CENTO";
        else if (a == 2) montagem += "DUZENTOS";
        else if (a == 3) montagem += "TREZENTOS";
        else if (a == 4) montagem += "QUATROCENTOS";
        else if (a == 5) montagem += "QUINHENTOS";
        else if (a == 6) montagem += "SEISCENTOS";
        else if (a == 7) montagem += "SETECENTOS";
        else if (a == 8) montagem += "OITOCENTOS";
        else if (a == 9) montagem += "NOVECENTOS";
        if (b == 1)
        {
            if (c == 0) montagem += ((a > 0) ? " E " : string.Empty) + "DEZ";
            else if (c == 1) montagem += ((a > 0) ? " E " : string.Empty) + "ONZE";
            else if (c == 2) montagem += ((a > 0) ? " E " : string.Empty) + "DOZE";
            else if (c == 3) montagem += ((a > 0) ? " E " : string.Empty) + "TREZE";
            else if (c == 4) montagem += ((a > 0) ? " E " : string.Empty) + "QUATORZE";
            else if (c == 5) montagem += ((a > 0) ? " E " : string.Empty) + "QUINZE";
            else if (c == 6) montagem += ((a > 0) ? " E " : string.Empty) + "DEZESSEIS";
            else if (c == 7) montagem += ((a > 0) ? " E " : string.Empty) + "DEZESSETE";
            else if (c == 8) montagem += ((a > 0) ? " E " : string.Empty) + "DEZOITO";
            else if (c == 9) montagem += ((a > 0) ? " E " : string.Empty) + "DEZENOVE";
        }
        else if (b == 2) montagem += ((a > 0) ? " E " : string.Empty) + "VINTE";
        else if (b == 3) montagem += ((a > 0) ? " E " : string.Empty) + "TRINTA";
        else if (b == 4) montagem += ((a > 0) ? " E " : string.Empty) + "QUARENTA";
        else if (b == 5) montagem += ((a > 0) ? " E " : string.Empty) + "CINQUENTA";
        else if (b == 6) montagem += ((a > 0) ? " E " : string.Empty) + "SESSENTA";
        else if (b == 7) montagem += ((a > 0) ? " E " : string.Empty) + "SETENTA";
        else if (b == 8) montagem += ((a > 0) ? " E " : string.Empty) + "OITENTA";
        else if (b == 9) montagem += ((a > 0) ? " E " : string.Empty) + "NOVENTA";
        if (strValor.Substring(1, 1) != "1" & c != 0 & montagem != string.Empty) montagem += " E ";
        if (strValor.Substring(1, 1) != "1")
            if (c == 1) montagem += "UM";
            else if (c == 2) montagem += "DOIS";
            else if (c == 3) montagem += "TRÊS";
            else if (c == 4) montagem += "QUATRO";
            else if (c == 5) montagem += "CINCO";
            else if (c == 6) montagem += "SEIS";
            else if (c == 7) montagem += "SETE";
            else if (c == 8) montagem += "OITO";
            else if (c == 9) montagem += "NOVE";
        return montagem;
    }
}
#endregion