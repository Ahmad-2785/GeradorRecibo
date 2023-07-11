using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GeradorRecibo.Model;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//TOTAL DE COLUNAS DO EXCEL
int totalCols = 40;
//LINHA DO CABECALHO
int cabecalho = 3;
//LINHA DO INICIO DO CONTEUDO DA TABELA
int conteudo = 4;

//O RECIBO POSSUI UMA SEQUENCIA A SER SEGUIDA
int sequencial = 1;
//VARIAVEL PARA GUARDAR O ENDERECO QUE SERA COLOCADO NO RECIBO
string endereco = "";

//DATA ATUAL PARA SER COLOCADA NO RECIBO
var dataAtual = DateTime.Today;
//CONFIGURACAO DA LINGUA USADA NO PROJETO
var cultureInfo = new CultureInfo("pt-BR");
//DATA POR EXTENSO PARA ASSINATURA DO RECIBO
var dataExtenso = $@"{dataAtual.Day} de {dataAtual.ToString("MMMM", cultureInfo)} de {dataAtual.Year}";
//DATA ABREVIADA PARA SER COLOCADA NO CANHOTO DO RECIBO
var dataAbreviada = $@"{dataAtual.ToString("MM/yyyy")}";

//CAMINHO DE ONDE QUER QUE OS ARQUIVOS SEJAM GRAVADOS
var caminhoGravacao = $@"C:\Users\sciam\Desktop";

//CAMINHO DO TEMPLATE DO EXCEL USADO
var pagamentoPath = $"{AppDomain.CurrentDomain.BaseDirectory}/Templates/PlanilhaPagamento-{DateTime.Today.Year}.xlsx";
//CAMINHO DO TEMPLATE DO DOCX DO RECIBO
var reciboPath = $"{AppDomain.CurrentDomain.BaseDirectory}/Templates/RECIBO-BASE.docx";

//INICIALIZACAO DA LISTA QUE VAI GUARDAR OS ITENS OBTIDOS DO EXCEL
List<MoradorModel> list = new List<MoradorModel>();

//DE MOMENTO DESABILITADO PARA SER MAIS RAPIDO A EXECUCAO
void Pausa()
{
    Console.WriteLine();
    Console.WriteLine("APERTE QUALQUER TECLA PARA CONTINUAR...");
    //Console.ReadKey();
}

Main();

void Main()
{
    Console.WriteLine();
    Console.WriteLine("CARREGANDO...");
    LerArquivo();
    GerarDadosRecibos();
}

void LerArquivo()
{
    Console.WriteLine();
    Console.WriteLine("INICIANDO LEITURA DO ARQUIVO...");
    Pausa();

    using (MemoryStream ms = new MemoryStream())
    {
        using (FileStream file = new FileStream(pagamentoPath, FileMode.Open, FileAccess.Read))
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
                    moradorModel.Meses = new List<MesesModel>();

                    //COLUNAS FIXAS
                    moradorModel.Id = Int32.Parse(worksheet.Cells[row,1].Value.ToString());
                    moradorModel.Casa = worksheet.Cells[row, 2].Value != null ? worksheet.Cells[row, 2].Value.ToString() : "";
                    moradorModel.Morador = worksheet.Cells[row, 3].Value != null ? worksheet.Cells[row, 3].Value.ToString() : "";
                    moradorModel.Email = worksheet.Cells[row, 4].Value != null ? worksheet.Cells[row, 4].Value.ToString() : "";

                    //COLUNAS QUE SE REPETEM
                    for (int i = 5; i <= totalCols; i += 3)
                    {
                        MesesModel mesModel = new MesesModel();

                        //NOME DO MES
                        mesModel.Mes = worksheet.Cells[3, i].Value.ToString();
                        //VALOR PAGO
                        mesModel.Valor = worksheet.Cells[row, i].Value != null ? worksheet.Cells[row, i].Value.ToString() : "";
                        //NOME DO COBRADOR
                        mesModel.Cobrador = worksheet.Cells[row, i + 1].Value != null ? worksheet.Cells[row, i + 1].Value.ToString() : "";
                        //RECIBO GERADO
                        mesModel.Gerado = worksheet.Cells[row, i + 2].Value.ToString() == "Sim" ? true : false;

                        moradorModel.Meses.Add(mesModel);
                    }

                    if(moradorModel.Id > 0)
                        list.Add(moradorModel);
                }

                Console.WriteLine();
                Console.WriteLine("ARQUIVO CARREGADO!");
            }
        }
        ms.Close();
    }
}

void GerarDadosRecibos()
{
    Console.WriteLine();
    Console.WriteLine("INICIANDO GERACAO DOS RECIBOS...");
    Pausa();

    foreach(var morador in list.OrderBy(x => x.Id))
    {
        //CRIACAO DA PASTA DA CASA E ANO REFERENTE
        var caminhoGravacaoBase = $@"{caminhoGravacao}\MatupaRecibos\{morador.Casa}\{DateTime.Today.Year}";
        
        if(!Directory.Exists(caminhoGravacaoBase))
            Directory.CreateDirectory(caminhoGravacaoBase);

        Console.WriteLine();
        Console.WriteLine($@"MORADOR: {morador.Morador}");
        Console.WriteLine($@"CASA: {morador.Casa}");

        var mesesPagos = morador.Meses.Where(x => !string.IsNullOrEmpty(x.Valor) && !x.Gerado);
        var qtdMesesPagos = mesesPagos.Count();
        var enderecoGravar = endereco.Replace("X", morador.Casa);
        var texto = "";

        if (qtdMesesPagos > 1)
        {
            var valor = decimal.Parse(mesesPagos.LastOrDefault().Valor == "X" ? mesesPagos.FirstOrDefault().Valor : mesesPagos.LastOrDefault().Valor);
            if (qtdMesesPagos > 2)
            {
                texto = $@"PGTO da cota {mesesPagos.FirstOrDefault().Mes} até {mesesPagos.LastOrDefault().Mes} / {DateTime.Today.Year}";
            }
            else
            {
                texto = $@"PGTO da cota {mesesPagos.FirstOrDefault().Mes} e {mesesPagos.LastOrDefault().Mes} / {DateTime.Today.Year}";
            }

            var nomeArquivo = $@"Recibo-{mesesPagos.FirstOrDefault().Mes}-{mesesPagos.LastOrDefault().Mes}.docx";

            GerarRecibos(caminhoGravacaoBase, nomeArquivo, sequencial, morador.Morador, morador.Casa, valor, texto, mesesPagos.LastOrDefault().Cobrador);
            GerarPDF(caminhoGravacaoBase, nomeArquivo);

            //INCREMENTAR O SEQUENCIAL
            sequencial++;
        }
        else if (qtdMesesPagos == 1)
        {
            var nomeArquivo = $@"Recibo-{mesesPagos.FirstOrDefault().Mes}.docx";

            texto = $@"manutenção de {mesesPagos.FirstOrDefault().Mes}";

            var valor = decimal.Parse(mesesPagos.FirstOrDefault().Valor);

            GerarRecibos(caminhoGravacaoBase, nomeArquivo, sequencial, morador.Morador, morador.Casa, valor, texto, mesesPagos.FirstOrDefault().Cobrador);
            GerarPDF(caminhoGravacaoBase, nomeArquivo);
            
            //INCREMENTAR O SEQUENCIAL
            sequencial++;
        }        
    }
}

//UTILIZA UMA BASE DE LAYOUT EM PDF E SUBSTITUI AS PALAVRAS PARA GERAR OS RECIBOS
void GerarRecibos(string caminhoGravacao, string nomeArquivo, int sequencia, string morador, string casa, decimal valor, string observacao, string cobrador)
{
    var path = $@"{caminhoGravacao}\{nomeArquivo}";

    File.Copy(reciboPath, path, true);

    using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
    {
        var body = doc.MainDocumentPart.Document.Body;
        foreach (var text in body.Descendants<Text>())
        { 
            text.Text = text.Text.Replace("{morador}", morador);
            text.Text = text.Text.Replace("{seq}", $"{sequencia}/{DateTime.Today.Year}");
            text.Text = text.Text.Replace("{casa}", casa);
            text.Text = text.Text.Replace("{valor}", valor.ToString("F"));
            text.Text = text.Text.Replace("{valorExtenso}", EscreverExtenso(valor));
            text.Text = text.Text.Replace("{observacao}", observacao);
            text.Text = text.Text.Replace("{cobrador}", cobrador);
            text.Text = text.Text.Replace("{data}", dataExtenso);
            text.Text = text.Text.Replace("{dataAbrev}", dataAbreviada);
        }
        doc.Save();
    }

    Console.WriteLine("RECIBO GERADO !!!");
}

//USADO PARA CONVERTER O DOCX PARA PDF COM O LIBREOFFICE COMO SERVICO DO WINDOWS
void GerarPDF(string caminhoGravacaoBase, string arquivoDocx)
{
    var path = $@"{caminhoGravacaoBase}\{arquivoDocx}";    
    var outFile = $@"{caminhoGravacaoBase}";

    try
    {
        Process process = new Process();
        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
        startInfo.FileName = @"C:\windows\system32\cmd.exe";
        startInfo.Arguments = "/c \"C:\\Program Files\\LibreOffice\\program\\soffice.exe\" --headless --convert-to pdf --outdir " + outFile + " " + path;
        process.StartInfo = startInfo;
        process.Start();
        process.WaitForExit();

        Console.WriteLine("PDF GERADO !!!");
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.ToString());
        Pausa();
    }
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