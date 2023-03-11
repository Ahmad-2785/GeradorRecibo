using DocumentFormat.OpenXml.Spreadsheet;
using GeradorRecibo.Model;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using OfficeOpenXml;
using System.Text;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

int totalCols = 27;
int cabecalho = 3;
int conteudo = 4;

int sequencial = 1;
string endereco = "";

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
    GerarRecibos();
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
                    moradorModel.MesesPagos = new List<MesesModel>();

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
                        else
                        {
                            if (i % 2 == 0)
                            {
                                var mes = worksheet.Cells[cabecalho, i].Value;
                                mesModel.Mes = mes.ToString();
                                mesModel.Pago = value == null ? null : value.ToString();
                            }
                            else
                            {
                                mesModel.Gerado = value == null ? false : true;
                                moradorModel.MesesPagos.Add(mesModel);
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

        foreach (var mes in morador.MesesPagos.Where(x => x.Pago != null && !x.Gerado))
        {
            var caminhoReferencia = $@"{caminhoGravacaoBase}\Recibo-{mes.Mes}.pdf";

            Console.WriteLine(mes.Mes.ToString());
            Console.WriteLine(mes.Gerado.ToString());

            GeradorPDF(caminhoReferencia);
        }
    }
}

void GeradorPDF(string caminho)
{
    #region CRIACAO DO ARQUIVO
    Document doc = new Document(PageSize.A4.Rotate());
    doc.SetMargins(20, 20, 20, 20);
    FileStream fs = new FileStream(caminho, FileMode.Create, FileAccess.Write);
    PdfWriter writer = PdfWriter.GetInstance(doc, fs);
    doc.Open();
    #endregion

    #region CONTEUDO
    var css = "";
    StringBuilder sb = new StringBuilder();
    sb.AppendLine("<div>CONTEUDO AQUI</div>");
    #endregion

    #region FINALIZACAO DO ARQUIVO
    XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString())), new MemoryStream(Encoding.UTF8.GetBytes(css.ToString())));
    doc.Close();
    #endregion
}
