using DocumentFormat.OpenXml.Spreadsheet;
using GeradorRecibo.Model;
using OfficeOpenXml;
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

int totalCols = 27;
int cabecalho = 2;
int conteudo = 3;
var path = @$"../../../PlanilhaPagamento-{DateTime.Today.Year}.xlsx";
List<MoradorModel> list = new List<MoradorModel>();

Console.WriteLine("CARREGANDO...");
Pausa();
LerArquivo();

void Pausa()
{
    Console.WriteLine("APERTE QUALQUER TECLA PARA CONTINUAR...");
    Console.ReadKey();
}    

void LerArquivo()
{
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
                                mesModel.Pago = value == null ? 0 : decimal.Parse(value.ToString());
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
                Pausa();
            }

        }
    }
}

void GerarRecibos()
{

}
