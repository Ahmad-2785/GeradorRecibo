namespace GeradorRecibo.Model;

public class MoradorModel
{
    public int Id { get; set; }
    public string Casa { get; set; }
    public string Morador { get; set; }

    public ICollection<MesesModel> MesesPagos { get; set; }    
}

public class MesesModel
{
    public string Mes { get; set; }
    public string? Pago { get; set; }
    public bool Gerado { get; set; }
}
