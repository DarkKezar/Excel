namespace ExcelCSharp;

public class Model
{
    public Int64 Id { get; set; }
    public String? Name { get; set; }
    public Double Price { get; set; }
    public DateOnly Date { get; set; }

    public override string ToString(){
        return Id + " " + Name + " " + Price + " " + Date.ToString();
    }
}
