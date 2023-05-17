using System.Text;
using OfficeOpenXml;

namespace ExcelCSharp;

public class ExcelReader
{
    private readonly String _path;
    

    public ExcelReader(String path){
        _path = path;
    }

    private string toCell(uint X, uint Y){
        StringBuilder ans = new StringBuilder();
        while(X > 26){
            ans.Append("A");
            X -= 26;
        }
        ans.Append(Convert.ToChar(65 + X));
        ans.Append(Y);
        return ans.ToString();
    }

    public List<Model> Read(int worksheet, uint startX, uint startY){
        List<Model> ans = new List<Model>();

        using(var package = new ExcelPackage(new FileInfo(_path)))
        {
            var sheet = package.Workbook.Worksheets[worksheet];
            bool isValid = true;
            do{
                string 
                    fieldA = sheet.Cells[toCell(startX + 0, startY + (uint)ans.Count)].Text,
                    fieldB = sheet.Cells[toCell(startX + 1, startY + (uint)ans.Count)].Text,
                    fieldC = sheet.Cells[toCell(startX + 2, startY + (uint)ans.Count)].Text,
                    fieldD = sheet.Cells[toCell(startX + 3, startY + (uint)ans.Count)].Text;
                try{
                    Int64 id = Int64.Parse(fieldA);
                    Double price = Double.Parse(fieldC);
                    DateOnly date = DateOnly.Parse(fieldD);

                    ans.Add(new Model(){
                        Id = id,
                        Name = fieldB,
                        Price = price,
                        Date = date
                    });
                }catch(Exception e){
                    e.GetType();
                    isValid = false;
                }
            }while(isValid);
        }

        return ans;
    }
}
