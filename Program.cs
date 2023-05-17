
using ExcelCSharp;
using OfficeOpenXml;

string path = "/Users/darkkezar/Documents/GitHub/ExcelCSharp/testBook.xlsx";


ExcelReader reader = new ExcelReader(path);

List<Model> ans = reader.Read(0, 1, 4);

foreach(var i in ans)
    Console.WriteLine(i.ToString());
