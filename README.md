# NUnit-DataDriven-tests-from-Excel-files

The purpose of Unit Testing is to validate that each unit of the software works as expected, so we're gonna through Nunit which is the most popular unit test framework for .NET and know how to read data from excel file and use this data through Nunit attribute(TestCaseSource). Let's start ;)

  # 1) At first the function `ReadFromExcel`:ي
   - it takes the `excelFileName` and `excelsheetTabName` and return a list of TestCaseData attribute. 
     - Here it gets the path to the excel file in your project. if the file not found it throws an exception. 
 ```c#
     string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
     string xslLocation = Path.Combine(executableLocation, "data/"+ excelFileName);    
     string cmdText = "SELECT * FROM [" + excelsheetTabName + "$]";
     if (!File.Exists(xslLocation))
        throw new Exception(string.Format("File name: {0}", xslLocation), new FileNotFoundException());
```                
  # 2) 
  - After getting the path we need to be able to read the file and get the data from it, so we have to open a connection through `OleDb`    `which is an API designed by Microsoft, allows accessing data from a variety of sources in a uniform manner.` Now we opened the           connection and will start reading the file row by row then add the row in our list.
  - See the code below: 
  

```c#
 public class ExcelReader
        {
            public static IEnumerable<TestCaseData> ReadFromExcel(string excelFileName, string excelsheetTabName)
            {     
                string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string xslLocation = Path.Combine(executableLocation, "data/"+ excelFileName);
        
                string cmdText = "SELECT * FROM [" + excelsheetTabName + "$]";
                if (!File.Exists(xslLocation))
                    throw new Exception(string.Format("File name: {0}", xslLocation), new FileNotFoundException());
                string connectionStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES\";", xslLocation);
                var testCases = new List<TestCaseData>();
                using (var connection = new OleDbConnection(connectionStr))
                {
                    connection.Open();
                    var command = new OleDbCommand(cmdText, connection);
                    var reader = command.ExecuteReader();
                    if (reader == null)
                        throw new Exception(string.Format("No data return from file, file name:{0}", xslLocation));
                    while (reader.Read())
                    {
                        var row = new List<string>();
                        var feildCnt = reader.FieldCount;
                        for (var i = 0; i < feildCnt; i++)
                            row.Add(reader.GetValue(i).ToString());
                        testCases.Add(new TestCaseData(row.ToArray()));
                    }
                }

                if (testCases != null)
                    foreach (TestCaseData testCaseData in testCases)
                        yield return testCaseData;            
            }
        }    
```
# Passing the data to the `TestCaseSource` attribute
  - In your test class, create a new function then send the `FILENAME` and `TabName`.
```c# 
     private static string FILENAME = "Registeration_Data.xlsx";

     public static IEnumerable<TestCaseData> RegisrtrationData()
     {
        return ExcelReader.ReadFromExcel(FILENAME, "Registeration");
     }
```
 # Here is an example:
  - In this example we validate the testcases of the `name` field in the registeration form. 
 ```c#
 [TestCaseSource("RegisrtrationData")]
        public void CheckName_Validations(string name, string email, string phone, string password,string confirmpassword)
        {
            Assert.Multiple(() =>
            {
                Assert.LessOrEqual(name.Length, 20);
                Assert.IsNotEmpty(name);
                Assert.IsFalse(name.Length > 20);
                for (int i = 0; i < Special_Chars.Length; i++)
                {
                    Assert.That(name, Does.Not.Contains(Special_Chars[i]));
                }
                Assert.IsFalse(name.Any(char.IsDigit));
                Assert.IsNotEmpty(email);
                Assert.IsNotEmpty(phone);
                Assert.IsNotEmpty(password);
                Assert.IsNotEmpty(confirmpassword);

            });
            
        }
 
 ```
 
  
