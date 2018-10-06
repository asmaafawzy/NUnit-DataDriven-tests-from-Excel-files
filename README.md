# NUnit-DataDriven-tests-from-Excel-files

 The purpose of Unit Testing is to validate that each unit of the software works as expected.
 So we're going through NUnit which is the most popular unit test framework for .NET. 
 Next, we will know how to read data from excel file and use this data through NUnit attribute `TestCaseSource`. 

Let's start ;) 

- See the code below and I will explain it. 
  
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

  1) At first the function `ReadFromExcel`:
  
     - It takes the `excelFileName`, `excelsheetTabName` and returns a list of `TestCaseData` attribute. 
     - Here it gets the path to the excel file in your project and if the file not found it throws an exception. 
     
 
 ```c#
     string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
     string xslLocation = Path.Combine(executableLocation, "data/"+ excelFileName);    
     string cmdText = "SELECT * FROM [" + excelsheetTabName + "$]";
     if (!File.Exists(xslLocation))
        throw new Exception(string.Format("File name: {0}", xslLocation), new FileNotFoundException());
```   

  2) After getting the path, we need to be able to read the file and get the data from it.
      We have to open a connection through `OleDb: which is an API designed by Microsoft, allows accessing data from a variety of             sources in a uniform manner`. 
      Now we open the connection and will start reading the file row by row then add each row to our list.
  
```c#
     string connectionStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0  Xml;HDR=YES\";", xslLocation);
                
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
  
```

  3) Passing data to the `TestCaseSource` attribute
  
   - In your test class, you should create a new function then pass the values of `FILENAME` and `TabName` of your data file.
  
```c# 
     private static string FILENAME = "Registeration_Data.xlsx";

     public static IEnumerable<TestCaseData> RegisrtrationData()
     {
        return ExcelReader.ReadFromExcel(FILENAME, "Registeration");
     }
```

 ### Here is an example:
 
  - In this example, we validate the test cases of the `name` field in the registration form. 
  
 ```c#
 [TestCaseSource("RegisrtrationData")]
        public void CheckName_Validations(string name, string email, string phone, string password)
        {
            Assert.Multiple(() =>
            {
                Assert.IsNotEmpty(name);
                ......
            });      
        }
 
 ```
 
  
