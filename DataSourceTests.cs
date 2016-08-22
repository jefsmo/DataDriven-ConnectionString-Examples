
using System;
using System.Configuration;
using System.Linq;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DataDrivenTests
{
    /// <summary>
    /// Represents tests of different ways of connecting to data for data driven tests.
    /// Includes examples of ODBC (Excel Driver, DSN), OLEDB, and VS TestTools XML and CSV providers.
    /// The tests consume data from XML, Excel, and CSV files used as data sources.
    /// </summary>
    /// <remarks>?
    /// The Excel ODBC driver is installed by downloading 'Microsoft Access Database Engine 2010 Redistributable'.
    /// It is stored locally on C:\Program Files (x86)\MSECache\AceRedist\1033\AceRedist.msi
    /// 
    /// The XML and CSV drivers are built into VS: Microsoft.VisualStudio.TestTools.DataSource.[XML / CVS]
    /// You can provide a relative path to the 'DBQ=' key file.
    /// You can provide the complete path to the 'DBQ=' key file.
    /// You can use |DataDirecotry| to refer to the default deployment directory path in the 'DBQ=' key file.
    /// Example: DBQ=|DataDirectory|TestData\\data.xlsx
    /// 
    /// The default |DataDirectory| is bin\debug folder or the deployment folder when deploying tests.
    /// May be able to set this via AppDomain class set method.
    /// 
    /// Add a project reference to System.Data to access data in the TestContext, i.e. TestContext.DataRow["foo"].
    /// </remarks>
    [TestClass]
    public class DataSourceTests
    {
        public TestContext TestContext { get; set; }

        [TestMethod]
        //[DeploymentItem("XMLFile1.xml")]
        [Description("Test sets XML provider in attribute.")]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.XML", "TestData\\XMLFile1.xml", "row", DataAccessMethod.Sequential)]
        public void Args_4_ShouldLoadXML_WhenProviderIsXML()
        {
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x,y,expected);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        //[DeploymentItem("CSVFile1.csv")]
        [Description("Test sets CSV provider in attribute.")]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "TestData\\CSVFile1.csv", "CSVFile1#csv", DataAccessMethod.Sequential)]
        public void Args_4_ShouldLoadCSV_WhenProviderIsCSV()
        {
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x, y, expected);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        //[DeploymentItem("data.xlsx")]
        [Description("Sets Excel ODBC Driver={Microsoft Excel Driver} provider in attribute.")]
        [DataSource("System.Data.Odbc", "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=TestData\\data.xlsx;", "Sheet1$", DataAccessMethod.Sequential)]
        public void Args_4_ShouldLoadExcel_WhenProviderIsOdbcDriver()
        {
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x, y, expected);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        //[DeploymentItem("data.xlsx")]
        [Description("Sets Excel ODBC DSN=Excel Files provider in attribute.")]
        [DataSource("System.Data.Odbc", "Dsn=Excel Files;Dbq=TestData\\data.xlsx", "Sheet1$", DataAccessMethod.Sequential)]
        public void Args_4_ShouldLoadExcel_WhenProviderIsOdbcDsn()
        {
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x, y, expected);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// Attribute contains data provider specific connection string and data table name.
        /// 2 args method seems to only work with OLEDB connection string; not ODBC connection strings.
        /// </summary>
        [TestMethod]
        //[DeploymentItem("data.xlsx")]
        [Description("Sets Excel OLEDB Provider=Microsoft.ACE.OLEDB in attribute.")]
        [DataSource("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=TestData\\data.xlsx; Extended Properties=\"Excel 12.0 Xml; HDR=YES\"", "Sheet1$")]
        public void Args_2_ShouldLoadExcel_WhenProviderIsOleDb()
        {
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x, y, expected);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        //[DeploymentItem("data.xlsx")]
        [Description("Sets Excel OLEDB Provider=Microsoft.ACE.OLEDB in attribute. DataAccessMethod null.")]
        [DataSource("System.Data.Oledb", "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=TestData\\data.xlsx; Extended Properties=\"Excel 12.0 Xml; HDR=YES\"", "Sheet1$", default(DataAccessMethod))]
        public void Args_4_1Null_ShouldLoadExcel_WhenProviderIsOleDb()
        {
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x, y, expected);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        /// DataSource stored in app.config file, not in DataSource attribute.
        /// ODBC provider invariant name = System.Data.Odbc.
        /// ODBC connection string = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=|DataDirectory|\\data.xlsx;"
        /// </summary>
        [TestMethod]
        //[DeploymentItem("data.xlsx")]
        [Description("Sets Excel ODBC Driver= provider in app.config.")]
        [DataSource("MyExcelDriverDataSource")]
        public void Args_1_ShouldLoadExcel_WhenProviderIsAppConfigOdbcDriver()
        {
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var config = (string)AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE");
            Console.WriteLine($"BaseDirectory = {baseDirectory}");
            Console.WriteLine($"App Config File = {config}");

            GetDataSourceInfo();
            GetConnectionStrings();
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x, y, expected);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        //[DeploymentItem("data.xlsx")]
        [Description("Sets Excel ODBC Dsn= provider in app.config.")]
        [DataSource("MyExcelDsnDataSource")]
        public void Args_1_ShouldLoadExcel_WhenProviderIsAppConfigOdbcDsn()
        {
            GetDataSourceInfo();
            GetConnectionStrings();
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x, y, expected);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        //[DeploymentItem("data.xlsx")]
        [Description("Sets Excel OLEDB ACE provider in app.config.")]
        [DataSource("MyExcelOleDbDataSource")]
        public void Args_1_ShouldLoadExcel_WhenProviderIsAppConfigOleDbACE()
        {
            GetDataSourceInfo();
            GetConnectionStrings();
            int x, y, expected;
            GetDataForTest(out x, out y, out expected);

            var actual = AddIntegers(x, y);

            TestContext.WriteLine("{0}, {1}, {2}", x, y, expected);
            Assert.AreEqual(expected, actual);
        }

        private void GetDataForTest(out int x, out int y, out int expected)
        {
            // Use Convert.ToInt() instead of int.Parse() etc.
            // Convert handles the DataRow object without the need to cast to string first.

            var arr = TestContext.DataRow.ItemArray.Select( Convert.ToInt32 ).ToArray();

            x = arr[0];
            y = arr[1];
            expected = arr[2];
        }

        private static int AddIntegers(int first, int second)
        {
            var sum = first;
            for (var i = 0; i < second; i++)
            {
                sum += 1;
            }
            return sum;
        }

        private static void GetDataSourceInfo()
        {
            // TestConfigurationSection: Provides access to data source configuration data. 
            var section = ((TestConfigurationSection)ConfigurationManager.GetSection("microsoft.visualstudio.testtools")).DataSources;
            foreach (DataSourceElement dse in section)
            {
                Console.WriteLine("DataSource Info");
                Console.WriteLine($"Name\t= {dse.Name}");
                Console.WriteLine($"Cnn\t= {dse.ConnectionString}");
                Console.WriteLine($"Table\t= {dse.DataTableName}");
                Console.WriteLine($"DataAcc\t= {dse.DataAccessMethod}");
            }
        }

        /// <summary>
        /// Displays the app.config connection settings section.
        /// </summary>
        private static void GetConnectionStrings()
        {
            var settings = ConfigurationManager.ConnectionStrings;
            if (settings == null) return;

            foreach (ConnectionStringSettings cs in settings)
            {
                Console.WriteLine("Connection string info");
                Console.WriteLine($"Name\t= {cs.Name}");
                Console.WriteLine($"Provider\t= {cs.ProviderName}");
                Console.WriteLine($"Cnn\t= {cs.ConnectionString}");
            }
        }
    }
}
