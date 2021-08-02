using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOIHelper;
using System.IO;
using System.Linq;

namespace NPOIHelperTest
{
    [TestClass]
    public class UnitTestExcelReader : BaseUnitTest
    {
        [TestMethod]
        public void TestReadList()
        {
            var users = NPOIHelperBuild.ReadExcel<ImportUser>("TestImportExcel/test1.xlsx");

            WriteLine(users);
            Assert.IsNotNull(users);
            Assert.IsTrue(users.Any());
            CollectionAssert.AllItemsAreNotNull(users.ToList());
        }

        [TestMethod]
        public void TestReadListStream()
        {
            var stream = File.OpenRead("TestImportExcel/test1.xlsx");
            // var users = NPOIHelperBuild.ReadExcel<ImportUser>(stream, "application/vnd.ms-excel");
            var users = NPOIHelperBuild.ReadExcel<ImportUser>(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            WriteLine(users);
            Assert.IsNotNull(users);
            Assert.IsTrue(users.Any());
            CollectionAssert.AllItemsAreNotNull(users.ToList());
        }

        [TestMethod]
        public void TestReadDataTable()
        {
            var dt = NPOIHelperBuild.ReadExcel("TestImportExcel/test1.xlsx");
            WriteLine(dt);
            Assert.IsNotNull(dt);

        }

        [TestMethod]
        public void TestReadDataTableStream()
        {
            var stream = File.OpenRead("TestImportExcel/test1.xlsx");
            var dt = NPOIHelperBuild.ReadExcel(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            WriteLine(dt);
            Assert.IsNotNull(dt);
        }
    }
}
