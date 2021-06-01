using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOIHelper;
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
        public void TestReadDataTable()
        {
            var dt = NPOIHelperBuild.ReadExcel("TestImportExcel/test1.xlsx");
            WriteLine(dt);
            Assert.IsNotNull(dt);
        }
    }
}
