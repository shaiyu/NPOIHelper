using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOIHelper;
using NPOIHelperTest.Model;
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
        public void TestReadListRecommend()
        {
            using var stream = File.OpenRead("TestImportExcel/test2.xlsx");
            using var reader = NPOIHelperBuild.GetReader(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            var users = reader.ReadSheet<ImportRecommend>(0);
            var users2 = reader.ReadSheet<ImportRecommend>(1);

            WriteLine(users);
            WriteLine(users2);
            Assert.IsNotNull(users);
            Assert.IsTrue(users.Any());
            CollectionAssert.AllItemsAreNotNull(users.ToList());
        }
        
        [TestMethod]
        [DataRow(0)]
        [DataRow(1)]
        public void TestReadListSheet(int sheetIndex)
        {
            var users = NPOIHelperBuild.ReadExcel<ImportUser2>("TestImportExcel/test1.xlsx", sheetIndex);

            WriteLine(users);
            Assert.IsNotNull(users);
            Assert.IsTrue(users.Any());
            CollectionAssert.AllItemsAreNotNull(users.ToList());
        }


        [TestMethod]
        public void TestReadListStream()
        {
            using var stream = File.OpenRead("TestImportExcel/test1.xlsx");
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
            using var stream = File.OpenRead("TestImportExcel/test1.xlsx");
            var dt = NPOIHelperBuild.ReadExcel(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            WriteLine(dt);
            Assert.IsNotNull(dt);
        }
    }
}
