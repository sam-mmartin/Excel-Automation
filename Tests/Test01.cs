using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestExcelDemo.Tests
{
   [TestClass]
   public class Test01 : DesktopAppSession
   {
      [TestMethod]
      public void ContentHeaderTest()
      {
         appWindow = AppDriver.Session!.FindElementByAccessibilityId("A1");
         Assert.AreEqual("cpf/cnpj", appWindow.Text);
      }

      [ClassInitialize]
      public static void ClassInitialize(TestContext context)
      {
         Setup(context);
      }

      [ClassCleanup]
      public static void ClassCleanup()
      {
         TearDown();
      }
   }
}
