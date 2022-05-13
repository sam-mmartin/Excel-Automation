using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace TestExcelDemo
{
   public class DesktopAppSession
   {
      private const string appID = @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE";
      private protected static WindowsElement? appWindow;

      public static void Setup(TestContext context)
      {
         AppDriver.InitApplication(context, "app", appID);

         AppDriver.ActiveApplication(context, "Excel");

         Assert.IsNotNull("Excel", AppDriver.Session!.Title);
      }

      public static void TearDown() => AppDriver.QuitApplication("");

      [TestInitialize]
      public void TestInitialize()
      {
         appWindow = AppDriver.Session!.FindElementByName("Página Inicial");
         appWindow.Clicar();

         appWindow = AppDriver.Session.FindElementByName("CadastroEmail.xlsx");
         appWindow.Clicar();

         appWindow = AppDriver.Session.FindElementByName("Grade");

         Assert.IsNotNull(appWindow);
      }

      protected static string SanitizeBackslashes(string input) => input.Replace("\\", Keys.Alt
                                                                                       + Keys.NumberPad9
                                                                                       + Keys.NumberPad2
                                                                                       + Keys.Alt);
   }
}