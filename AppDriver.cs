using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;

namespace TestExcelDemo
{
   internal class AppDriver
   {
      protected const string WinAppDriverURL = "http://127.0.0.1:4723";
      protected static WindowsDriver<WindowsElement>? session;
      protected static WindowsDriver<WindowsElement>? winSession;
      public static WindowsDriver<WindowsElement>? Session
      {
         get => session;
         set => session = value;
      }
      public static WindowsDriver<WindowsElement>? WinSession
      {
         get => winSession;
         set => winSession = value;
      }

      public static void InitSystem(TestContext context)
      {
         if (context is null)
         {
            throw new ArgumentNullException(nameof(context));
         }

         if (winSession is null)
         {
            // Create session at root level
            var rootCapabilities = new AppiumOptions();
            rootCapabilities.AddAdditionalCapability("app", "Root");
            winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);

            Assert.IsNotNull(winSession);
            Assert.IsNotNull(winSession.SessionId);

            winSession.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
         }
      }

      public static void InitApplication(TestContext context, string capability, string AppId)
      {
         if (context is null)
         {
            throw new ArgumentNullException(nameof(context));
         }

         // Create a new session to launch Notepad application
         var appCapabilities = new AppiumOptions();
         appCapabilities.AddAdditionalCapability(capability, AppId);
         session = new WindowsDriver<WindowsElement>(new Uri(WinAppDriverURL), appCapabilities);

         Assert.IsNotNull(session);
         Assert.IsNotNull(session.SessionId);

         // Set implicit timeout to 1.5 seconds to make element search to retry every 500ms for at most three times
         session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
      }

      public static void ActiveApplication(TestContext context, string name)
      {
         InitSystem(context);

         if (winSession != null)
         {
            var element = winSession.FindElementByName(name);
            var appTopLevelWindowHandle = element.GetAttribute("NativeWindowHandle");
            appTopLevelWindowHandle = (int.Parse(appTopLevelWindowHandle)).ToString("x"); // Convert to Hex

            InitApplication(context, "appTopLevelWindow", appTopLevelWindowHandle);
         }
      }

      public static void QuitApplication(string btnSave)
      {
         if (Session != null)
         {
            Session.Close();

            try
            {
               if (!string.IsNullOrEmpty(btnSave))
               {
                  // Dismiss save dialog if it is blocking the exit
                  Session.FindElementByName(btnSave).Click();
               }
            }
            catch { }

            Session.Quit();
            Session = null;
         }

         if (WinSession != null)
         {
            WinSession.Close();
            WinSession.Quit();
            WinSession = null;
         }
      }
   }
}
