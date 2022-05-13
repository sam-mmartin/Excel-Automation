using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Threading;

namespace TestExcelDemo
{
   public static class ElementExtensions
   {
      public static void ClearEditor(this WindowsElement editBox)
      {
         editBox.SendKeys(Keys.Control + "a" + Keys.Control);
         editBox.SendKeys(Keys.Delete);
      }

      public static void Clicar(this WindowsElement element, int milliseconds = 0)
      {
         element.Click();
         AppDriver.Session!.Waiting(milliseconds);
      }

      public static void Waiting(this WindowsDriver<WindowsElement> session, int milliseconds)
      {
         Thread.Sleep(TimeSpan.FromMilliseconds(milliseconds));
      }
   }
}
