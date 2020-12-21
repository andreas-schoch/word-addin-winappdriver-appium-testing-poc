using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;

// If a popup error appears during debugging saying: "To prevent an unsafe abort when evaluating the function..." --> https://stackoverflow.com/a/58797847
namespace word_addin_automation_poc_01
{
    // NOTE: These [SomethingSomething] square bracket thingys are so called "Attributes" in C#
    // In this scenario the ones used here are predefined and define the purpose of a method (setup, cleanup, actual test, etc)
    [TestClass]
    public class WordAddinTests
    {
        static WindowsDriver<WindowsElement> session;
        static AppiumLocalService winAppDriverService;
        static WebDriverWait wait;
        private static TestContext objTestContext;

        [ClassInitialize]
        public static void Init(TestContext testContext)
        {
            Debug.WriteLine("Inside Init()");

            // Alternative to running the WinAppDriver.exe file manually but can lead to issues as it sometimes breaks future testruns due to allocated port 4723 already being in use (it doesn't stop on it's own 100% of the time, need to manually kill process)
            // winAppDriverService = new AppiumServiceBuilder().UsingPort(4723).Build();
            // winAppDriverService.Start();

            AppiumOptions options = new AppiumOptions();
            // Find out hidden AUMID of any installed app: get-StartApps | Select-String -Pattern "<Regular app name>"
            options.AddAdditionalCapability("app", @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"); // AUMID doesn't seem to work for word, so actual path is used
            options.AddAdditionalCapability("appArguments", "/q"); // /q starts word without the loading "splash screen"
            //string rootPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "..\\..\\..\\");
            //options.AddAdditionalCapability("appArguments", Path.Combine(rootPath, "my-document.docx")); // open with specific document (path is unreliable outside VS)

            session = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), options);  // replace Uri with winAppDriverService when using local service instead of manually starting WinAppDriver.exe beforehand
            session.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000); // Might improve to overall automation stability & laggy UI errors
            session.Manage().Window.Maximize();

            wait = new WebDriverWait(session, TimeSpan.FromSeconds(10));

            objTestContext = testContext;
        }

        [ClassCleanup]
        public static void TearDown()
        {
            Debug.WriteLine("Inside TearDown()");
            if (session != null)
            {
                //session.Quit(); // close session (Word app in this case) when all tests ran through
            }
        }

        [TestInitialize]
        public void BeforeEach()
        {
            Debug.WriteLine("Inside BeforeEach()");
        }

        [TestCleanup]
        public void AfterEach()
        {
            Debug.WriteLine("Inside AfterEach()");
        }

        [TestMethod]
        public void Test_0001_WordOpens()  // TODO find better way to order test execution (e.g playlist with ms-test) or make sure it is always completely independent
        {
            Debug.WriteLine("Inside Test_0001_WordOpens()");
            // Title is not always a foolproof way to assert things when OS language isn't known in advance. Example: 'Alarm & Clock' vs 'Alarm und Uhr'
            // Assert.AreEqual("my-document.docx - Word", session.Title, $"Expected Title 'my-document.docx - Word' didn't match the actual Title: '{session.Title}'");
            Assert.AreEqual("Word", session.Title, $"Expected Title 'Word' didn't match the actual Title: '{session.Title}'");
        }

        [TestMethod]
        public void Test_0002_AddinLaunchesSuccessfullyInsideWord()
        {
            Debug.WriteLine("Inside Test_0002_AddinLaunchesSuccessfullyInsideWord()");
            // Could fail, as Title is not always a foolproof way to assert things when OS language isn't known in advance. Example: 'Alarm & Clock' vs 'Alarm und Uhr'
            // Assert.AreEqual("my-document.docx - Word", session.Title, $"Expected Title 'Word' didn't match the actual Title: '{session.Title}'");
            Assert.AreEqual("Word", session.Title, $"Expected Title 'Word' didn't match the actual Title: '{session.Title}'");

            clickElementByName(null, "Blank document");
            clickElementByName(null, "Script Lab");
            clickElementByXPath(null, @"//Group[@ClassName=""NetUIChunk""][@Name=""Script""]//Button[@ClassName=""NetUIRibbonButton""][@Name=""Code""]");

            var tab = session.FindElementByAccessibilityId("Pivot21-Tab2");
            wait.Until(pred => tab.Displayed);

            tab.Click();

            //var rect = tab.GetAttribute("BoundingRectangle");

            //"BoundingRectangle"

            //Debug.WriteLine($"-------------------- {tab.Location.X}, {tab.Location.Y} --------- attr ${tab.GetAttribute("ClickablePoint")}");
            Debug.WriteLine($"MYDATA-------------------- {tab.Coordinates.LocationInViewport}  -  {tab.Coordinates.LocationInDom} - {tab.Location} - {tab.GetAttribute("BoundingRectangle")}");
            // Location: 3170, 525 (in legacy webview EdgeHTML only, correct in webview2!)
            // ClickablePoint: 3206, 541
            // ACTUAL on screen: 1708, 283 (from top left corner)


            //session.FindElementByImage()

            var ribbon = session.FindElementByName("Ribbon"); // used as an anchor to know the top left corner
            //Actions action = new Actions(session);
            //action.MoveToElement(tab, -1470, -250);
            //action.MoveToElement(null, 1708, 283);
            //action.MoveToElement(ribbon, tab.Coordinates.LocationOnScreen.X, tab.Coordinates.LocationOnScreen.Y);

            //action.Click();
            //action.MoveByOffset(1708, 283);
            //action.Click(); // it was mentioned that ContextMenu could fail sometimes and that a .Click() can solve the issue somehow
            //action.ContextClick();
            //action.Perform();


            //clickElementByXPath(null, @"//Group[@ClassName=""NetUIChunk""][@Name=""Add-ins""]//SplitButton[@ClassName=""NetUISplitButtonAnchor""][@Name=""My Add-ins""]//MenuItem[@ClassName=""NetUIRibbonButton""][@Name=""More Options""]");
            //// Alternative query in two steps without using XPath syntax
            //// var parentContext = session.FindElementByName("Add-ins");
            //// clickElementByName(parentContext, "More Options");
            //clickElementByName(null, "My Office Add-in");
            //clickElementByName(null, "Show Taskpane");


            //System.Threading.Thread.Sleep(5000);
            ////clickElementByXPath(null, @"//Pane[@ClassName=""Internet Explorer_Server""][@Name=""My Office Add-in""]//Pane[@Name=""My Office Add-in""]//Group[position()=2]//Button[@Name=""Run""]");

            ////session.FindElementByCssSelector("div[role]").Click();
            //var runBtn = session.FindElementByAccessibilityId("my-button-id");

            //ReadOnlyCollection<string> contextNames = session.Contexts;
            //foreach (string context in contextNames) {
            //    Debug.WriteLine($"----Context: {context}");
            //}


            List<string> AllContexts = new List<string>();
            foreach (var context in (session.Contexts))
            {
                AllContexts.Add(context);
            }




            System.Threading.Thread.Sleep(1000);


        }

        private static void clickElementByName(WindowsElement customContext, string elementName)
        {
            // TODO find a way to be able pass parameter that is either WindowsElement or WindowsDriver<WindowsElement> instead of using customContext or session and the if-else-statement
            if (customContext == null)
            {
                var element = session.FindElementByName(elementName);
                // wait.Until(pred => element.Displayed); // could fail if element displayed but not clickable yet.
                wait.Until(ExpectedConditions.ElementToBeClickable(element));
                element.Click();
            }
            else
            {
                var element = customContext.FindElementByName(elementName);
                wait.Until(ExpectedConditions.ElementToBeClickable(element));
                element.Click();
            }
            
        }

        private static void clickElementByXPath(WindowsElement customContext, string elementXPath)

        {
            if (customContext == null)
            {
                var element = session.FindElementByXPath(elementXPath);
                wait.Until(ExpectedConditions.ElementToBeClickable(element));
                element.Click();
            }
            else
            {
                var element = customContext.FindElementByXPath(elementXPath);
                wait.Until(ExpectedConditions.ElementToBeClickable(element));
                element.Click();
            }
        }
    }
}
