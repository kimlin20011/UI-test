using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Remote;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Linq;
using System.Diagnostics;
using System.Threading;

namespace SkypeAutoCapture
{
    class SkypeSession
    {
        // Note: append /wd/hub to the URL if you're directing the test at Appium
        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //決定app id windows app appId
        private const string CalculatorAppId = "Microsoft.SkypeApp_kzf8qxf38zg5c!App";

        protected WindowsDriver<WindowsElement> session;

        public void Setup()
        {
            // Launch Calculator application if it is not yet launched
            if (session == null)
            {
                // Create a new session to bring up an instance of the Calculator application
                // Note: Multiple calculator windows (instances) share the same process Id
                Console.WriteLine("Connecting App ...");
                Process[] wad = Process.GetProcessesByName("WinAppDriver");
                if (wad.Length < 1)
                {
                    //啓動WinAppDriver.exe
                    Process.Start(@".\Windows Application Driver\WinAppDriver.exe");
                }

                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", CalculatorAppId);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                // Assert.IsNotNull(session);

                // Set implicit timeout to 1.5 seconds to make element search to retry every 500 ms for at most three times
                session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
                session.Manage().Window.Maximize(); //放到最大
            }
        }
        //截取一張相片
        public bool takeScreenShotOnce()
        {
            //("ScrollViewer")[1]代表Skype右邊的方框欄位；SaveAsFile將截圖放到某一個路徑
            session.FindElementsByClassName("ScrollViewer")[1].GetScreenshot().SaveAsFile("ScreenshotAlarmPivotItem.png", ScreenshotImageFormat.Png);

            return true;
        }


        //滑動skype右方欄位，並截取所有聊天室的對話圖片
        public bool autoScrollScreenShotOnePage(string folderPath, int side, int num)
        //folderPath為儲存路徑；side為滑動方向 0為up 1為down；num代表要截圖的數量，可以先設定MAX再讓系統判斷
        {
            WindowsElement ContentScrollViewer = null;
            XElement ContentScrollBar = null;


            //如果要截圖的數量大於0
            while (num-- > 0)
            {
                try
                {
                    //先找ScrollViewer[1]對話框 = 右邊的聊天室 ScrollViewer是skype裏面設定的名稱
                    var ScrollViewerElements = session.FindElementsByClassName("ScrollViewer");
                    //如沒找到對話框，可能是使用者把界面關掉了
                    if (ScrollViewerElements.Count == 0)
                    {
                        throw new ArgumentException("Can't Find View");
                    }

                    //將設定聊天室元件變數
                    ContentScrollViewer = ScrollViewerElements[1];

                    //PageSource為該頁面的内容: xml
                    string source = session.PageSource;

                    //讀取該頁面的XML  
                    XDocument doc = XDocument.Parse(session.PageSource);

                    //將ScrollViewer元件的XML讀取出來
                    var xmlScrollViewerElements = doc.Descendants()
                              .Where(x =>
                              x.Attribute("ClassName").Value == "ScrollViewer");

                    //將滑動杆的XML讀取出來
                    var xmlContentScrollBarElements = doc.Descendants()
                              .Where(x =>
                              x.Attribute("ClassName").Value == "ScrollBar");

                    //先判斷xmlScrollViewerElements是、 是可以垂直滑動的
                    if (xmlScrollViewerElements.ElementAt(1).Attribute("VerticallyScrollable").Value == "True")
                    {
                        if (xmlScrollViewerElements.ElementAt(0).Attribute("VerticallyScrollable").Value == "False")
                        {
                            ContentScrollBar = xmlContentScrollBarElements.ElementAt(0);
                        }
                        else
                        {
                            ContentScrollBar = xmlContentScrollBarElements.ElementAt(1);
                        }
                    }

                    //將原件對應到聊天框的element
                    var matchingElements = doc.Descendants()
                              .Where(x =>
                              x.Attribute("Name").Value != "" &&
                              x.Attribute("y").Value != "0" &&
                              x.Attribute("ClassName").Value != "TextBlock" &&
                              x.Attribute("ClassName").Value != "RichTextBlock" &&
                              x.Attribute("LocalizedControlType").Value == "group");

                    //設定filename
                    string filename = DateTime.Now.Month.ToString() + "-" +
                        DateTime.Now.Day.ToString() + "-" +
                        DateTime.Now.Hour.ToString() + "-" +
                        DateTime.Now.Minute.ToString() + "-" +
                        DateTime.Now.Second.ToString() + "-" +
                        DateTime.Now.Millisecond.ToString();

                    //產生截圖並儲存
                    Console.WriteLine("Getting Picture ... " + filename);
                    //使用WinDriver將畫面截圖保存
                    session.GetScreenshot().SaveAsFile(folderPath + "\\" + filename + ".png", ScreenshotImageFormat.Png);

                    //產生文字檔並儲存
                    Console.WriteLine("Processing Text ... " + filename);
                    StreamWriter sw = File.CreateText(folderPath + "\\" + filename + ".txt");
                    //將所有xml文字寫入檔案
                    foreach (var element in matchingElements)
                    {
                        sw.WriteLine(element.Attribute("Name").Value);
                    }
                    sw.Close();


                    //如果不能滑動則停止function
                    if (null == ContentScrollBar)

                    {
                        break;
                    }

                    //如果可以滑動，則將聊天框往上滑，0的屬性代表往上滑
                    if (0 == side)
                    {
                        ContentScrollViewer.SendKeys(Keys.PageUp);
                        //value= 0代表滑到頂，則停止程式執行
                        if (ContentScrollBar.Attribute("Value").Value == "0")
                        {
                            break;
                        }
                    }
                    else
                    {
                        //若爲1，則執行往下滑
                        ContentScrollViewer.SendKeys(Keys.PageDown);
                        if (Math.Abs(double.Parse(ContentScrollBar.Attribute("Maximum").Value) - double.Parse(ContentScrollBar.Attribute("Value").Value)) < 10 )
                        {
                            break;
                        }
                    }
                }
                catch
                {
                    Thread.Sleep(1000);
                    continue;
                }
            }
            //滑倒低，則將程式完成
            Console.WriteLine("Complete");
            return true;
        }

        //點擊所有聯絡人
        public bool autoClickContactPersonScrollScreenShotEachPage(string folderPath, int folderName)
        {
            XDocument doc = null;
            WindowsElement ContactPersonScrollViewer = null;
            while (true)
            {
                try
                {
                    //先找到左邊聯絡人欄位
                    var ScrollViewerElements = session.FindElementsByClassName("ScrollViewer"); //尋找兩個下拉框
                    //如果沒有找到任何欄位，表示程式沒有完整開啓
                    if (ScrollViewerElements.Count == 0)
                    {
                        throw new ArgumentException("Can't Find View");
                    }

                    doc = XDocument.Parse(session.PageSource); //PageSource取得XML結構  XDocument C#提供API

                    //將掃描到的XML中Scrollviweer相關的資訊放到XML element中
                    var xmlScrollViewerElements = doc.Descendants()
                                  .Where(x =>
                                  x.Attribute("ClassName").Value == "ScrollViewer");

                    //確認滑動杆可不可以上下滑動   
                    if (xmlScrollViewerElements.ElementAt(0).Attribute("VerticallyScrollable").Value == "True")
                    {
                        //可以的話就設定聯絡人部分滑動杆的變數
                        ContactPersonScrollViewer = ScrollViewerElements[0];
                    }

                    if (ContactPersonScrollViewer != null)
                    {
                        while (true)
                        {
                            //將滑動杆的XML資訊存爲變數
                            var ScrollBarElements = doc.Descendants()
                                  .Where(x =>
                                  x.Attribute("ClassName").Value == "ScrollBar");


                            //透過檢查ScrollBarElements的Value是否為0來確認是否滑倒頂
                            if (double.Parse(ScrollBarElements.ElementAt(0).Attribute("Value").Value) != 0)
                            {
                                ContactPersonScrollViewer.SendKeys(Keys.PageUp);
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    break;
                }
                catch
                {
                    Thread.Sleep(1000);
                    continue;
                }
            }

            //建立hashtable資料結構，將已點擊過的聯絡人資訊儲存在hashtable的資料結構
            HashSet<string> AlreadlyClickID = new HashSet<string>();

            while (true)
            {
                try
                {
                    if (session.FindElementsByClassName("ScrollViewer").Count == 0)
                    {
                        throw new ArgumentException("Can't Find View");
                    }


                    //將頁面的XML資訊存到doc變數
                    doc = XDocument.Parse(session.PageSource);

                    //source變數儲存原始頁面資訊
                    string source = session.PageSource;

                    //取得所有聯絡人頭像元件
                    var matchingElements = doc.Descendants()
                                  .Where(x =>
                                  x.Attribute("Name").Value != "" &&
                                  x.Attribute("x").Value == "0" &&
                                  x.Attribute("y").Value != "0" &&
                                  x.Attribute("LocalizedControlType").Value == "button" &&
                                  x.Attribute("IsKeyboardFocusable").Value == "False");

                    //變例所有元件
                    foreach (var element in matchingElements)
                    {
                        //取得該元件的的RuntimeId
                        string ID = element.Attribute("RuntimeId").Value;
                        //先確認RuntimeId是否已經被點擊過，若沒被點擊過，先把他加到HASHTABLE
                        if (AlreadlyClickID.Add(ID))
                        {
                            while (true)
                            {
                                try
                                {
                                    //點擊RuntimeId，就可以點解聯絡人
                                    var clickElement = session.FindElementById(ID);
                                    clickElement.Click();
                                    clickElement.Click();
                                    break;
                                }
                                catch
                                {
                                    Thread.Sleep(1000);
                                    continue;
                                }
                            }

                            //最後先建立文件檔案
                            Directory.CreateDirectory(folderPath + "\\" + Convert.ToString(folderName));
                            //執行該聯絡人的聊天室截圖
                            autoScrollScreenShotOnePage(folderPath + "\\" + Convert.ToString(folderName), 0, Int32.MaxValue);
                            folderName++;
                        }
                    }
                }
                catch
                {
                    Thread.Sleep(1000);
                    continue;
                }

                if (null == ContactPersonScrollViewer)
                {
                    break;
                }

                while (true)
                {
                    try
                    {
                        doc = XDocument.Parse(session.PageSource);

                        var ScrollBarElements = doc.Descendants()
                          .Where(x =>
                          x.Attribute("ClassName").Value == "ScrollBar");

                        if (ScrollBarElements.ElementAt(0).Attribute("Value").Value == ScrollBarElements.ElementAt(0).Attribute("Maximum").Value)
                        {
                            return true;
                        }
                        else
                        {
                            ContactPersonScrollViewer.SendKeys(Keys.PageDown);
                        }
                        break;
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            return true;
        }



        public void TearDown()
        {
            // Close the application and delete the session
            if (session != null)
            {
                session = null;
            }
        }

        private void KillProcessesByName(params string[] namesOfProcessesToKill)
        {
            foreach (var name in namesOfProcessesToKill)
            {
                Process[] existingProcesses = Process.GetProcessesByName(name);
                if (existingProcesses.Length > 0)
                {
                    foreach (var p in existingProcesses)
                    {
                        p.Kill();
                    }
                }
            }
        }
    }
}
