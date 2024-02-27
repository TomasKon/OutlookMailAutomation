using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;

class Program
{
    static void Main(string[] args)
    {
        // Instantiate Firefox WebDriver
        FirefoxDriverService service = FirefoxDriverService.CreateDefaultService(@"F:\ֻחרויע\geckodriver-v0.34.0-win64");
        IWebDriver driver = new FirefoxDriver(service);

        // Navigate to Outlook Mail
        driver.Navigate().GoToUrl("https://go.microsoft.com/fwlink/p/?linkid=2125442&clcid=0x409&culture=enus&country=us");

        // Login to Outlook Mail 
        System.Threading.Thread.Sleep(3000);
        driver.FindElement(By.XPath("//input[@id='i0116']")).SendKeys("testsel24@hotmail.com");
        System.Threading.Thread.Sleep(5000);
        driver.FindElement(By.XPath("//button[@id='idSIButton9']")).Click();
        System.Threading.Thread.Sleep(3000);
        driver.FindElement(By.XPath("//input[@id='i0118']")).SendKeys("iwQ&q6x3._4r-Kq");
        System.Threading.Thread.Sleep(3000);
        driver.FindElement(By.XPath("//button[@id='idSIButton9']")).Click();
        System.Threading.Thread.Sleep(3000);
        driver.FindElement(By.XPath("//button[@id='acceptButton']")).Click();
        // Wait for the page to load
        System.Threading.Thread.Sleep(8000);

        driver.FindElement(By.XPath("//button[@id='3']")).Click();
        System.Threading.Thread.Sleep(5000);
        driver.FindElement(By.XPath("//button[@id='643']")).Click();
        System.Threading.Thread.Sleep(5000);

        // Verify that by default "Focused" section is selected
        IWebElement element = driver.FindElement(By.XPath("/html/body/div[2]/div[4]/div/div/div/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div/div[1]"));
       

        if (!element.Selected)
        {
            element.Click();
        }
        
        driver.FindElement(By.XPath("/html/body/div[2]/div[4]/div/div/div/div[2]/div[2]/div/div[3]/div[3]/button[1]")).Click();

        // Wait for the page to load
        System.Threading.Thread.Sleep(5000);
        driver.FindElement(By.XPath("/html/body/div[2]/div[4]/div/div/div/div[2]/div[2]/div/div[3]/div[1]/div[2]/button/span/i")).Click();
        // Select the first two emails using the checkboxes
        driver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div[1]/div")).Click();
        driver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div[1]/div/div")).Click();


        // Verify that the text "2 items selected" is shown
        IWebElement selectedItemText = driver.FindElement(By.XPath("//*[@id='EmptyState_MainMessage']"));
        if (selectedItemText.Text.Equals("2 conversations selected"))
        {
            Console.WriteLine("Verification: '2 conversations selected' text is displayed.");
        }
        else
        {
            Console.WriteLine("Verification Failed: '2 conversations selected' text is not displayed.");
        }

        
        int n = 1; 
        GetEmailDetails(driver, n);

     
    }

    static void GetEmailDetails(IWebDriver driver, int n)
    {
       
        IWebElement sender = driver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[1]"));
        IWebElement subject = driver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div[1]/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[2]"));

        Console.WriteLine($"Sender: {sender.Text}");
        Console.WriteLine($"Subject: {subject.Text}");
    }

}

