package com.practice;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.*;

import java.time.Duration;

public class DataProviderExcel{
    private static ThreadLocal<WebDriver> driver = new ThreadLocal<>();
    private static ThreadLocal<WebDriverWait> wait = new ThreadLocal<>();

    @Test(dataProvider = "excelData", dataProviderClass = DPExcel1.class)
    public void search(String keyWord1, String keyWord2) throws InterruptedException {
        WebDriver localDriver = driver.get();
        WebDriverWait localWait = wait.get();

        WebElement txtBox = localWait.until(ExpectedConditions.visibilityOfElementLocated(By.name("q")));
        txtBox.sendKeys(keyWord1 + " " + keyWord2);
        System.out.println("Keyword entered: " + keyWord1 + " and " + keyWord2);
        txtBox.sendKeys(Keys.ENTER);
        System.out.println("Search result displayed");

        Thread.sleep(3000);
        String searchPageTitle = localDriver.getTitle();
        System.out.println("Page title: " + searchPageTitle);

        Assert.assertTrue(searchPageTitle.toLowerCase().contains(keyWord1.toLowerCase()),
                "Search is not successful - Title does not contain: " + keyWord1);
    }

    @BeforeMethod
    public void beforeMethod() {
        System.out.println("Start of the test");  
        ChromeOptions options = new ChromeOptions();
        options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
        options.addArguments("--disable-blink-features=AutomationControlled");
        options.setExperimentalOption("excludeSwitches", new String[]{"enable-automation"});
        options.setExperimentalOption("useAutomationExtension", false);

        WebDriver localDriver = new ChromeDriver(options);
        localDriver.manage().window().maximize();
        localDriver.get("https://google.co.in/");

        driver.set(localDriver);
        wait.set(new WebDriverWait(localDriver, Duration.ofSeconds(20)));
    }

    @AfterMethod
    public void afterMethod() {
        if (driver.get() != null) {
            driver.get().quit();
            driver.remove();
            wait.remove();
        }
    }
}