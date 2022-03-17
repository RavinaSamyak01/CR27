package BasePackage;

import org.openqa.selenium.Dimension;
import org.openqa.selenium.Platform;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseInit {
	public static StringBuilder msg = new StringBuilder();
	public static WebDriver ChDriver;

	@BeforeSuite
	public void startup() {
		DesiredCapabilities capabilities = new DesiredCapabilities();
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		// options.addArguments("headless");
		options.addArguments("headless");
		options.addArguments("--incognito");
		options.addArguments("--test-type");
		options.addArguments("--no-proxy-server");
		options.addArguments("--proxy-bypass-list=*");
		options.addArguments("--disable-extensions");
		options.addArguments("--no-sandbox");
		options.addArguments("--headless");
		options.addArguments("window-size=1366x788");
		capabilities.setPlatform(Platform.ANY);
		capabilities.setCapability(ChromeOptions.CAPABILITY, options);
		ChDriver = new ChromeDriver(options);
		// Default size
		Dimension currentDimension = ChDriver.manage().window().getSize();
		int height = currentDimension.getHeight();
		int width = currentDimension.getWidth();
		System.out.println("Current height: " + height);
		System.out.println("Current width: " + width);
		System.out.println("window size==" + ChDriver.manage().window().getSize());

		// Set new size
		Dimension newDimension = new Dimension(1366, 788);
		ChDriver.manage().window().setSize(newDimension);

		// Getting
		Dimension newSetDimension = ChDriver.manage().window().getSize();
		int newHeight = newSetDimension.getHeight();
		int newWidth = newSetDimension.getWidth();
		System.out.println("Current height: " + newHeight);
		System.out.println("Current width: " + newWidth);
	}

	@AfterSuite
	public void end() {
		ChDriver.close();
		ChDriver.quit();
	}
}
