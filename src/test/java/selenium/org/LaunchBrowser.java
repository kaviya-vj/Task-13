package selenium.org;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class LaunchBrowser {

	public static void main(String[] args) throws InterruptedException {
		// set the property
		WebDriverManager.chromedriver().setup();
		// Initialise the WebDriver for chrome
		WebDriver driver = new ChromeDriver();
		// path or selenium.dev
		driver.get("https://selenium.dev/");
		// to maximize the chrome page
		driver.manage().window().maximize();
		// wait for the result
		Thread.sleep(2000);
		// to close the window
		driver.close();

	}

}
