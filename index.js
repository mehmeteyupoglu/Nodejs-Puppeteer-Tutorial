const puppeteer = require("puppeteer");
const readline = require("readline");
const ExcelJS = require("exceljs");

// Create an interface for user input
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// Prompt the user for sector keywords and number of news articles
rl.question("Enter sector keywords: ", (keywords) => {
  rl.question(
    "Enter the number of news articles to extract: ",
    async (numArticles) => {
      try {
        // Launch the browser
        const browser = await puppeteer.launch({
          headless: false,
          defaultViewport: false,
          userDataDir: "./tmp",
        });

        // Open a new page
        const page = await browser.newPage();

        // Go to Google and search for the keywords
        await page.goto("https://www.google.com");
        await page.waitForSelector('textarea[name="q"]'); // Wait for the search input to be available
        await page.type('textarea[name="q"]', keywords);
        await page.keyboard.press("Enter");
        await page.waitForNavigation(); // Wait for the search results page to load

        // Click on the "Haberler" (News) tab
        await page.waitForSelector('a[href*="tbm=nws"]'); // Wait for the "Haberler" tab to be available
        await page.click('a[href*="tbm=nws"]');
        await page.waitForSelector("a.WlydOe"); // Wait for the news results to load

        // Extract the specified number of news links
        const newsLinks = await page.evaluate((numArticles) => {
          const links = [];
          const items = document.querySelectorAll("a.WlydOe");
          for (let i = 0; i < numArticles && i < items.length; i++) {
            const titleElement = items[i].querySelector(".n0jPhd");
            const dateElement = items[i].querySelector(".OSrXXb span");
            if (titleElement && dateElement) {
              links.push({
                title: titleElement.innerText,
                url: items[i].href,
                date: dateElement.innerText,
              });
            }
          }
          return links;
        }, numArticles);

        // Close the browser
        await browser.close();

        // Create a new Excel workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("News");

        // Add column headers
        worksheet.columns = [
          { header: "Title", key: "title", width: 50 },
          { header: "URL", key: "url", width: 100 },
          { header: "Date", key: "date", width: 20 },
        ];

        // Add the news links to the worksheet
        newsLinks.forEach((link) => {
          worksheet.addRow({
            title: link.title,
            url: { text: link.url, hyperlink: link.url },
            date: link.date,
          });
        });

        // Save the workbook to a file
        await workbook.xlsx.writeFile("news.xlsx");

        console.log("News exported to news.xlsx");
      } catch (error) {
        console.error("An error occurred:", error);
      } finally {
        // Close the readline interface
        rl.close();
      }
    }
  );
});
