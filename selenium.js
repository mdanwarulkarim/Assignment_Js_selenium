const { Builder, By } = require('selenium-webdriver');
const xlsx = require('xlsx');
const fs = require('fs');

(async function main() {
  const driver = await new Builder().forBrowser('chrome').build();

  try {
    // Get today's day of the week
    const today = new Date().toLocaleString('en-us', { weekday: 'long' }).toLowerCase();

    // Simulated data for the week
    const dataByDay = {
      monday: ["keyword1", "keyword2"],
      tuesday: ["keyword3", "keyword4"],
      // Add more days here
    };

    const keywords = dataByDay[today];

    if (!keywords || keywords.length === 0) {
      console.log(`No keywords available for ${today}.`);
      return;
    }

    const results = [];

    for (const keyword of keywords) {
      await driver.get('https://www.google.com');

      // Accept cookies if needed
      try {
        const acceptBtn = await driver.findElement(By.xpath('//button[text()="Accept all"]'));
        await acceptBtn.click();
      } catch (err) {
        // No cookie prompt
      }

      // Search for the keyword
      const searchBox = await driver.findElement(By.name('q'));
      await searchBox.sendKeys(keyword, '\n');

      // Wait for search results to load
      await driver.sleep(2000);

      // Get all search result titles
      const titles = await driver.findElements(By.css('h3'));

      let longestTitle = '';
      let shortestTitle = '';

      for (const title of titles) {
        const text = await title.getText();
        if (text) {
          if (!longestTitle || text.length > longestTitle.length) longestTitle = text;
          if (!shortestTitle || text.length < shortestTitle.length) shortestTitle = text;
        }
      }

      results.push({ keyword, longestTitle, shortestTitle });
    }

    // Save to Excel
    const workbook = xlsx.utils.book_new();
    const worksheetData = [['Keyword', 'Longest Title', 'Shortest Title']];
    results.forEach(({ keyword, longestTitle, shortestTitle }) => {
      worksheetData.push([keyword, longestTitle, shortestTitle]);
    });

    const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Results');
    const filePath = `results_${today}.xlsx`;
    xlsx.writeFile(workbook, filePath);

    console.log(`Results saved to ${filePath}`);
  } catch (err) {
    console.error('Error:', err);
  } finally {
    await driver.quit();
  }
})();
