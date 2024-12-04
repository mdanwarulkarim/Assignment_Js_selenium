const { Builder, By } = require('selenium-webdriver');
const xlsx = require('xlsx');
const fs = require('fs');

(async function main() {
  const driver = await new Builder().forBrowser('chrome').build();

  try {
    // Load the Excel file and get today's sheet
    const filePath = './data.xlsx'; // Replace with the path to your file
    const workbook = xlsx.readFile(filePath);
    const today = new Date().toLocaleString('en-us', { weekday: 'long' });
    const sheet = workbook.Sheets[today];

    if (!sheet) {
      console.log(`No data found for ${today}.`);
      return;
    }

    // Parse keywords from today's sheet
    const data = xlsx.utils.sheet_to_json(sheet);
    const keywords = data.map(row => row['Keyword']).filter(Boolean);

    if (!keywords || keywords.length === 0) {
      console.log(`No keywords available for ${today}.`);
      return;
    }

    const results = [];

    for (const keyword of keywords) {
      await driver.get('https://www.google.com');

      // Accept cookies if necessary
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

      results.push({ Keyword: keyword, LongestOption: longestTitle, ShortestOption: shortestTitle });
    }

    // Update Excel with results
    const updatedData = data.map(row => {
      const result = results.find(r => r.Keyword === row['Keyword']);
      return {
        ...row,
        'Longest Option': result?.LongestOption || '',
        'Shortest Option': result?.ShortestOption || '',
      };
    });

    const updatedSheet = xlsx.utils.json_to_sheet(updatedData);
    workbook.Sheets[today] = updatedSheet;
    xlsx.writeFile(workbook, filePath);

    console.log(`Results updated in ${filePath}.`);
  } catch (err) {
    console.error('Error:', err);
  } finally {
    await driver.quit();
  }
})();
