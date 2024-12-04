
//Updated Selenium Script for Friday:
const today = 'Friday'; // Force the script to work on Friday data

const { Builder, By } = require('selenium-webdriver');
const xlsx = require('xlsx');

(async function main() {
  const driver = await new Builder().forBrowser('chrome').build();

  try {
    // Load the Excel file
    const filePath = './data.xlsx';
    const workbook = xlsx.readFile(filePath);

    // Set the day to Friday
    const today = 'Friday';
    console.log(`Processing data for: ${today}`);

    // Load the Friday sheet
    const sheet = workbook.Sheets[today];
    if (!sheet) {
      console.error(`No data found for ${today}`);
      return;
    }

    // Parse the sheet data
    const data = xlsx.utils.sheet_to_json(sheet);
    const keywords = data.map(row => row['Keyword']).filter(Boolean);

    if (keywords.length === 0) {
      console.error(`No keywords found for ${today}`);
      return;
    }

    const results = [];

    for (const keyword of keywords) {
      console.log(`Searching for keyword: ${keyword}`);
      await driver.get('https://www.google.com');

      // Handle cookie dialog if present
      try {
        const acceptButton = await driver.findElement(By.xpath('//button[text()="Accept all"]'));
        await acceptButton.click();
      } catch (err) {
        // Ignore if no cookie dialog
      }

      // Search for the keyword
      const searchBox = await driver.findElement(By.name('q'));
      await searchBox.sendKeys(keyword, '\n');

      // Wait for search results
      await driver.sleep(2000);

      // Get search result titles
      const titles = await driver.findElements(By.css('h3'));
      let longestTitle = '';
      let shortestTitle = '';

      for (const titleElement of titles) {
        const text = await titleElement.getText();
        if (text) {
          if (!longestTitle || text.length > longestTitle.length) longestTitle = text;
          if (!shortestTitle || text.length < shortestTitle.length) shortestTitle = text;
        }
      }

      // Store results
      results.push({ Keyword: keyword, LongestOption: longestTitle, ShortestOption: shortestTitle });
    }

    // Update Excel with the results
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
  } catch (error) {
    console.error(`Error: ${error}`);
  } finally {
    await driver.quit();
  }
})();
