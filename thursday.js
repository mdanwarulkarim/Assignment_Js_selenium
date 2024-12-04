const today = 'Thursday'; // Forces script to process Thursday data

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

  let longestTitle = '';
let shortestTitle = '';

for (const titleElement of titles) {
  const text = await titleElement.getText();
  if (text) {
    if (!longestTitle || text.length > longestTitle.length) longestTitle = text;
    if (!shortestTitle || text.length < shortestTitle.length) shortestTitle = text;
  }
}
