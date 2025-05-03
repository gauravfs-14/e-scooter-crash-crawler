# E-Scooter Crash Crawler

## Description

The E-Scooter Crash Crawler is a JavaScript-based project designed to scrape and analyze data related to e-scooter crashes. This tool aims to gather insights and provide meaningful statistics to improve safety and awareness around e-scooter usage.

## Features

- Web scraping for e-scooter crash data.
- Data deduplication using fingerprints.
- Exporting data to Excel format.
- Configurable search terms and date ranges.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/gauravfs-14/e-scooter-crash-crawler
   ```
2. Navigate to the project directory:
   ```bash
   cd e-scooter-crash-crawler
   ```
3. Install dependencies:
   ```bash
   npm install
   ```

## Configuration

1. Open the `index.js` file.
2. Update the `config` object as needed:
   - **Search Terms**: Add or modify the `searchTerms` array to include additional keywords.
   - **Date Range**: Adjust `yearsToSearch` to change the range of articles to search.
   - **Pages Per Term**: Modify `pagesPerTerm` to control the number of Google News pages to scrape per search term.
   - **Delay Between Requests**: Adjust `delayBetweenRequests` to avoid being flagged as a bot.
   - **Output Files**: Ensure `outputFile` and `fingerprintsFile` paths are correct.

## Usage

1. Run the crawler:
   ```bash
   npm run start
   ```
2. The crawler will:
   - Search Google News for the specified terms.
   - Scrape articles and extract relevant data.
   - Save the data to an Excel file (`escooter_crash_news.xlsx`).
   - Store fingerprints in a JSON file (`fingerprints.json`).

## Output

- The collected data will be saved in an Excel file (`escooter_crash_news.xlsx`) with the following columns:
  - News Media Name
  - Date
  - Title of the News
  - Descriptive Text
  - URL
- Fingerprints of processed articles will be stored in `fingerprints.json` to avoid duplicates in future runs.

## Requirements

- Node.js (v14 or higher)
- npm (v6 or higher)

## Error Handling

- If the crawler encounters errors:
  - Check the console logs for details.
  - Ensure the internet connection is stable.
  - Verify that the search terms and URLs are valid.

## Future Enhancements

- Add support for additional languages or regions.
- Integrate a database (e.g., MongoDB) for better scalability.
- Use machine learning to classify articles based on relevance.

## Technical Stack

- **Programming Language**: JavaScript (Node.js)
- **Libraries Used**:
  - Puppeteer: For web scraping and automation.
  - Axios: For HTTP requests.
  - Cheerio: For HTML parsing.
  - XLSX: For Excel file generation.
  - Moment: For date manipulation.
- **File Formats**:
  - Output: Excel file (`escooter_crash_news.xlsx`)
  - Fingerprints: JSON file (`fingerprints.json`)
