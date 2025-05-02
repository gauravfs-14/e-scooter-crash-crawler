const axios = require("axios");
const cheerio = require("cheerio");
const XLSX = require("xlsx");
const puppeteer = require("puppeteer");
const moment = require("moment");
const fs = require("fs");

const config = {
  searchTerms: [
    "e-scooter crash",
    "electric scooter accident",
    "scooter collision",
    "e-scooter injury",
    "scooter fatality",
    "micromobility crash",
    "electric scooter death",
    "e-scooter trauma center",
    "e-scooter hospitalization",
    "scooter head injury",
    "e-scooter spinal injury",
    "electric scooter accident victim",
    "scooter hit by car",
    "e-scooter broken bones",
    "e-scooter hospital admission",
    "e-scooter safety concerns",
    "electric scooter lawsuit",
    "e-scooter regulation",
    "e-scooter public health",
    "scooter accident news",
    "electric scooter danger",
    "micromobility injury statistics",
    "scooter recall news",
    "scooter rental accident",
    "shared scooter crash",
    "e-scooter dui crash",
    "electric scooter police report",
    "scooter traffic citation",
    "scooter rider killed",
    "e-scooter hit and run",
    "fatal scooter collision",
    "e-scooter child injury",
    "scooter elderly crash",
    "e-scooter injuries downtown",
    "campus e-scooter accident",
    "e-scooter school zone crash",
    "e-scooter pedestrian struck",
    "bike lane scooter crash",
  ],
  yearsToSearch: 4,
  outputFile: "escooter_crash_news.xlsx",
  fingerprintsFile: "fingerprints.json",
  pagesPerTerm: 5,
  delayBetweenRequests: 2000,
  userAgent:
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36",
};

// Load or initialize fingerprint store
let fingerprints = new Set();
if (fs.existsSync(config.fingerprintsFile)) {
  fingerprints = new Set(
    JSON.parse(fs.readFileSync(config.fingerprintsFile, "utf-8"))
  );
}

const results = [];
const processedUrls = new Set();

function generateFingerprint(title, content) {
  const simplifiedTitle = title
    .toLowerCase()
    .replace(/[^\w\s]/g, "")
    .trim();
  const simplifiedContent = content
    .toLowerCase()
    .replace(/[^\w\s]/g, "")
    .substring(0, 150)
    .trim();
  return `${simplifiedTitle}-${simplifiedContent}`;
}

function isDuplicate(url, title, content) {
  const fingerprint = generateFingerprint(title, content);
  if (processedUrls.has(url) || fingerprints.has(fingerprint)) {
    console.log(`🟡 Skipping duplicate: ${url}`);
    return true;
  }
  processedUrls.add(url);
  fingerprints.add(fingerprint);
  return false;
}

function getDateRange() {
  const end = moment();
  const start = moment().subtract(config.yearsToSearch, "years");
  return {
    startDate: start.format("MM/DD/YYYY"),
    endDate: end.format("MM/DD/YYYY"),
  };
}

async function searchGoogleNews(searchTerm) {
  const { startDate, endDate } = getDateRange();
  const browser = await puppeteer.launch({
    headless: true,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-blink-features=AutomationControlled",
    ],
  });

  const page = await browser.newPage();
  await page.setUserAgent(config.userAgent);
  await page.setViewport({ width: 1280, height: 800 });
  await page.evaluateOnNewDocument(() => {
    Object.defineProperty(navigator, "webdriver", { get: () => false });
  });

  const allLinks = new Set();
  for (let i = 0; i < config.pagesPerTerm; i++) {
    const start = i * 10;
    const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(
      searchTerm
    )}&tbm=nws&tbs=cdr:1,cd_min:${startDate},cd_max:${endDate}&start=${start}`;

    console.log(`🔎 Page ${i + 1}: ${searchUrl}`);
    await page.goto(searchUrl, { waitUntil: "networkidle2" });
    await new Promise((res) => setTimeout(res, 2000));

    const links = await page.evaluate(() => {
      return Array.from(document.querySelectorAll("a"))
        .map((a) => a.href)
        .filter(
          (href) =>
            href.startsWith("http") &&
            !href.includes("google.com") &&
            !href.includes("/settings") &&
            !href.includes("/policies")
        );
    });

    links.forEach((link) => allLinks.add(link));
  }

  await browser.close();
  return Array.from(allLinks);
}

async function scrapeArticle(url) {
  try {
    console.log(`📰 Scraping: ${url}`);
    const response = await axios.get(url, {
      headers: { "User-Agent": config.userAgent },
      timeout: 10000,
    });

    const $ = cheerio.load(response.data);
    const title = $("h1").first().text().trim() || $("title").text().trim();

    let dateText =
      $('meta[property="article:published_time"]').attr("content") ||
      $("time").first().attr("datetime") ||
      $("time").first().text() ||
      $('meta[name="date"]').attr("content") ||
      $('[class*="date"], [class*="time"], [class*="publish"]').first().text();

    let date = moment(dateText, moment.ISO_8601, true).isValid()
      ? moment(dateText).format("YYYY-MM-DD")
      : "Unknown";

    if (date === "Invalid date") date = "Unknown";

    const paragraphs = $("p")
      .map((i, el) => $(el).text().trim())
      .get();
    const content = paragraphs.filter((p) => p.length > 50).join("\n\n");

    if (!content || content.length < 50 || isDuplicate(url, title, content)) {
      return null;
    }

    return {
      mediaName: new URL(url).hostname.replace("www.", ""),
      date,
      title,
      content: content.substring(0, 5000),
      url,
    };
  } catch (error) {
    console.error(`❌ Failed to scrape ${url}: ${error.message}`);
    return null;
  }
}

function saveToExcel(newData) {
  let existingData = [];

  if (fs.existsSync(config.outputFile)) {
    const workbook = XLSX.readFile(config.outputFile);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    existingData = XLSX.utils.sheet_to_json(sheet);
  }

  const combined = [...existingData, ...newData];

  // Remove duplicates by URL
  const uniqueMap = new Map();
  combined.forEach((item) => {
    uniqueMap.set(item.URL, item); // Later ones overwrite
  });

  const deduped = Array.from(uniqueMap.values());

  const worksheet = XLSX.utils.json_to_sheet(deduped);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Crash News");

  worksheet["!cols"] = [
    { wch: 25 },
    { wch: 12 },
    { wch: 60 },
    { wch: 100 },
    { wch: 50 },
  ];

  workbook.Props = {
    Title: "E-Scooter Crash News Dataset",
    Author: "E-Scooter News Bot",
    CreatedDate: new Date(),
  };

  XLSX.writeFile(workbook, config.outputFile);
  console.log(
    `✅ Saved ${deduped.length} total articles to "${config.outputFile}"`
  );
}

async function main() {
  for (const term of config.searchTerms) {
    const links = await searchGoogleNews(term);

    for (const url of links) {
      const article = await scrapeArticle(url);
      if (article) {
        results.push({
          "News Media Name": article.mediaName,
          Date: article.date,
          "Title of the News": article.title,
          "Descriptive Text": article.content,
          URL: article.url,
        });
      }
      await new Promise((res) => setTimeout(res, config.delayBetweenRequests));
    }
  }

  saveToExcel(results);

  // Save updated fingerprints
  fs.writeFileSync(
    config.fingerprintsFile,
    JSON.stringify([...fingerprints], null, 2)
  );
  console.log(
    `🧠 Stored ${fingerprints.size} total fingerprints in ${config.fingerprintsFile}`
  );
}

main().catch((err) => console.error("❌ Fatal error:", err.message));
