import puppeteer from 'puppeteer';
import ExcelJS from 'exceljs';

// Function to scrape TikTok video metadata using Puppeteer
async function scrapeTikTokVideoMetadata(videoUrl: string): Promise<{views: string, likes: string, comments: string, shares: string, timestamp: string} | null> {
    try {
        const browser = await puppeteer.launch();
        const page = await browser.newPage();
        
        // Go to the TikTok video URL
        await page.goto(videoUrl, { waitUntil: 'networkidle2' });
        
        // Wait for the necessary elements to load
        await page.waitForSelector('strong');  // Adjust this if needed for better selectors

        // Scrape the metadata
        const metadata = await page.evaluate(() => {
            const views = document.querySelector('strong[data-e2e="like-count"]')?.textContent || 'N/A';
            const likes = document.querySelector('strong[data-e2e="comment-count"]')?.textContent || 'N/A';
            const comments = document.querySelector('strong[data-e2e="undefined-count"]')?.textContent || 'N/A';
            const shares = document.querySelector('strong[data-e2e="share-count"]')?.textContent || 'N/A';
            
            return {
                views,
                likes,
                comments,
                shares
            };
        });

        // Add the current timestamp
        const timestamp = new Date().toISOString();  // Use ISO string for a clean format
        const metadataWithTime = {
            ...metadata,
            timestamp
        };

        printScrapedData([metadataWithTime]);

        await browser.close();
        return metadataWithTime;

    } catch (error) {
        console.error(`Error scraping metadata for ${videoUrl}: `, error);
        return null;
    }
}

// Function to export data to an Excel sheet and append new rows with timestamps
async function exportToExcel(scrapedData: Array<{views: string, likes: string, comments: string, shares: string, timestamp: string}>) {
    const workbook = new ExcelJS.Workbook();

    // Try loading an existing file, otherwise create a new one
    try {
        await workbook.xlsx.readFile('tiktok_video_metadata.xlsx');
    } catch (error) {
        console.log('Creating a new Excel file...');
    }

    let worksheet = workbook.getWorksheet('TikTok Metadata');
    if (!worksheet) {
        worksheet = workbook.addWorksheet('TikTok Metadata');

        // Add header row if the worksheet is new
        worksheet.columns = [
            { header: 'Views', key: 'views', width: 15 },
            { header: 'Likes', key: 'likes', width: 15 },
            { header: 'Comments', key: 'comments', width: 15 },
            { header: 'Shares', key: 'shares', width: 15 },
            { header: 'Timestamp', key: 'timestamp', width: 25 },
        ];
    }

    // Add data rows
    scrapedData.forEach((data) => {
        worksheet.addRow(data);
    });

    // Write the updated data to the Excel file
    await workbook.xlsx.writeFile('tiktok_video_metadata.xlsx');
    console.log('Data exported to Excel successfully!');
}

// Function to print the scraped data to the console
function printScrapedData(scrapedData: Array<{views: string, likes: string, comments: string, shares: string, timestamp: string}>) {
    console.log('Scraped TikTok Video Metadata:');
    scrapedData.forEach((data, index) => {
        console.log(`Video ${index + 1}:`);
        console.log(`  Views: ${data.views}`);
        console.log(`  Likes: ${data.likes}`);
        console.log(`  Comments: ${data.comments}`);
        console.log(`  Shares: ${data.shares}`);
        console.log(`  Scraped At: ${data.timestamp}`);
        console.log('-------------------------');
    });
}

// Function to periodically scrape and export TikTok video metadata
async function runScraperPeriodically() {
    const videoUrls = [
        'https://www.tiktok.com/@audracoteee/video/7403581594914606378?is_from_webapp=1',  // Example TikTok video URL
    ];

    // Loop every 10 minutes (600000 ms)
    setInterval(async () => {
        console.log('Starting new scraping iteration...');

        const scrapedData: Array<{views: string, likes: string, comments: string, shares: string, timestamp: string}> = [];

        // Loop through each video URL and scrape metadata
        for (const url of videoUrls) {
            const metadata = await scrapeTikTokVideoMetadata(url);
            if (metadata) {
                scrapedData.push(metadata);
            }
        }

        // Print the scraped data to the console
        // printScrapedData(scrapedData);

        // Export the scraped data to Excel
        await exportToExcel(scrapedData);
    }, 10000);  // 10 seconds
}

// Run the scraper periodically until interrupted
runScraperPeriodically().catch(console.error);
