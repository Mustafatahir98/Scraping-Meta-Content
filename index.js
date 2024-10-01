const axios = require('axios');
const cheerio = require('cheerio');
const { parseStringPromise } = require('xml2js');
const ExcelJS = require('exceljs');
const fs = require('fs/promises');
const path = require('path');
const nodemailer = require('nodemailer');
 
require('dotenv').config();
 
const OLD_DATA_FILE = path.join(__dirname, 'oldScrapedData.json');
// Fetch and parse the sitemap
async function fetchSitemap(url) {
  try {
    const response = await axios.get(url);
    const sitemap = await parseStringPromise(response.data);
    return sitemap.urlset.url.map(u => u.loc[0]);
  } catch (error) {
    console.error(`Error fetching sitemap from ${url}:`, error.message);
    return [];
  }
}
 
async function scrapePage(url) {
    try {
      const response = await axios.get(url);
      const $ = cheerio.load(response.data);
      let metaData = {
        url,
        locale: $('meta[property="og:locale"]').attr('content'),
        type: $('meta[property="og:type"]').attr('content'),
        title: $('meta[property="og:title"]').attr('content'),
        description: $('meta[property="og:description"]').attr('content'),
        siteName: $('meta[property="og:site_name"]').attr('content'),
        updatedTime: $('meta[property="og:updated_time"]').attr('content'),
        image: $('meta[property="og:image"]').attr('content'),
        imageWidth: $('meta[property="og:image:width"]').attr('content'),
        imageHeight: $('meta[property="og:image:height"]').attr('content'),
        imageAlt: $('meta[property="og:image:alt"]').attr('content'),
        imageType: $('meta[property="og:image:type"]').attr('content'),
        video: $('meta[property="og:video"]').attr('content'),
        videoDuration: $('meta[property="video:duration"]').attr('content'),
        schema: getSchemaTypes($),
        robots: $('meta[name="robots"]').attr('content'),
        isIndexable: function() {
          var robotsContent = $('meta[name="robots"]').attr('content');
          return !(robotsContent && /noindex/i.test(robotsContent));
      }(),
      };
      return metaData;
    } catch (error) {
      console.error(`Error scraping ${url}:`, error.message);
      return null;
    }
  }
  function getSchemaTypes($) {
    // Find script tags that contain application/ld+json and parse them
    let schemaTypes = [];
    $('script[type="application/ld+json"]').each((i, elem) => {
        try {
            const data = JSON.parse($(elem).html());
            // Traverse through the parsed JSON data and extract schema names
            if (data['@graph']) {
                data['@graph'].forEach(item => {
                    if (item['@type']) {
                        schemaTypes.push(item['@type']);
                    }
                });
            } else if (data['@type']) { // Directly look for @type if no @graph
                schemaTypes.push(data['@type']);
            }
        } catch (e) {
            console.error('Error parsing JSON-LD schema:', e.message);
        }
    });
    // Join all schema types with a comma and return the result
    return schemaTypes.join(', ');
}
  async function saveNewData(data) {
    await fs.writeFile(OLD_DATA_FILE, JSON.stringify(data, null, 2));
  }
  
  async function readOldData() {
    try {
      const data = await fs.readFile(OLD_DATA_FILE, 'utf8');
      return JSON.parse(data);
    } catch (error) {
      console.error(`Error reading old data file: ${error.message}`);
      return [];
    }
  }
 

  async function readExcelFile(filePath) {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet('Scraped Data');
      
      if (!worksheet) {
        throw new Error(`Worksheet 'Scraped Data' does not exist in file ${filePath}`);
      }
      
      const data = [];
      worksheet.eachRow((row) => {
        const rowData = {};  // rowData needs to be declared here
        row.eachCell((cell, colNumber) => {
          rowData[worksheet.columns[colNumber - 1].key] = cell.value;
        });
        data.push(rowData);
      });
      
      return data;
    } catch (error) {
      console.error(`Error reading Excel file at ${filePath}:`, error.message);
      return [];
    }
  }
  
  
async function createAndCompareExcelFile(newData, fileName) {
    console.log(`Creating Excel file with new data: ${fileName}`);
    const oldData = await readOldData();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Scraped Data');
  
    // Setup headers based on newData structure
    worksheet.columns = Object.keys(newData[0]).map(key => ({
      header: key.charAt(0).toUpperCase() + key.slice(1),
      key: key,
      width: 20
    }));
  
    // Populate rows with new data and compare for differences
    newData.forEach(newItem => {
      const row = worksheet.addRow(newItem);
      const oldItem = oldData.find(old => old.url === newItem.url);
  
      if (oldItem) {
        Object.keys(newItem).forEach(key => {
          const cell = row.getCell(key);
          const newVal = newItem[key] ? newItem[key].toString().trim() : '';
          const oldVal = oldItem[key] ? oldItem[key].toString().trim() : '';
          if (newVal !== oldVal) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFF00' }  // Yellow fill for differences
            };
          }
        });
      }
    });
    console.log(`Writing to file: ${fileName}`);
    await workbook.xlsx.writeFile(fileName);
    console.log(`Excel file has been created/updated: ${fileName}`);
  }
 
// Main function
// const url = require('url');
 
async function main() {
  // Read old data before starting the scraping process
  const oldData = await readOldData();
  const sitemapUrls = [
    'https://www.personaldrivers.com/MySecureSitemap/page-sitemap.xml',
    'https://www.personaldrivers.com/MySecureSitemap/page-sitemap.xml',
    'https://www.professionaldrivers.com/post-sitemap.xml',
    'https://www.professionaldrivers.com/page-sitemap.xml',
    'https://amride.com/page-sitemap.xml',
    'https://amride.com/services-sitemap.xml',
    // 'https://www.idriveyourcar.com/sitemap.xml'
    // 'https://www.blacklane.com/sitemap.xml',
    // 'https://www.staging1.theplantconcierge.com/post-sitemap.xml',
    // 'https://www.staging1.theplantconcierge.com/page-sitemap.xml'
    // 'https://amride.com/page-sitemap.xml',
    // 'https://www.theperfectlawn.com/post-sitemap.xml',
    // 'https://www.theperfectlawn.com/page-sitemap.xml'
    // ... other sitemap URLs
  ];
  // Group URLs by their domain name
  const groupedUrls = sitemapUrls.reduce((acc, sitemapUrl) => {
    const hostname = new URL(sitemapUrl).hostname;
    acc[hostname] = [...(acc[hostname] || []), sitemapUrl];
    return acc;
  }, {});
 
  for (const [hostname, urls] of Object.entries(groupedUrls)) {
    const safeHostname = hostname.replace(/\W+/g, '_'); // Replace non-word chars with underscore
    const dateSuffix = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
    const fileName = `ScrapedData_${safeHostname}_${dateSuffix}.xlsx`;
    const filePath = path.join(__dirname, fileName);
 
    let oldData = [];
    try {
      await fs.access(filePath);
      oldData = await readExcelFile(filePath);
    } catch (error) {
      console.log(`No previous file found for ${hostname}:`, error.message);
    }
 
    let allResults = [];
    for (const sitemapUrl of urls) {
      const singleSitemapUrls = await fetchSitemap(sitemapUrl);
      const scrapePromises = singleSitemapUrls.map(url => scrapePage(url));
      const results = (await Promise.allSettled(scrapePromises))
        .filter(result => result.status === 'fulfilled')
        .map(result => result.value)
        .filter(value => value !== null);
 
      allResults = [...allResults, ...results];
    }
 
    if (allResults.length > 0) {
        await createAndCompareExcelFile(allResults, filePath); 
        await saveNewData(allResults);
        // Now you can convert the new Excel file to HTML and send it.
        const htmlContent = await convertExcelToHTML(filePath);
        await sendEmailWithAttachment(filePath, htmlContent);
      } else {
        console.log(`No data was scraped for domain: ${hostname}`);
      }
  }
}
 
 
async function convertExcelToHTML(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet('Scraped Data');

  // Start the HTML table
  let html = '<table border="2"><tr>';

  // Add the headers
  worksheet.getRow(1).eachCell((cell) => {
    html += `<th>${cell.value}</th>`;
  });
  html += '</tr>';

  // Add the data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header row
    html += '<tr>';
    row.eachCell({ includeEmpty: true }, (cell) => {
      // Check if the cell has a fill indicating a change
      let style = '';
      if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor && cell.fill.fgColor.argb === 'FFFFFF00') {
        style = ' style="background-color: yellow;"'; // Apply yellow background if changes are detected
      }
      const cellValue = cell.value instanceof Date ? cell.value.toLocaleDateString() : cell.value; // Handle date formatting
      html += `<td${style}>${cellValue || ''}</td>`; // Display the cell value, using empty string if null
    });
    html += '</tr>';
  });

  html += '</table>';
  return html;
}

// Function to send an email with the Excel file attachment
async function sendEmailWithAttachment(filePath) {
  const htmlTable = await convertExcelToHTML(filePath);
  const transporter = nodemailer.createTransport({
    host: "smtp-mail.outlook.com",
    port: 587,
    secure: false, // true for 465, false for other ports
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
    tls: {
      rejectUnauthorized: false // Adjust TLS settings to meet your requirements
    }
  });
 
  const mailOptions = {
    from: `"Dilshad Delawalla" <${process.env.EMAIL_USER}>`, // sender address
    to: 'mustafa@delawallabiz.com' ,  // list of receivers
    subject: 'Scraped Data Excel File', // Subject line
    text: 'Attached is the latest scraped data Excel file.', // plain text body
    html: htmlTable,
    attachments: [{
      filename: path.basename(filePath),
      path: filePath
    }]
  };
 
  await transporter.sendMail(mailOptions);
}
 
main().catch(console.error);