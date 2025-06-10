const express = require('express');
const axios = require('axios');
const bodyParser = require('body-parser');
const Outscraper = require('outscraper');
require('dotenv').config(); // To access process.env

const app = express();
const port = 3000;
const path = require('path');
const xlsx = require('xlsx'); // Import the xlsx library
const fs = require('fs');


// Set the view engine to EJS
app.set('view engine', 'ejs');

// Set the views folder (optional, defaults to "views")
app.set('views', './views');

// Use body parser to handle POST request body
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

// Serve static files like CSS, JS
app.use(express.static('public'));

// Initialize Outscraper with the API key
let client = new Outscraper(process.env.SCRAPPER_KEY);

// Route to render the form (get request)
app.get('/', (req, res) => {
    // Check if the leadsData.json file exists
    const leadsFilePath = 'leadsData.json';
    let leads = [];

    if (fs.existsSync(leadsFilePath)) {
        // Read and parse the data from the file
        leads = JSON.parse(fs.readFileSync(leadsFilePath, 'utf-8'));
    }

    const leadsFromFile = JSON.parse(fs.readFileSync('leadsData.json', 'utf-8'));

    // Pass the loaded leads to the EJS template
    res.render('index', { data: leadsFromFile });
});

// Route to handle the search and scrape data
app.post('/search', async (req, res) => {
    const { name,area, limit } = req.body;
    const query=name+" "+area
    try {
        // Initiate the search
        const response = await client.googleMapsSearch(query, limit, { language: 'en', region: 'IN' });

        if (response.status === 'Pending') {
            // Poll for the results
            const resultsLocation = response.results_location;

            // Check the status until it's not 'Pending' anymore
            const checkStatus = setInterval(async () => {
                const statusResponse = await axios.get(resultsLocation);

                // If the status is no longer 'Pending', stop polling
                if (statusResponse.data.status !== 'Pending') {
                    clearInterval(checkStatus);
                    const resultData = statusResponse.data.data;

                    // Handle the nested array of results
                    if (Array.isArray(resultData) && resultData.length > 0) {
                        // Flattening the nested arrays
                        const leads = resultData.flat();

                        // Extract relevant data
                        let scrapedData = leads.map(place => {
                            return {
                                name: place.name || 'N/A',
                                phone: place.phone || 'N/A',
                                website: place.site || 'N/A',
                                photosCount: place.photos_count || 0,
                                isJustdial: place.site && place.site.includes('justdial.com') ? 'Yes' : 'No',
                                isTripadvisor: place.site && place.site.includes('tripadvisor.com') ? 'Yes' : 'No',
                                locationLink: place.location_link || '#',
                                address: place.full_address || 'N/A',
                                rating: place.rating || 'N/A'
                            };
                        });

                        // Save the scraped data into leadsData.json
                        // First, try reading existing leads
                        let existingLeads = [];

                        try {
                            existingLeads = JSON.parse(fs.readFileSync('leadsData.json', 'utf-8'));
                        } catch (error) {
                            console.log('No existing leads, starting fresh.');
                        }

                        // Then, add new scraped data on top
                        const updatedLeads = [...scrapedData, ...existingLeads];

                        // Now write the combined leads back
                        fs.writeFileSync('leadsData.json', JSON.stringify(updatedLeads, null, 2));


                        // Render the results
                        const leadsFromFile = JSON.parse(fs.readFileSync('leadsData.json', 'utf-8'));

                        // Pass the loaded leads to the EJS template
                        res.render('index', { data: leadsFromFile });
                    } else {
                        res.render('index', { data: [] });
                    }
                }
            }, 5000); // Poll every 5 seconds
        } else {
            const resultData = response.data;

            // Handle the nested array of results
            if (Array.isArray(resultData) && resultData.length > 0) {
                // Flattening the nested arrays
                const leads = resultData.flat();
// isJustdial: place.site && place.site.includes('justdial.com') ? 'Yes' : 'No',
// isTripadvisor: place.site && place.site.includes('tripadvisor.com') ? 'Yes' : 'No',
                // Extract relevant data
                let scrapedData = leads.map(place => ({
                    name: place.name || 'N/A',
                    phone: place.phone || 'N/A',
                    website: place.site || 'N/A',
                    photosCount: place.photos_count || 0,
                    
                    locationLink: place.location_link || '#',
                    address: place.full_address || 'N/A',
                    rating: place.rating || 'N/A'
                }));

                // Save the scraped data into leadsData.json
                // First, try reading existing leads
                let existingLeads = [];

                try {
                    existingLeads = JSON.parse(fs.readFileSync('leadsData.json', 'utf-8'));
                } catch (error) {
                    console.log('No existing leads, starting fresh.');
                }

                // Then, add new scraped data on top
                const updatedLeads = [...scrapedData, ...existingLeads];

                // Now write the combined leads back
                fs.writeFileSync('leadsData.json', JSON.stringify(updatedLeads, null, 2));


                // Render the results
                res.render('index', { data: scrapedData });
            } else {
                const leadsFromFile = JSON.parse(fs.readFileSync('leadsData.json', 'utf-8'));

                // Pass the loaded leads to the EJS template
                res.render('index', { data: leadsFromFile });
            }
        }
    } catch (error) {
        console.error(error);
        const leadsFromFile = JSON.parse(fs.readFileSync('leadsData.json', 'utf-8'));

        // Pass the loaded leads to the EJS template
        res.render('index', { data: leadsFromFile });
    }
});

// Route to export the leads data to Excel
const ExcelJS = require('exceljs');


app.get('/export', async (req, res) => {
    const leadsFilePath = 'leadsData.json';

    if (!fs.existsSync(leadsFilePath)) {
        return res.status(404).send('No leads data found');
    }

    const leads = JSON.parse(fs.readFileSync(leadsFilePath, 'utf-8'));
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Leads');

    // Define headers from keys
    const headers = Object.keys(leads[0]);

    // Create styled header row
    sheet.addRow(headers);
    const headerRow = sheet.getRow(1);
    headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF1F497D' } // dark blue
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
        };
    });

    // Add data rows with centered alignment
    leads.forEach((lead) => {
        const row = sheet.addRow(headers.map(h => lead[h]));
        row.eachCell((cell) => {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });
    });

    // Auto-adjust column widths
    sheet.columns.forEach((col, i) => {
        const maxLength = Math.max(
            headers[i].length,
            ...sheet.getColumn(i + 1).values.map(val => (val ? val.toString().length : 0))
        );
        col.width = maxLength + 4;
    });

    const filePath = path.join(__dirname, 'leadsData.xlsx');
    await workbook.xlsx.writeFile(filePath);

    res.download(filePath, 'leadsData.xlsx', (err) => {
        if (!err) fs.unlinkSync(filePath);
    });
});




app.post('/clear', (req, res) => {
    fs.writeFileSync('leadsData.json', JSON.stringify([], null, 2)); // Clear by writing empty array
    res.redirect('/'); // After clearing, redirect back to home
});

// Start the server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
