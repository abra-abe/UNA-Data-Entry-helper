const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const axios = require('axios');
const path = require('path');
const app = express();
const port = 3000;

//configuring the view-engine
app.set('view-engine', 'ejs');

//configuring body parser(from express)
app.use(express.json());
app.use(express.urlencoded({extended: true}));

app.use(express.static(path.join(__dirname, 'public')));

//configuring multer
const storage = multer.memoryStorage();
const upload = multer({storage: storage});

//routes
app.get('/', (req, res) => {
    res.render('index.ejs');
});

app.post('/check-websites', upload.single('file'), async (req, res) => {
    try{
        const fileBuffer = req.file.buffer;
        const organizations = await readXls(fileBuffer);

        //function to check if the website exists
        async function checkWebsite(orgName) {
            const baseUrl = 'https://';  // Adjust this based on the protocol (http or https)
        
            // Assuming orgName is abbreviated, you can generate alternative URLs
            const alternativeUrls = [
                `${baseUrl}${orgName}.com`,
                `${baseUrl}${orgName}.co.tz`,
                `${baseUrl}${orgName}.ac.tz`,
                `${baseUrl}${orgName}.or.tz`,
                // Add more alternative URL patterns as needed
            ];
        
            // Function to check if a website exists
            async function tryUrls(urls) {
                const matchedUrls = [];
                for (const url of urls) {
                    try {
                        const response = await axios.head(url);
                        if (response.status === 200) {
                            console.log(url);
                            matchedUrls.push(url)
                        }
                    } catch (error) {
                        // Ignore errors and try the next URL
                    }
                }
                return matchedUrls;
            }
        
            return tryUrls(alternativeUrls);
        }        

        //add a new field for website existence
        const updatedOrganizations = await Promise.all(organizations.map(async (org) => {
            const websites = await checkWebsite(org['Organization Name']);
            org['Website Exists'] = websites.length > 0;
            org['Websites'] = websites;
            return org;
        }));

        res.render('result.ejs', { organizations: updatedOrganizations });
    }catch(err){
        res.status(500).send('Internal Server Error: Error processing file upload');
    }
});

app.get('/download', (req, res) => {
    const updatedCsv = parse(updatedOrganizations, { header: true }).data;
    const csvContent = parse(updatedCsv, { header: true, quotes: true }).csv;

    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename=updated_document.csv');
    res.send(csvContent);
});

//function to read xls document
function readXls(fileBuffer) {
    try {
        const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
        const result = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        console.log('this is working...');
        return Promise.resolve(result);
    } catch (error) {
        console.error(error);
        return Promise.reject(error);
    }
}

app.listen(port, () => {
    console.log(`Server is listening on port: ${port}`);
})