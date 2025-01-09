const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');

const app = express();
const PORT = 3000;

// Middleware to serve static files
app.use(express.static(path.join(__dirname, '../public')));
app.use(bodyParser.urlencoded({ extended: true }));

// Set view engine to EJS
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, '../views'));

// Function to check Employee ID in Excel
function getSeatNumber(employeeId) {
    const workbook = xlsx.readFile(path.join(__dirname, '../seatPage.xlsx'));
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);
    for (const row of data) {
        if (row.EMIP == employeeId) {
            return [row.No,row.Seat];
        }
    }
    return null;
}

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../index.html'));
});

app.post('/submit', (req, res) => {
    const employeeId = req.body.emp_id;
    const seatNumber = getSeatNumber(employeeId);
     if (seatNumber[0]) {
        if(seatNumber[1] === "Chair"){
            const numberOF =seatNumber[0] 
            res.render('seat', { numberOF});
        }
        else{
            const numberOF =seatNumber[0] 
            res.render('table', { numberOF});
        }
       
    } else {
         res.send(`
             <h1 style="text-align:center; color:red;">Invalid Employee ID</h1>
             <a href="/" style="display:block; text-align:center; margin-top:20px;">Go Back</a>
         `);
 }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
