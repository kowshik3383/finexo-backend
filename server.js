// Import dependencies
const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const XLSX = require('xlsx');
const bodyParser = require('body-parser');
const cors = require('cors');
const dotenv = require('dotenv')
dotenv.config()
// Initialize the app
const app = express();
app.use(cors());
app.use(bodyParser.json());

// Connect to MongoDB
mongoose.connect(process.env.MONGO_URI, {

});

const db = mongoose.connection;
db.on('error', console.error.bind(console, 'Connection error:'));
db.once('open', () => {
  console.log('Connected to MongoDB');
});

// Define the schema and model for data
const dataSchema = new mongoose.Schema({
  sheetName: String,
  data: Array,
});

const DataModel = mongoose.model('Data', dataSchema);

// Configure file upload using multer
const storage = multer.memoryStorage();
const upload = multer({ storage });

// Endpoint to upload and process Excel file
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetNames = workbook.SheetNames;

    for (const sheetName of sheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // Save data to MongoDB
      const data = new DataModel({ sheetName, data: jsonData });
      await data.save();
    }

    res.status(200).send({ message: 'File uploaded and processed successfully', sheetNames });
  } catch (error) {
    console.error(error);
    res.status(500).send({ message: 'Error processing file', error });
  }
});

// Endpoint to fetch data for a specific sheet
app.get('/api/data/:sheetName', async (req, res) => {
  try {
    const { sheetName } = req.params;
    const data = await DataModel.findOne({ sheetName });

    if (!data) {
      return res.status(404).send({ message: 'Sheet not found' });
    }

    res.status(200).send(data);
  } catch (error) {
    console.error(error);
    res.status(500).send({ message: 'Error fetching data', error });
  }
});

// Endpoint to update a row in a sheet
app.put('/api/data', async (req, res) => {
  try {
    const { sheetName, rowIndex, updatedRow } = req.body;
    const data = await DataModel.findOne({ sheetName });

    if (!data) {
      return res.status(404).send({ message: 'Sheet not found' });
    }

    data.data[rowIndex] = updatedRow;
    await data.save();

    res.status(200).send({ message: 'Row updated successfully' });
  } catch (error) {
    console.error(error);
    res.status(500).send({ message: 'Error updating row', error });
  }
});

// Endpoint to delete a row from a sheet
app.delete('/api/data/:sheetName/:rowIndex', async (req, res) => {
  try {
    const { sheetName, rowIndex } = req.params;
    const data = await DataModel.findOne({ sheetName });

    if (!data) {
      return res.status(404).send({ message: 'Sheet not found' });
    }

    data.data.splice(rowIndex, 1);
    await data.save();

    res.status(200).send({ message: 'Row deleted successfully' });
  } catch (error) {
    console.error(error);
    res.status(500).send({ message: 'Error deleting row', error });
  }
});

// Endpoint to create a new row in a sheet
app.post('/api/data', async (req, res) => {
  try {
    const { sheetName, newRow } = req.body;
    const data = await DataModel.findOne({ sheetName });

    if (!data) {
      return res.status(404).send({ message: 'Sheet not found' });
    }

    data.data.unshift(newRow);
    await data.save();

    res.status(200).send({ message: 'Row added successfully' });
  } catch (error) {
    console.error(error);
    res.status(500).send({ message: 'Error adding row', error });
  }
});

// Start the server
const PORT = 5000;
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
