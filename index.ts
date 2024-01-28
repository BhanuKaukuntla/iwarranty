import express, { Request, Response } from 'express';
import multer from 'multer';
import xlsx from 'xlsx';
import { MongoClient } from 'mongodb';

const app = express();
const port = 3000;

// Set up multer for handling file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage });

// MongoDB connection string
const mongoURI = 'yourMongodbURL';
const dbName = 'yourDatabaseName';

// Create a MongoDB client
const client = new MongoClient(mongoURI);

app.post('/upload', upload.single('file'), async (req: Request, res: Response) => {
  try {
    // Check if a file was provided
    if (!req.file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    // Parse the Excel file
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert the sheet data to JSON
    const jsonData: any[] = xlsx.utils.sheet_to_json(sheet);

    // Reuse the existing MongoDB connection
    await client.connect();
    const db = client.db(dbName);
    const collection = db.collection('yourCollectionName');

    await collection.insertMany(jsonData);

    res.status(200).json({ data: jsonData });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Internal server error' });
  } finally {
    // Close the MongoDB client connection
    await client.close();
  }
});

app.get('/data', async (req: Request, res: Response) => {
  try {
    // Reuse the existing MongoDB connection
    await client.connect();

    const db = client.db(dbName);
    const collection = db.collection('yourCollectionName');

    const data = await collection.find().toArray();

    res.status(200).json({ data });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Internal server error' });
  } finally {
    // Close the MongoDB client connection
    await client.close();
  }
});

app.get('/search/:searchTerm', async (req: Request, res: Response) => {
  try {
    const searchTerm: string = req.params.searchTerm;

    if (!searchTerm) {
      return res.status(400).json({ error: 'Search term is required' });
    }

    // Reuse the existing MongoDB connection
    await client.connect();

    const db = client.db(dbName);
    const collection = db.collection('yourCollectionName');

    // Construct a dynamic $or query for all fields recursively
    const buildOrQuery = (obj: any): any => {
        const orQuery: any[] = [];
        for (const key in obj) {
          if (Object.prototype.hasOwnProperty.call(obj, key)) {
            const value = obj[key];
            if (typeof value === 'string' && value.toLowerCase().includes(searchTerm.toLowerCase())) {
              orQuery.push({ [key]: value });
            } else if (typeof value === 'object') {
              const nestedQuery = buildOrQuery(value);
              if (nestedQuery.length > 0) {
                orQuery.push({ [key]: { $or: nestedQuery } });
              }
            }
          }
        }
        return orQuery;
      };
  
      const documents = await collection.find().toArray();
  
      // Build $or query for all documents
      const orQueries = documents.map((doc) => buildOrQuery(doc));
  
      // Combine all $or queries with $or operator
      const finalQuery = { $or: [].concat(...orQueries) };
  
      const results = await collection.find(finalQuery).toArray();
  
    res.json(results);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Internal server error' });
  } finally {
    // Close the MongoDB client connection
    await client.close();
  }
});

// Start the server
client.connect().then(() => {
  app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
  });
});
