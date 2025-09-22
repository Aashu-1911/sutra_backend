// Import necessary modules using ES module syntax
import express from 'express';
import xlsx from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs/promises'; // Import the file system module
import 'dotenv/config'; // Loads .env file
import { GoogleGenerativeAI } from "@google/generative-ai";
import multer from 'multer'; // Add multer for file uploads
import cors from 'cors'; // Add this line

// --- Setup ---

// Recreate __dirname for ES modules, which is not available by default
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Initialize Express app
const app = express();
const port = 3000;

// Add CORS middleware BEFORE other middleware
app.use(cors({
    origin: [
        'http://localhost:3000', 
        'http://localhost:3001', 
        'http://localhost:5173', 
        'http://localhost:4200',
        'http://192.168.171.70:8080',  // Add your actual frontend URL
        'http://192.168.171.70:*',     // Allow any port on this IP
        /^http:\/\/192\.168\.171\.\d+:\d+$/  // Allow any IP in your network range
    ],
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true
}));

// Set up EJS as the view engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Serve static files (if any, e.g., CSS)
app.use(express.static(path.join(__dirname, 'public')));

// Middleware to parse JSON and URL-encoded data
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// --- Multer Configuration for File Upload ---

// Ensure uploads directory exists
const uploadsDir = path.join(__dirname, 'uploads');

// Create uploads directory if it doesn't exist
const ensureUploadsDir = async () => {
    try {
        await fs.access(uploadsDir);
    } catch {
        await fs.mkdir(uploadsDir, { recursive: true });
        console.log('Created uploads directory');
    }
};

// Initialize uploads directory
await ensureUploadsDir();

// Configure multer for file upload
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/'); // Files will be saved in uploads folder
    },
    filename: function (req, file, cb) {
        // Keep original filename
        cb(null, file.originalname);
    }
});

// File filter to allow CSV and Excel files
const fileFilter = (req, file, cb) => {
    const allowedExtensions = ['.xlsx', '.xls', '.csv'];
    const fileExtension = path.extname(file.originalname).toLowerCase();
    
    if (allowedExtensions.includes(fileExtension)) {
        cb(null, true);
    } else {
        cb(new Error('Only Excel files (.xlsx, .xls) and CSV files are allowed!'), false);
    }
};

const upload = multer({
    storage: storage,
    fileFilter: fileFilter,
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    }
});

// --- Helper Functions ---

// Helper function to detect data type
function detectDataType(sample) {
    const keys = Object.keys(sample).map(k => k.toLowerCase());
    
    if (keys.some(k => k.includes('teacher') || k.includes('faculty') || k.includes('instructor'))) {
        return 'faculty';
    } else if (keys.some(k => k.includes('subject') || k.includes('course'))) {
        return 'subjects';
    } else if (keys.some(k => k.includes('room') || k.includes('classroom') || k.includes('venue'))) {
        return 'rooms';
    } else {
        return 'students';
    }
}

// Helper function to parse CSV file
async function parseCSVFile(filePath) {
    const csvContent = await fs.readFile(filePath, 'utf-8');
    const lines = csvContent.trim().split('\n');
    
    if (lines.length < 2) {
        throw new Error('CSV file must have at least a header and one data row');
    }
    
    const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
    const data = [];
    
    for (let i = 1; i < lines.length; i++) {
        if (lines[i].trim()) {
            const values = lines[i].split(',').map(v => v.trim().replace(/"/g, ''));
            const row = {};
            headers.forEach((header, index) => {
                row[header] = values[index] || '';
            });
            data.push(row);
        }
    }
    
    return data;
}

// --- Routes ---

// GET route for the home page
app.get('/', (req, res) => {
    res.render('index');
});

// POST route to handle file upload
app.post('/upload', upload.single('excelFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ 
                success: false, 
                error: 'Please select a file to upload.' 
            });
        }

        const filePath = req.file.path;
        const fileName = req.file.filename;
        const fileExtension = path.extname(fileName).toLowerCase();

        // Handle different file types
        if (fileExtension === '.csv') {
            // For CSV files, verify they can be parsed
            try {
                const csvData = await parseCSVFile(filePath);
                if (csvData.length === 0) {
                    await fs.unlink(filePath);
                    return res.status(400).json({
                        success: false,
                        error: 'CSV file contains no data.'
                    });
                }
            } catch (csvError) {
                await fs.unlink(filePath);
                return res.status(400).json({
                    success: false,
                    error: `Invalid CSV file: ${csvError.message}`
                });
            }
        } else {
            // For Excel files, verify structure
            try {
                const workbook = xlsx.readFile(filePath);
                const sheetNames = workbook.SheetNames;
                
                if (sheetNames.length === 0) {
                    await fs.unlink(filePath);
                    return res.status(400).json({
                        success: false,
                        error: 'Excel file contains no sheets.'
                    });
                }
                
                // For DS*_Dataset.xlsx files, check required sheets
                if (fileName.match(/^DS\d+_Dataset\.xlsx$/)) {
                    const requiredSheets = ['Theory Courses', 'Lab Courses', 'Faculty', 'Load Dist', 'Batch Details', 'Venue'];
                    const missingSheets = requiredSheets.filter(sheet => !sheetNames.includes(sheet));
                    
                    if (missingSheets.length > 0) {
                        await fs.unlink(filePath);
                        return res.status(400).json({
                            success: false,
                            error: `Missing required sheets: ${missingSheets.join(', ')}`
                        });
                    }
                }
                
            } catch (excelError) {
                await fs.unlink(filePath);
                return res.status(400).json({
                    success: false,
                    error: 'Invalid Excel file format or corrupted file.'
                });
            }
        }

        res.json({
            success: true,
            message: `File '${fileName}' uploaded successfully to backend!`,
            fileName: fileName,
            filePath: filePath,
            fileType: fileExtension.substring(1).toUpperCase()
        });

    } catch (error) {
        console.error('Upload error:', error);
        res.status(500).json({
            success: false,
            error: error.message || 'An error occurred during file upload.'
        });
    }
});

// GET route to list uploaded files
app.get('/files', async (req, res) => {
    try {
        const files = await fs.readdir(uploadsDir);
        const dataFiles = files.filter(file => {
            const ext = path.extname(file).toLowerCase();
            return ['.xlsx', '.xls', '.csv'].includes(ext);
        });
        
        // Get file stats for additional info
        const fileDetails = await Promise.all(
            dataFiles.map(async (file) => {
                const filePath = path.join(uploadsDir, file);
                const stats = await fs.stat(filePath);
                return {
                    name: file,
                    size: stats.size,
                    uploadDate: stats.mtime,
                    type: path.extname(file).substring(1).toUpperCase()
                };
            })
        );
        
        res.json({
            success: true,
            files: dataFiles,
            fileDetails: fileDetails
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: 'Failed to retrieve file list.'
        });
    }
});

// GET route to parse and return data from a specific file
app.get('/parse/:filename', async (req, res) => {
    try {
        const filename = req.params.filename;
        const filePath = path.join(uploadsDir, filename);
        
        // Check if file exists
        try {
            await fs.access(filePath);
        } catch {
            return res.status(404).json({
                success: false,
                error: 'File not found.'
            });
        }
        
        const fileExtension = path.extname(filename).toLowerCase();
        let parsedData = {};
        
        if (fileExtension === '.csv') {
            // Parse CSV file
            const csvData = await parseCSVFile(filePath);
            if (csvData.length > 0) {
                const dataType = detectDataType(csvData[0]);
                parsedData[dataType] = csvData;
            }
        } else {
            // Parse Excel file
            const workbook = xlsx.readFile(filePath);
            const sheetNames = workbook.SheetNames;
            
            // Helper function to safely read a sheet
            const getSheetData = (sheetName) => {
                const sheet = workbook.Sheets[sheetName];
                if (!sheet) return [];
                return xlsx.utils.sheet_to_json(sheet);
            };
            
            // Parse common sheet names and map to standard format
            if (sheetNames.includes('Theory Courses')) {
                parsedData.subjects = getSheetData('Theory Courses');
            }
            if (sheetNames.includes('Lab Courses')) {
                if (!parsedData.subjects) parsedData.subjects = [];
                parsedData.subjects = [...parsedData.subjects, ...getSheetData('Lab Courses')];
            }
            if (sheetNames.includes('Faculty')) {
                parsedData.faculty = getSheetData('Faculty');
            }
            if (sheetNames.includes('Venue')) {
                parsedData.rooms = getSheetData('Venue');
            }
            if (sheetNames.includes('Batch Details')) {
                parsedData.students = getSheetData('Batch Details');
            }
            
            // If no standard sheets found, try to detect data type from first sheet
            if (Object.keys(parsedData).length === 0 && sheetNames.length > 0) {
                const firstSheetData = getSheetData(sheetNames[0]);
                if (firstSheetData.length > 0) {
                    const dataType = detectDataType(firstSheetData[0]);
                    parsedData[dataType] = firstSheetData;
                }
            }
        }
        
        res.json({
            success: true,
            data: parsedData,
            filename: filename
        });
        
    } catch (error) {
        console.error('Parse error:', error);
        res.status(500).json({
            success: false,
            error: 'Failed to parse file.'
        });
    }
});

// DELETE route to remove uploaded files
app.delete('/files/:filename', async (req, res) => {
    try {
        const filename = req.params.filename;
        const filePath = path.join(uploadsDir, filename);
        
        // Check if file exists
        try {
            await fs.access(filePath);
        } catch {
            return res.status(404).json({
                success: false,
                error: 'File not found.'
            });
        }
        
        // Delete the file
        await fs.unlink(filePath);
        
        res.json({
            success: true,
            message: `File '${filename}' deleted successfully.`
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            error: 'Failed to delete file.'
        });
    }
});

// POST route to trigger timetable generation
app.post('/generate', async (req, res) => {
    try {
        const uploadsPath = path.join(__dirname, 'uploads');
        const allFiles = await fs.readdir(uploadsPath);

        // Dynamically find all files matching the DS*_Dataset.xlsx pattern
        const divisionFiles = allFiles.filter(file => /^DS\d+_Dataset\.xlsx$/.test(file));

        if (divisionFiles.length === 0) {
            throw new Error("No division data files found in the 'uploads' folder. Please upload files named like 'DS1_Dataset.xlsx', 'DS2_Dataset.xlsx', etc.");
        }
        
        console.log(`Found ${divisionFiles.length} division files to process:`, divisionFiles);

        // Initialize master arrays to hold combined data from all divisions
        let allTheoryCourses = [], allLabCourses = [], allFaculty = [], allLoadDist = [], allBatches = [], allVenues = [];

        // Helper function to safely read a sheet and throw a clear error if it's missing.
        const getSheetData = (workbook, sheetName, division) => {
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) {
                throw new Error(`The sheet named "${sheetName}" was not found for Division ${division}. Please check the file. The available sheets are: [${workbook.SheetNames.join(", ")}]`);
            }
            return xlsx.utils.sheet_to_json(sheet);
        };

        // Loop through each found division file and aggregate its data
        for (const file of divisionFiles) {
            const filePath = path.join(uploadsPath, file);
            const divisionMatch = file.match(/^DS(\d+)_Dataset\.xlsx$/);
            const division = parseInt(divisionMatch[1]);

            console.log(`Processing Division ${division} from ${file}...`);
            const workbook = xlsx.readFile(filePath);

            // Extract and tag data with its division number
            allTheoryCourses.push(...getSheetData(workbook, 'Theory Courses', division).map(course => ({ ...course, division })));
            allLabCourses.push(...getSheetData(workbook, 'Lab Courses', division).map(course => ({ ...course, division })));
            allFaculty.push(...getSheetData(workbook, 'Faculty', division).map(f => ({ ...f, division })));
            allLoadDist.push(...getSheetData(workbook, 'Load Dist', division).map(load => ({ ...load, division })));
            allBatches.push(...getSheetData(workbook, 'Batch Details', division).map(batch => ({ ...batch, division })));
            
            // Venues are a shared resource, so we only need to read them once from the first file.
            if (allVenues.length === 0) {
                allVenues = getSheetData(workbook, 'Venue', division);
            }
        }

        // Generate the timetable using the new, fully combined data structure
        const generatedTimetable = await generateTimetableWithGemini({
            theoryCourses: allTheoryCourses,
            labCourses: allLabCourses,
            faculty: allFaculty,
            loadDist: allLoadDist,
            venues: allVenues,
            batches: allBatches,
            divisionCount: divisionFiles.length
        });
        
        // Parse the Markdown table response from Gemini
        const parsedTable = parseMarkdownTable(generatedTimetable);

        res.render('result', { timetable: parsedTable, error: null });

    } catch (error) {
        console.error('An error occurred during timetable generation:', error);
        res.render('result', { timetable: null, error: error.message || 'An unknown error occurred.' });
    }
});

// --- AI and Helper Functions ---

/**
 * Generates a combined timetable for multiple divisions using the Gemini API.
 * @param {object} allData - An object containing combined and tagged data for all divisions.
 * @returns {Promise<string>} - The generated timetable as a Markdown string.
 */
async function generateTimetableWithGemini(allData) {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
        throw new Error("GEMINI_API_KEY is not set in the .env file.");
    }

    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" });

    // The data is now pre-combined and tagged with division numbers.
    const dataString = `
        All Theory Courses (with division tags): ${JSON.stringify(allData.theoryCourses, null, 2)}
        All Lab Courses (with division tags): ${JSON.stringify(allData.labCourses, null, 2)}
        All Faculty Assignments (with division tags): ${JSON.stringify(allData.faculty, null, 2)}
        All Weekly Load Distributions (with division tags): ${JSON.stringify(allData.loadDist, null, 2)}
        All Student Batches (with division tags): ${JSON.stringify(allData.batches, null, 2)}
        Shared Available Venues: ${JSON.stringify(allData.venues, null, 2)}
    `;

    const prompt = `
        You are an expert university scheduler creating a combined, collision-free timetable for ${allData.divisionCount} separate first-year B.Tech divisions.
        Your task is to create a single, unified 5-day (Monday to Friday) school week timetable that includes classes for all divisions.
        The time slots are: 9:00-10:00, 10:00-11:00, 11:00-12:00, 12:00-1:00, (1:00-2:00 LUNCH), 2:00-3:00, 3:00-4:00, 4:00-5:00.

        Faculty and Venues are shared resources between all divisions.

        Here is the combined data for all divisions. Each item has a "division" field indicating which division it belongs to.
        ${dataString}

        You must follow these constraints strictly for the COMBINED timetable:
        1.  **Faculty Collision:** A faculty member CANNOT be scheduled in two different places (for any division) at the same time.
        2.  **Venue Collision:** A venue CANNOT be used for two different classes (from any division) at the same time.
        3.  **Student Collision:** A student batch from one division cannot have a schedule conflict.
        4.  **Workload Fulfillment:** The total scheduled hours for each course for each division must match its 'Total' hours from the 'Weekly Load Distribution' data.
        5.  **Use the "division" tag:** When scheduling a course or batch, you MUST use the "division" field from the data to correctly assign it. For example, Batches 11-14 are for division 1, and Batches 21-24 are for division 2.
        6.  **Lunch Break:** 1:00 PM to 2:00 PM is always the LUNCH break for everyone.

        Provide the final combined timetable in a single, clean Markdown table.
        The table MUST have columns: 'Day', 'Time', 'Division', 'Class/Batch', 'Course Name', 'Faculty', 'Venue'.
        The 'Division' column is the most important part of the output and MUST be populated with the correct division number for every single class. Do not omit this column.
    `;

    try {
        const result = await model.generateContent(prompt);
        const response = await result.response;
        return response.text();
    } catch (error) {
        console.error("Error calling Gemini API:", error);
        throw new Error("Failed to get a response from the AI model.");
    }
}

/**
 * Parses a Markdown table string into a structured object, ensuring all rows have a consistent number of columns.
 * @param {string} markdown - The Markdown table string.
 * @returns {{headers: string[], rows: string[][]}} - An object with headers and rows.
 */
function parseMarkdownTable(markdown) {
    if (!markdown) return { headers: [], rows: [] };

    // Filter out lines that don't look like table rows (must contain '|') or are the markdown separator '---'
    const lines = markdown.trim().split('\n').filter(line => line.includes('|') && !line.includes('---'));
    
    if (lines.length < 1) return { headers: [], rows: [] };

    // The first line is always the header. Get column names and the required number of columns.
    const headers = lines[0].split('|').slice(1, -1).map(h => h.trim()).filter(Boolean);
    const numColumns = headers.length;
    
    // Process the rest of the lines as data rows
    let rows = lines.slice(1).map(line => 
        line.split('|').slice(1, -1).map(cell => cell.trim())
    );

    // Filter out any leftover separator-like rows (e.g., a row of '-----------')
    rows = rows.filter(row => row.some(cell => !cell.startsWith('---')));

    // CRITICAL FIX: Ensure every row has the same number of columns as the header.
    // Pad any short rows with empty strings. This prevents a broken table layout.
    const consistentRows = rows.map(row => {
        const newRow = [...row];
        while (newRow.length < numColumns) {
            newRow.push(''); // Add empty cells to match header count
        }
        // Also handle rows that might be too long, though less common
        return newRow.slice(0, numColumns);
    });

    return { headers, rows: consistentRows };
}

// --- Error Handling Middleware ---
app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({
                success: false,
                error: 'File size too large. Maximum size is 10MB.'
            });
        }
    }
    
    res.status(500).json({
        success: false,
        error: error.message || 'Internal server error'
    });
});

// --- Server Start ---
app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
    console.log(`Uploads directory: ${uploadsDir}`);
});
