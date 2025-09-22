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
        'http://localhost:8080', 
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

// GET route to get available branches and divisions from uploaded data
app.get('/api/branches-divisions', async (req, res) => {
    try {
        const uploadsPath = path.join(__dirname, 'uploads');
        const allFiles = await fs.readdir(uploadsPath);
        
        const excelFiles = allFiles.filter(file => 
            (file.endsWith('.xlsx') || file.endsWith('.xls')) && 
            !file.startsWith('timetable_') // Exclude generated timetable files
        );
        
        console.log('Found Excel files:', excelFiles);
        
        if (excelFiles.length === 0) {
            return res.json({
                success: true,
                data: { branches: [], divisions: [], message: 'No Excel files uploaded yet' }
            });
        }
        
        const branchesSet = new Set();
        const divisionsSet = new Set();
        
        // Try to extract from filename first (DS1_Dataset.xlsx pattern)
        excelFiles.forEach(file => {
            const match = file.match(/^([A-Za-z]+)(\d+)_Dataset\.(xlsx|xls)$/);
            if (match) {
                const branch = match[1]; // e.g., "DS"
                const division = match[2]; // e.g., "1"
                branchesSet.add(branch);
                divisionsSet.add(division);
                console.log(`Extracted from filename ${file}: Branch=${branch}, Division=${division}`);
            }
        });
        
        // If no data from filenames, try reading Excel content
        if (branchesSet.size === 0 || divisionsSet.size === 0) {
            const filePath = path.join(uploadsPath, excelFiles[0]);
            const workbook = xlsx.readFile(filePath);
            
            // Check multiple sheets for branch/division data
            const sheetsToCheck = ['Theory Courses', 'Lab Courses', 'Faculty', 'Batch Details'];
            
            for (const sheetName of sheetsToCheck) {
                if (workbook.Sheets[sheetName]) {
                    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
                    console.log(`Checking sheet ${sheetName}, found ${data.length} rows`);
                    
                    data.forEach(row => {
                        const branch = row.Branch || row.branch || row.BRANCH;
                        const division = row.Division || row.division || row.DIVISION || row.Div;
                        
                        if (branch) {
                            branchesSet.add(branch.toString());
                            console.log(`Found branch in data: ${branch}`);
                        }
                        if (division) {
                            divisionsSet.add(division.toString());
                            console.log(`Found division in data: ${division}`);
                        }
                    });
                }
            }
        }
        
        const branches = Array.from(branchesSet).sort();
        const divisions = Array.from(divisionsSet).sort();
        
        console.log('Final branches:', branches);
        console.log('Final divisions:', divisions);
        
        res.json({
            success: true,
            data: {
                branches,
                divisions,
                availableFiles: excelFiles,
                debug: {
                    totalFiles: excelFiles.length,
                    branchesFound: branches.length,
                    divisionsFound: divisions.length
                }
            }
        });
        
    } catch (error) {
        console.error('Error fetching branches/divisions:', error);
        res.status(500).json({
            success: false,
            error: 'Failed to fetch available branches and divisions: ' + error.message
        });
    }
});

// GET route to retrieve stored timetables
app.get('/api/timetables', async (req, res) => {
    try {
        const { branch, division } = req.query;
        const uploadsPath = path.join(__dirname, 'uploads');
        
        console.log('=== TIMETABLES API DEBUG ===');
        console.log('Query params:', { branch, division });
        console.log('Uploads path:', uploadsPath);
        
        const allFiles = await fs.readdir(uploadsPath);
        console.log('All files in uploads:', allFiles);
        
        // Find timetable files
        let timetableFiles = allFiles.filter(file => 
            file.startsWith('timetable_') && file.endsWith('.json')
        );
        console.log('Timetable files found:', timetableFiles);
        
        // Filter by branch/division if specified
        if (branch || division) {
            const originalCount = timetableFiles.length;
            timetableFiles = timetableFiles.filter(file => {
                const matchesBranch = !branch || file.toLowerCase().includes(branch.toLowerCase());
                const matchesDivision = !division || file.includes(division.toString());
                console.log(`File ${file}: branch match=${matchesBranch}, division match=${matchesDivision}`);
                return matchesBranch && matchesDivision;
            });
            console.log(`Filtered from ${originalCount} to ${timetableFiles.length} files`);
        }
        
        const timetables = [];
        
        for (const file of timetableFiles) {
            try {
                const filePath = path.join(uploadsPath, file);
                console.log(`Reading file: ${filePath}`);
                const content = await fs.readFile(filePath, 'utf-8');
                const timetableData = JSON.parse(content);
                
                console.log(`File ${file} contains:`, {
                    branch: timetableData.branch,
                    division: timetableData.division,
                    hasRows: !!timetableData.timetable?.rows?.length
                });
                
                timetables.push({
                    filename: file,
                    ...timetableData
                });
            } catch (readError) {
                console.error(`Could not read timetable file ${file}:`, readError.message);
            }
        }
        
        // Sort by generation date (newest first)
        timetables.sort((a, b) => new Date(b.generatedAt) - new Date(a.generatedAt));
        
        console.log(`Returning ${timetables.length} timetables`);
        console.log('===========================');
        
        res.json({
            success: true,
            data: timetables
        });
        
    } catch (error) {
        console.error('Error fetching timetables:', error);
        res.status(500).json({
            success: false,
            error: 'Failed to fetch timetables: ' + error.message
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

// POST route to generate timetable for specific branch/division
app.post('/generate', async (req, res) => {
    try {
        const { branch, division, year, theoryDuration, labDuration, shortBreaks, longBreaks } = req.body;
        
        console.log('Generate request received:', { branch, division, year });
        
        const uploadsPath = path.join(__dirname, 'uploads');
        const allFiles = await fs.readdir(uploadsPath);

        let filePath = null;

        // Method 1: Look for specific branch-division file (like DS1_Dataset.xlsx)
        const specificFile = allFiles.find(file => {
            const fileName = file.toLowerCase();
            return (
                fileName.includes(branch?.toLowerCase() || '') && 
                fileName.includes(division?.toLowerCase() || '') &&
                (fileName.endsWith('.xlsx') || fileName.endsWith('.xls'))
            );
        });

        if (specificFile) {
            filePath = path.join(uploadsPath, specificFile);
            console.log(`Found specific file: ${specificFile}`);
        } else {
            // Method 2: Look for any Excel file and filter data from it
            const excelFiles = allFiles.filter(file => 
                (file.endsWith('.xlsx') || file.endsWith('.xls')) && 
                !file.startsWith('timetable_')
            );
            
            if (excelFiles.length === 0) {
                throw new Error("No Excel files found in uploads folder. Please upload data first.");
            }
            
            // Use the first Excel file and filter data from it
            filePath = path.join(uploadsPath, excelFiles[0]);
            console.log(`Using file: ${excelFiles[0]} and will filter data`);
        }

        // Read and parse the Excel file
        const workbook = xlsx.readFile(filePath);
        console.log('Available sheets:', workbook.SheetNames);

        // Helper function to safely read a sheet
        const getSheetData = (sheetName) => {
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) {
                console.log(`Sheet ${sheetName} not found`);
                return [];
            }
            return xlsx.utils.sheet_to_json(sheet);
        };

        // Extract data from sheets
        let theoryCourses = getSheetData('Theory Courses');
        let labCourses = getSheetData('Lab Courses');
        let faculty = getSheetData('Faculty');
        let loadDist = getSheetData('Load Dist');
        let batches = getSheetData('Batch Details');
        let venues = getSheetData('Venue');

        // Filter data by branch and division if they exist in the data
        if (branch && division) {
            const filterByBranchDiv = (data) => {
                return data.filter(item => {
                    const itemBranch = item.Branch || item.branch || '';
                    const itemDivision = item.Division || item.division || '';
                    
                    return (
                        itemBranch.toString().toLowerCase().includes(branch.toLowerCase()) &&
                        itemDivision.toString().toLowerCase().includes(division.toLowerCase())
                    );
                });
            };

            theoryCourses = filterByBranchDiv(theoryCourses);
            labCourses = filterByBranchDiv(labCourses);
            faculty = filterByBranchDiv(faculty);
            loadDist = filterByBranchDiv(loadDist);
            batches = filterByBranchDiv(batches);
            // Venues are usually shared, so don't filter them
        }

        console.log(`Filtered data counts:`, {
            theoryCourses: theoryCourses.length,
            labCourses: labCourses.length,
            faculty: faculty.length,
            batches: batches.length
        });

        if (theoryCourses.length === 0 && labCourses.length === 0) {
            throw new Error(`No data found for branch: ${branch}, division: ${division}. Please check your data.`);
        }

        // Prepare data for timetable generation
        const timetableData = {
            theoryCourses: theoryCourses.map(course => ({ ...course, division: division })),
            labCourses: labCourses.map(course => ({ ...course, division: division })),
            faculty: faculty.map(f => ({ ...f, division: division })),
            loadDist: loadDist.map(load => ({ ...load, division: division })),
            venues: venues,
            batches: batches.map(batch => ({ ...batch, division: division })),
            divisionCount: 1,
            branch: branch,
            division: division,
            constraints: {
                theoryDuration: theoryDuration || 60,
                labDuration: labDuration || 120,
                shortBreaks: shortBreaks || 2,
                longBreaks: longBreaks || 1
            }
        };

        // Generate the timetable using Gemini AI
        const generatedTimetable = await generateTimetableWithGemini(timetableData);
        const parsedTable = parseMarkdownTable(generatedTimetable);

        // Store the generated timetable (optional - for future retrieval)
        const timetableFile = `timetable_${branch}_${division}_${Date.now()}.json`;
        const timetableData_store = {
            branch,
            division,
            year,
            generatedAt: new Date().toISOString(),
            timetable: parsedTable
        };
        
        try {
            await fs.writeFile(
                path.join(uploadsPath, timetableFile), 
                JSON.stringify(timetableData_store, null, 2)
            );
            console.log(`Timetable saved as: ${timetableFile}`);
        } catch (saveError) {
            console.log('Could not save timetable file:', saveError.message);
        }

        // Return JSON response instead of rendering (for API usage)
        res.json({
            success: true,
            data: {
                branch,
                division,
                year,
                timetable: parsedTable,
                generatedAt: new Date().toISOString()
            }
        });

    } catch (error) {
        console.error('Timetable generation error:', error);
        res.status(500).json({
            success: false,
            error: error.message || 'An unknown error occurred during timetable generation.'
        });
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
        You are an expert university scheduler creating a timetable for ${allData.branch} branch, Division ${allData.division}.
        Your task is to create a 5-day (Monday to Friday) school week timetable.
        The time slots are: 9:00-10:00, 10:00-11:00, 11:00-12:00, 12:00-1:00, (1:00-2:00 LUNCH), 2:00-3:00, 3:00-4:00, 4:00-5:00.

        Here is the data for this specific branch and division:
        ${dataString}

        You must follow these constraints strictly:
        1. **Faculty Collision:** A faculty member CANNOT be scheduled in two different places at the same time.
        2. **Venue Collision:** A venue CANNOT be used for two different classes at the same time.
        3. **Student Collision:** A student batch cannot have a schedule conflict.
        4. **Workload Fulfillment:** The total scheduled hours for each course must match the 'Total' hours from the 'Weekly Load Distribution' data.
        5. **Lunch Break:** 1:00 PM to 2:00 PM is always the LUNCH break for everyone.

        Provide the final timetable in a single, clean Markdown table.
        The table MUST have columns: 'Day', 'Time', 'Class/Batch', 'Course Name', 'Faculty', 'Venue'.
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
