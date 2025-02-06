import axios from 'axios';
import ExcelJS from 'exceljs';
import yargs from 'yargs';

// Kimai demo API URL and token (super admin or admin token)
const KIMAI_API_URL = 'https://demo.kimai.org/api';
const API_TOKEN = 'token_super';  // Replace with your super admin API token

// Axios instance with authentication
const axiosInstance = axios.create({
    baseURL: KIMAI_API_URL,
    headers: {
        'Authorization': `Bearer ${API_TOKEN}`,
        'Content-Type': 'application/json',
    },
});

// Fetch all projects from the API to get the project ID by name
async function fetchProjectIdByName(projectName) {
    try {
        const response = await axiosInstance.get('/projects');
        const project = response.data.find(p => p.name === projectName);

        if (project) {
            return project.id;  // Return the project ID
        } else {
            console.error('Project not found.');
            return null;
        }
    } catch (error) {
        console.error('Error fetching projects:', error.message);
        return null;
    }
}

// Fetch all timesheets for a given project (all users)
async function fetchTimesheetsForProject(projectId) {
    try {
        const response = await axiosInstance.get('/timesheets', {
            params: {
                project: projectId,
                user: 'all',  // Fetch timesheets for all users
            },
        });
        console.log('Timesheets data fetched:', response.data.length); // Log the length of data for debugging
        return response.data;  // Return the list of timesheets
    } catch (error) {
        console.error('Error fetching timesheets:', error.message);
        return [];
    }
}

// Convert timesheets data into an array suitable for Excel
function convertTimesheetsToExcelData(timesheets) {
    if (!timesheets || timesheets.length === 0) {
        console.log('No timesheet data available.');
        return [];
    }

    return timesheets.map(timesheet => ({
        'Timesheet ID': timesheet.id || 'N/A',
        'Project ID': timesheet.project || 'N/A',
        'User ID': timesheet.user || 'N/A',
        'Activity ID': timesheet.activity || 'N/A',
        'Start Time': timesheet.begin || 'N/A',
        'End Time': timesheet.end || 'N/A',
        'Duration (seconds)': timesheet.duration || 'N/A',
        'Description': timesheet.description || 'No Description',  // Correcting the default behavior
        'Billable': timesheet.billable ? 'Yes' : 'No',  // Convert boolean to 'Yes'/'No'
    }));
}

// Save data to Excel file using ExcelJS
async function saveToExcelFile(excelData, fileName) {
    if (excelData.length === 0) {
        console.log('No data to save to Excel.');
        return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Timesheets');

    // Add column headers
    worksheet.columns = [
        { header: 'Timesheet ID', key: 'timesheet_id', width: 15 },
        { header: 'Project ID', key: 'project_id', width: 15 },
        { header: 'User ID', key: 'user_id', width: 15 },
        { header: 'Activity ID', key: 'activity_id', width: 15 },
        { header: 'Start Time', key: 'start_time', width: 25 },
        { header: 'End Time', key: 'end_time', width: 25 },
        { header: 'Duration (seconds)', key: 'duration', width: 20 },
        { header: 'Description', key: 'description', width: 50 },
        { header: 'Billable', key: 'billable', width: 10 },
    ];

    // Add rows one by one
    console.log('Adding rows to Excel...');
    excelData.forEach(dataRow => {
        // Convert the row object to an array matching the column order
        const row = [
            dataRow['Timesheet ID'],
            dataRow['Project ID'],
            dataRow['User ID'],
            dataRow['Activity ID'],
            dataRow['Start Time'],
            dataRow['End Time'],
            dataRow['Duration (seconds)'],
            dataRow['Description'],
            dataRow['Billable'],
        ];
        worksheet.addRow(row);
    });

    // Write the workbook to a file
    await workbook.xlsx.writeFile(fileName);
    console.log(`Excel file saved as: ${fileName}`);
}

// Main function to fetch, process and save timesheets data to Excel
async function generateTimesheetReport() {
    console.log('Parsed arguments:', process.argv);
    const argv = yargs(process.argv.slice(2))
        .option('project', {
            alias: 'p',        // Alias for the project argument
            description: 'Name of the project to fetch timesheets for',
            type: 'string',    // Ensure the project name is expected to be a string
            demandOption: true // Mark the project argument as required
        })
        .argv;
    const project = argv.project;
    console.log('Project name:', project);
    if (!project) {
        console.log('Please provide a project name using --project=<project-name>');
        return;
    }

    // Step 1: Fetch project ID by project name
    const projectId = await fetchProjectIdByName(project);

    if (!projectId) {
        console.log('Could not find project ID for the given project name.');
        return;
    }

    // Step 2: Fetch timesheets for the project
    const timesheets = await fetchTimesheetsForProject(projectId);

    if (timesheets.length > 0) {
        // Step 3: Convert timesheets to Excel format
        const excelData = convertTimesheetsToExcelData(timesheets);

        // Step 4: Save data to Excel file
        const fileName = `${project}_timesheets.xlsx`;  // File name with project name
        await saveToExcelFile(excelData, fileName);
    } else {
        console.log('No timesheets found for this project.');
    }
}

// Generate the timesheet report and save to Excel
generateTimesheetReport();
