import axios from 'axios';
import ExcelJS from 'exceljs';
import fs from 'fs';
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

// Fetch project ID by name
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

// Fetch customer ID by name
async function fetchCustomerIdByName(customerName) {
    try {
        const response = await axiosInstance.get('/customers');
        const customer = response.data.find(c => c.name === customerName);

        if (customer) {
            return customer.id;  // Return the customer ID
        } else {
            console.error('Customer not found.');
            return null;
        }
    } catch (error) {
        console.error('Error fetching customers:', error.message);
        return null;
    }
}

// Fetch timesheets for a given project (all users)
async function fetchTimesheetsForProject(projectId) {
    try {
        const response = await axiosInstance.get('/timesheets', {
            params: {
                project: projectId,
                user: 'all',  // Fetch timesheets for all users
            },
        });
        return response.data;  // Return the list of timesheets
    } catch (error) {
        console.error('Error fetching timesheets for project:', error.message);
        return [];
    }
}

// Fetch timesheets for a given customer (all users)
async function fetchTimesheetsForCustomer(customerId) {
    try {
        const response = await axiosInstance.get('/timesheets', {
            params: {
                customer: customerId,
                user: 'all',  // Fetch timesheets for all users
            },
        });
        return response.data;  // Return the list of timesheets
    } catch (error) {
        console.error('Error fetching timesheets for customer:', error.message);
        return [];
    }
}

// Convert timesheets data into an array suitable for Excel
function convertTimesheetsToExcelData(timesheets, projectName, customerName) {
    if (!timesheets || timesheets.length === 0) {
        console.log('No timesheet data available.');
        return [];
    }

    return timesheets.map(timesheet => ({
        'Timesheet ID': timesheet.id || 'N/A',
        'Project Name': projectName || 'N/A',
        'Customer Name': customerName || 'N/A',
        'User ID': timesheet.user || 'N/A',
        'Activity ID': timesheet.activity || 'N/A',
        'Start Time': timesheet.begin || 'N/A',
        'End Time': timesheet.end || 'N/A',
        'Duration (seconds)': timesheet.duration || 'N/A',
        'Description': timesheet.description || 'No Description',
        'Billable': timesheet.billable ? 'Yes' : 'No',
    }));
}

// Save data to Excel file using ExcelJS
async function saveToExcelFile(workbook, excelData, sheetName) {
    if (excelData.length === 0) {
        console.log('No data to save to Excel.');
        return;
    }

    const worksheet = workbook.addWorksheet(sheetName);

    // Add column headers
    worksheet.columns = [
        { header: 'Timesheet ID', key: 'timesheet_id', width: 15 },
        { header: 'Project Name', key: 'project_name', width: 25 },
        { header: 'Customer Name', key: 'customer_name', width: 25 },
        { header: 'User ID', key: 'user_id', width: 15 },
        { header: 'Activity ID', key: 'activity_id', width: 15 },
        { header: 'Start Time', key: 'start_time', width: 25 },
        { header: 'End Time', key: 'end_time', width: 25 },
        { header: 'Duration (seconds)', key: 'duration', width: 20 },
        { header: 'Description', key: 'description', width: 50 },
        { header: 'Billable', key: 'billable', width: 10 },
    ];

    // Add rows one by one
    excelData.forEach(dataRow => {
        const row = [
            dataRow['Timesheet ID'],
            dataRow['Project Name'],
            dataRow['Customer Name'],
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
}

// Main function to fetch, process and save timesheets data to Excel
async function generateTimesheetReport() {
    // Use yargs to parse arguments
    const argv = yargs(process.argv.slice(2))
        .option('project', {
            description: 'Name of the project to fetch timesheets for',
            type: 'string',
        })
        .option('customer', {
            description: 'Name of the customer to fetch timesheets for',
            type: 'string',
        })
        .argv;

    const { project, customer } = argv;

    // Check if at least one argument is provided
    if (!project && !customer) {
        console.log('Please provide either a project name or a customer name.');
        return;
    }

    // Initialize the workbook for the Excel file
    const workbook = new ExcelJS.Workbook();

    // Fetch project timesheets if project name is provided
    if (project) {
        const projectId = await fetchProjectIdByName(project);
        if (projectId) {
            const timesheets = await fetchTimesheetsForProject(projectId);
            const projectExcelData = convertTimesheetsToExcelData(timesheets, project, null);
            await saveToExcelFile(workbook, projectExcelData, 'Project Timesheets');
        }
    }

    // Fetch customer timesheets if customer name is provided
    if (customer) {
        const customerId = await fetchCustomerIdByName(customer);
        if (customerId) {
            const timesheets = await fetchTimesheetsForCustomer(customerId);
            const customerExcelData = convertTimesheetsToExcelData(timesheets, null, customer);
            await saveToExcelFile(workbook, customerExcelData, 'Customer Timesheets');
        }
    }

    // Save the workbook to the file once both sheets are added
    const fileName = 'timesheets_report.xlsx';
    await workbook.xlsx.writeFile(fileName);
    console.log(`Excel file saved as: ${fileName}`);
}

// Generate the timesheet report and save to Excel
generateTimesheetReport();
