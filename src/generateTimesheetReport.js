import axios from 'axios';
import ExcelJS from 'exceljs';
import yargs from 'yargs';
import fs from 'fs';
import path from 'path';

// Get the directory name of the current module using import.meta.url
const __filename = new URL(import.meta.url).pathname;
const __dirname = path.dirname(__filename);

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

const readCustomersFromFile = (filePath) => {
    try {
        const fileContent = fs.readFileSync(filePath, 'utf8');

        // Parse the file content (assuming a text file with one customer per line)
        const customerNames = fileContent.split('\n').map(line => line.trim()).filter(line => line);

        return customerNames;
    } catch (error) {
        console.error('Error reading customers file:', error.message);
        return [];
    }
};

// Fetch all customers and return a map of customer names to IDs
async function fetchAllCustomers() {
    try {
        const response = await axiosInstance.get('/customers');
        const customerMap = {};

        // Create a map of customer names to their corresponding IDs
        response.data.forEach(customer => {
            customerMap[customer.name] = customer.id;
        });

        return customerMap;
    } catch (error) {
        console.error('Error fetching all customers:', error.message);
        return {};
    }
}


// Fetch project ID by name
async function fetchProjectIdByName(projectName) {
    try {
        const response = await axiosInstance.get('/projects');
        const project = response.data.find(p => p.name === projectName);

        if (project) {
            return project.id;  // Return the project ID
        } else {
            console.error(`Project ${projectName} not found.`);
            return null;
        }
    } catch (error) {
        console.error('Error fetching projects:', error.message);
        return null;
    }
}

// Fetch customer ID from project ID
async function fetchCustomerIdByProjectId(projectId) {
    try {
        const response = await axiosInstance.get(`/projects/${projectId}`);
        const project = response.data;

        if (project && project.customer) {
            return project.customer;  // Return the customer ID associated with the project
        } else {
            console.error('Customer ID not found for this project.');
            return null;
        }
    } catch (error) {
        console.error('Error fetching project details:', error.message);
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
            console.error(`Customer ${customerName} not found.`);
            return null;
        }
    } catch (error) {
        console.error('Error fetching customers:', error.message);
        return null;
    }
}

// Fetch customer details by customer ID
async function fetchCustomerDetailsById(customerId) {
    try {
        const response = await axiosInstance.get(`/customers/${customerId}`);
        const customer = response.data;

        if (customer) {
            return customer.name;  // Return the customer name
        } else {
            console.error('Customer not found.');
            return 'Unknown Customer';
        }
    } catch (error) {
        console.error('Error fetching customer:', error.message);
        return 'Unknown Customer';
    }
}

// Fetch timesheets for a given project (all users) within a specific date range
async function fetchTimesheetsForProject(projectId, startDate, endDate) {
    try {
        console.log(projectId, startDate, endDate);
        const response = await axiosInstance.get('/timesheets', {
            params: {
                project: projectId,
                user: 'all',  // Fetch timesheets for all users
                begin: startDate,
                end: endDate,
            },
        });
        return response.data;  // Return the list of timesheets
    } catch (error) {
        console.error('Error fetching timesheets for project:', error.message);
        return [];
    }
}

// Fetch timesheets for a given customer (all users) within a specific date range
async function fetchTimesheetsForCustomer(customerId, startDate, endDate) {
    try {
        const response = await axiosInstance.get('/timesheets', {
            params: {
                customer: customerId,
                user: 'all',  // Fetch timesheets for all users
                begin: startDate,
                end: endDate,
            },
        });
        return response.data;  // Return the list of timesheets
    } catch (error) {
        console.error('Error fetching timesheets for customer:', error.message);
        return [];
    }
}

// Fetch user details by user ID
async function fetchUserDetailsById(userId) {
    try {
        const response = await axiosInstance.get(`/users/${userId}`);
        const user = response.data;

        if (user) {
            return {
                name: user.alias || 'N/A',
                userName: user.username || 'N/A',
            };
        } else {
            console.error('User not found.');
            return { firstName: 'Unknown', lastName: 'User' };
        }
    } catch (error) {
        console.error('Error fetching user details:', error.message);
        return { firstName: 'Unknown', lastName: 'User' };
    }
}

// Convert timesheets data into an array suitable for Excel
async function convertTimesheetsToExcelData(timesheets, sheetName) {
    if (!timesheets || timesheets.length === 0) {
        console.log(`No timesheet data available for: ${sheetName}`);
        return [];
    }

    const excelData = [];

    for (const timesheet of timesheets) {
        // Fetch project name by project ID
        const projectId = timesheet.project;
        const projectResponse = await axiosInstance.get(`/projects/${projectId}`);
        const projectName = projectResponse.data.name || 'Unknown Project';

        // Fetch customer ID from the project
        const customerId = await fetchCustomerIdByProjectId(projectId);

        // Fetch customer name using the customer ID
        const customerName = await fetchCustomerDetailsById(customerId);

        // Fetch user details by user ID
        const userDetails = await fetchUserDetailsById(timesheet.user);
        const name = `${userDetails.name}`;
        const userName = `${userDetails.userName}`;

        excelData.push({
            'Project Name': projectName || 'N/A',
            'Customer Name': customerName || 'N/A',
            'User Name': name || 'N/A',
            'User Login': userName || 'N/A',
            'Activity ID': timesheet.activity || 'N/A',
            'Start Time': timesheet.begin || 'N/A',
            'End Time': timesheet.end || 'N/A',
            'Duration (seconds)': timesheet.duration || 'N/A',
            'Description': timesheet.description || 'No Description',
            'Billable': timesheet.billable ? 'Yes' : 'No',
        });
    }

    return excelData;
}

// Save data to Excel file using ExcelJS
async function saveToExcelFile(workbook, excelData, sheetName) {
    if (excelData.length === 0) {
        console.log(`No data to save to Excel for sheet: ${sheetName}`);
        return;
    }

    const worksheet = workbook.addWorksheet(sheetName);

    // Add column headers
    worksheet.columns = [
        { header: 'Project Name', key: 'project_name', width: 25 },
        { header: 'Customer Name', key: 'customer_name', width: 25 },
        { header: 'User Name', key: 'user_name', width: 25 },
        { header: 'User Login', key: 'user_id', width: 15 },
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
            dataRow['Project Name'],
            dataRow['Customer Name'],
            dataRow['User Name'],
            dataRow['User Login'],
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
        .option('start_date', {
            description: 'Start date for the time period (YYYY-MM-DD)',
            type: 'string',
        })
        .option('end_date', {
            description: 'End date for the time period (YYYY-MM-DD)',
            type: 'string',
        })
        .argv;

    const { project, customer, start_date, end_date } = argv;

    // Get current date and timestamp for filename
    const currentDate = new Date();
    const timestamp = currentDate.toISOString().replace(/[-:.]/g, '').slice(0, 15); // e.g., 20250210T120000
    const reportsFolderPath = path.resolve(__dirname, '..', 'reports');

    const firstDayOfLastMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
    const lastDayOfLastMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 0);
    // If no start date or end date provided, use the default date is last month start and end
    let startDate = start_date || firstDayOfLastMonth.toLocaleDateString('en-CA');
    let endDate = end_date || lastDayOfLastMonth.toLocaleDateString('en-CA');
    // Format start date to YYYY-MM-DDThh:mm:ss (last day of previous month, with time at 00:00:00)
    const startDateFormatted = `${startDate}T00:00:00`;
    // Format end date to YYYY-MM-DDThh:mm:ss (last day of current month, with time at 23:59:59)
    const endDateFormatted = `${endDate}T23:59:59`;
    // Log the start and end dates
    console.log('Current Date:', currentDate);
    console.log('Start Date:', startDateFormatted);
    console.log('End Date:', endDateFormatted);


    // Ensure the reports folder exists
    if (!fs.existsSync(reportsFolderPath)) {
        fs.mkdirSync(reportsFolderPath);
    }

    // Generate the full file path with timestamp
    const fileName = `timesheet_report_${timestamp}.xlsx`;
    const filePath = path.join(reportsFolderPath, fileName);

    console.log('Saving report to:', filePath);

    // Initialize the workbook for the Excel file
    const workbook = new ExcelJS.Workbook();

    // Fetch project timesheets if project name is provided
    if (project) {
        const projectId = await fetchProjectIdByName(project);
        if (projectId) {
            const timesheets = await fetchTimesheetsForProject(projectId, startDateFormatted, endDateFormatted);
            const projectExcelData = await convertTimesheetsToExcelData(timesheets, 'Project Timesheets');
            await saveToExcelFile(workbook, projectExcelData, 'Project Timesheets');
        }
    }

    const customersFilePath = path.resolve(__dirname, '..', 'files', 'customer_details.txt');
    const customerNames = fs.existsSync(customersFilePath) ? readCustomersFromFile(customersFilePath) : [];
    if (customerNames.length === 0) {
        console.log('No customer list found or no customers in the file. Proceeding with individual customer or project fetch.');
    }
    // Fetch customer timesheets if customer name is provided
    if (customer) {
        customerNames.push(customer)
        /*const customerId = await fetchCustomerIdByName(customer);
        if (customerId) {
            const timesheets = await fetchTimesheetsForCustomer(customerId, start_date, end_date);
            const customerExcelData = await convertTimesheetsToExcelData(timesheets, 'Customer Timesheets');
            await saveToExcelFile(workbook, customerExcelData, 'Customer Timesheets');
        }*/
    }
    // If customer list is present, fetch timesheets for each customer
    if (customerNames.length > 0) {
        const customerMap = await fetchAllCustomers(); // Ensure this is awaited and completed

        for (const customerName of customerNames) {
            // Look up customer ID from the local map
            const customerId = customerMap[customerName];
            if (customerId) {
                const timesheets = await fetchTimesheetsForCustomer(customerId, startDateFormatted, endDateFormatted);
                const customerExcelData = await convertTimesheetsToExcelData(timesheets, `Timesheets for ${customerName}`);
                await saveToExcelFile(workbook, customerExcelData, `Timesheets for ${customerName}`);
            } else {
                console.log(`Customer not found: ${customerName}`);
            }
        }
    }

    // Save the workbook to the file
    if(workbook.worksheets.length === 0) {
        console.log(`No data to save to Excel`);
    }
    else
    {
        await workbook.xlsx.writeFile(filePath);
        console.log(`Excel file saved as: ${filePath}`);
    }

}

// Generate the timesheet report and save to Excel
generateTimesheetReport();
