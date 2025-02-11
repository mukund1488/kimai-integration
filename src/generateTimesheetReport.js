import axios from 'axios';
import ExcelJS from 'exceljs';
import yargs from 'yargs';
import fs from 'fs';
import path from 'path';
import dotenv from 'dotenv';

// Load environment variables from .env file
dotenv.config();

// Access your variables from process.env
const KIMAI_API_URL = process.env.KIMAI_API_URL;
const API_TOKEN = process.env.API_TOKEN;

console.log('Kimai API URL:', KIMAI_API_URL);
console.log('API Token:', API_TOKEN);

// Get the directory name of the current module using import.meta.url
const __filename = new URL(import.meta.url).pathname;
const __dirname = path.dirname(__filename);

// Axios instance with authentication
const axiosInstance = axios.create({
  baseURL: KIMAI_API_URL,
  headers: {
    Authorization: `Bearer ${API_TOKEN}`,
    'Content-Type': 'application/json',
  },
});

const readCustomersFromFile = (filePath) => {
  try {
    const fileContent = fs.readFileSync(filePath, 'utf8');

    // Parse the file content (assuming a text file with one customer per line)
    const customerNames = fileContent
      .split('\n')
      .map((line) => line.trim())
      .filter((line) => line);

    return customerNames;
  } catch (error) {
    console.error('Error reading customers file:', error.message);
    return [];
  }
};

const readProjectsFromFile = (filePath) => {
  try {
    const fileContent = fs.readFileSync(filePath, 'utf8');

    // Parse the file content (assuming a text file with one customer per line)
    const projectNames = fileContent
      .split('\n')
      .map((line) => line.trim())
      .filter((line) => line);

    return projectNames;
  } catch (error) {
    console.error('Error reading project files:', error.message);
    return [];
  }
};

// Fetch all customers and return a map of customer names to IDs
async function fetchAllCustomers() {
  try {
    const response = await axiosInstance.get('/customers');
    const customerMap = {};

    // Create a map of customer names to their corresponding IDs
    response.data.forEach((customer) => {
      customerMap[customer.name] = customer.id;
    });

    return customerMap;
  } catch (error) {
    console.error('Error fetching all customers:', error.message);
    return {};
  }
}

// Fetch all customers and return a map of customer names to IDs
async function fetchAllProjects() {
  try {
    const response = await axiosInstance.get('/projects');
    const projectMap = {};

    // Create a map of customer names to their corresponding IDs
    response.data.forEach((project) => {
      projectMap[project.name] = project.id;
    });

    return projectMap;
  } catch (error) {
    console.error('Error fetching all projects:', error.message);
    return {};
  }
}

// Fetch customer ID from project ID
async function fetchCustomerIdByProjectId(projectId) {
  try {
    const response = await axiosInstance.get(`/projects/${projectId}`);
    const project = response.data;

    if (project && project.customer) {
      return project.customer; // Return the customer ID associated with the project
    } else {
      console.error('Customer ID not found for this project.');
      return null;
    }
  } catch (error) {
    console.error('Error fetching project details:', error.message);
    return null;
  }
}

// Fetch customer details by customer ID
async function fetchCustomerDetailsById(customerId) {
  try {
    const response = await axiosInstance.get(`/customers/${customerId}`);
    const customer = response.data;

    if (customer) {
      return customer.name; // Return the customer name
    } else {
      console.error('Customer not found.');
      return 'Unknown Customer';
    }
  } catch (error) {
    console.error('Error fetching customer:', error.message);
    return 'Unknown Customer';
  }
}

async function fetchTimesheetsForEntity(
  entityId,
  startDate,
  endDate,
  entityType = 'customer',
) {
  let allTimesheets = [];
  let page = 1;
  let hasMore = true;

  while (hasMore) {
    try {
      // Prepare the params based on the entity type (customer or project)
      const params = {
        user: 'all', // Fetch timesheets for all users
        begin: startDate,
        end: endDate,
        page: page,
      };
      // Add the entity-specific parameter (either `customer` or `project`)
      params[entityType] = entityId;

      const response = await axiosInstance.get('/timesheets', { params });

      // Add the fetched timesheets to the list
      allTimesheets = allTimesheets.concat(response.data);

      // If the response contains fewer records than expected, stop the loop
      hasMore = response.data.length === 50;
      page++;
    } catch (error) {
      console.error(
        `Error fetching timesheets for ${entityType}:`,
        error.message,
      );
      return [];
    }
  }

  return allTimesheets;
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

// Fetch activity details by activity ID, including the description
async function fetchActivityDetailsById(activityId) {
  try {
    const response = await axiosInstance.get(`/activities/${activityId}`);
    const activity = response.data;

    if (activity) {
      return {
        name: activity.name || 'Unknown Activity',
        description: activity.comment || 'No Description Available',
      };
    } else {
      console.error('Activity not found.');
      return {
        name: 'Unknown Activity',
        description: 'No Description Available',
      };
    }
  } catch (error) {
    console.error('Error fetching activity details:', error.message);
    return {
      name: 'Unknown Activity',
      description: 'No Description Available',
    };
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

    // Fetch activity details by activity ID (name and description)
    const { name: activityName, description: activityDescription } =
      await fetchActivityDetailsById(timesheet.activity);

    // Fetch user details by user ID
    const userDetails = await fetchUserDetailsById(timesheet.user);
    const name = `${userDetails.name}`;
    const userName = `${userDetails.userName}`;
    const durationInHours =
      timesheet.duration == null || timesheet.duration === ''
        ? 'N/A'
        : (timesheet.duration / 3600).toFixed(2);

    excelData.push({
      'Customer Name': customerName || 'N/A',
      'Project Name': projectName || 'N/A',
      'User Name': name || 'N/A',
      'User Login': userName || 'N/A',
      'Activity Name': activityName || 'N/A', // Activity Name
      'Activity Description': activityDescription || 'No Description Available',
      'Start Time': timesheet.begin || 'N/A',
      'Duration (hours)': durationInHours,
      Description: timesheet.description || 'No Description',
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
    { header: 'Customer Name', key: 'customer_name', width: 25 },
    { header: 'Project Name', key: 'project_name', width: 25 },
    { header: 'User Name', key: 'user_name', width: 25 },
    { header: 'User Login', key: 'user_id', width: 15 },
    { header: 'Activity Name', key: 'activity_name', width: 30 },
    { header: 'Activity Description', key: 'activity_description', width: 100 }, //
    { header: 'Start Time', key: 'start_time', width: 25 },
    { header: 'Duration (hours)', key: 'duration', width: 20 },
    { header: 'Description', key: 'description', width: 50 },
  ];

  // Add rows one by one
  excelData.forEach((dataRow) => {
    const row = [
      dataRow['Customer Name'],
      dataRow['Project Name'],
      dataRow['User Name'],
      dataRow['User Login'],
      dataRow['Activity Name'], // Activity Name
      dataRow['Activity Description'],
      dataRow['Start Time'],
      dataRow['Duration (hours)'],
      dataRow['Description'],
    ];
    worksheet.addRow(row);
  });
}

async function writeExcel(workbook, filePath) {
  if (workbook.worksheets.length === 0) {
    console.log(`Skipping empty sheet`);
  } else {
    await workbook.xlsx.writeFile(filePath);
    console.log(`Excel file saved as: ${filePath}`);
  }
}

function deriveStartEndDate(currentDate, start_date, end_date) {
  const firstDayOfLastMonth = new Date(
    currentDate.getFullYear(),
    currentDate.getMonth() - 1,
    1,
  );
  const lastDayOfLastMonth = new Date(
    currentDate.getFullYear(),
    currentDate.getMonth(),
    0,
  );
  // If no start date or end date provided, use the default date is last month start and end
  let startDate = start_date || firstDayOfLastMonth.toLocaleDateString('en-CA');
  let endDate = end_date || lastDayOfLastMonth.toLocaleDateString('en-CA');
  // Format start date to YYYY-MM-DDThh:mm:ss (last day of previous month, with time at 00:00:00)
  const startDateFormatted = `${startDate}T00:00:00`;
  // Format end date to YYYY-MM-DDThh:mm:ss (last day of current month, with time at 23:59:59)
  const endDateFormatted = `${endDate}T23:59:59`;
  console.log('Start Date:', startDateFormatted);
  console.log('End Date:', endDateFormatted);
  return { startDateFormatted, endDateFormatted };
}

async function generateTimeSheetReportForProjects(
  project,
  timestamp,
  reportsFolderPath,
  startDateFormatted,
  endDateFormatted,
) {
  const projectFilePath = path.resolve(
    __dirname,
    '..',
    'files',
    'project_details.txt',
  );
  const projectNames = fs.existsSync(projectFilePath)
    ? readProjectsFromFile(projectFilePath)
    : [];

  // Fetch project timesheets if project name is provided
  if (project) {
    projectNames.push(project);
  }

  if (projectNames.length > 0) {
    // Generate the full file path with timestamp
    const fileName = `timesheet_report_project_${timestamp}.xlsx`;
    const filePath = path.join(reportsFolderPath, fileName);
    console.log('Saving report to:', filePath);
    const workbook = new ExcelJS.Workbook();
    const projectMap = await fetchAllProjects(); // Ensure this is awaited and completed
    for (const projectName of projectNames) {
      // Look up customer ID from the local map
      const projectId = projectMap[projectName];
      if (projectId) {
        const timesheets = await fetchTimesheetsForEntity(
          projectId,
          startDateFormatted,
          endDateFormatted,
          'project',
        );
        const projectExcelData = await convertTimesheetsToExcelData(
          timesheets,
          `${projectName}`,
        );
        await saveToExcelFile(workbook, projectExcelData, `${projectName}`);
      } else {
        console.log(`Project not found: ${projectName}`);
      }
    }
    // Save the workbook to the file
    await writeExcel(workbook, filePath);
  }
}

async function generateTimeSheetReportForCustomers(
  customer,
  timestamp,
  reportsFolderPath,
  startDateFormatted,
  endDateFormatted,
) {
  const customersFilePath = path.resolve(
    __dirname,
    '..',
    'files',
    'customer_details.txt',
  );
  const customerNames = fs.existsSync(customersFilePath)
    ? readCustomersFromFile(customersFilePath)
    : [];
  if (customerNames.length === 0) {
    console.log(
      'No customer list found or no customers in the file. Proceeding with individual customer or project fetch.',
    );
  }
  // Fetch customer timesheets if customer name is provided
  if (customer) {
    customerNames.push(customer);
  }
  // If customer list is present, fetch timesheets for each customer
  if (customerNames.length > 0) {
    // Generate the full file path with timestamp
    const fileName = `timesheet_report_customer_${timestamp}.xlsx`;
    const filePath = path.join(reportsFolderPath, fileName);
    const workbook = new ExcelJS.Workbook();
    console.log('Saving report to:', filePath);
    const customerMap = await fetchAllCustomers(); // Ensure this is awaited and completed

    for (const customerName of customerNames) {
      // Look up customer ID from the local map
      const customerId = customerMap[customerName];
      if (customerId) {
        const timesheets = await fetchTimesheetsForEntity(
          customerId,
          startDateFormatted,
          endDateFormatted,
          'customer',
        );
        const customerExcelData = await convertTimesheetsToExcelData(
          timesheets,
          `${customerName}`,
        );
        await saveToExcelFile(workbook, customerExcelData, `${customerName}`);
      } else {
        console.log(`Customer not found: ${customerName}`);
      }
    }
    // Save the workbook to the file
    await writeExcel(workbook, filePath);
  }
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
    }).argv;

  const { project, customer, start_date, end_date } = argv;

  const currentDate = new Date();
  const { startDateFormatted, endDateFormatted } = deriveStartEndDate(
    currentDate,
    start_date,
    end_date,
  );

  const timestamp = currentDate
    .toISOString()
    .replace(/[-:.]/g, '')
    .slice(0, 15); // e.g., 20250210T120000
  const reportsFolderPath = path.resolve(__dirname, '..', 'reports');
  // Ensure the reports folder exists
  if (!fs.existsSync(reportsFolderPath)) {
    fs.mkdirSync(reportsFolderPath);
  }

  // Run both async functions concurrently and wait for both to complete
  await Promise.all([
    generateTimeSheetReportForProjects(
      project,
      timestamp,
      reportsFolderPath,
      startDateFormatted,
      endDateFormatted,
    ),
    generateTimeSheetReportForCustomers(
      customer,
      timestamp,
      reportsFolderPath,
      startDateFormatted,
      endDateFormatted,
    ),
  ]);
}

// Generate the timesheet report and save to Excel
generateTimesheetReport();
