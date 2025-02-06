import axios from 'axios';
import ExcelJS from 'exceljs';

const KIMAI_API_URL = 'https://demo.kimai.org/api';  // Replace with your Kimai URL
const API_TOKEN = 'token_super';  // Replace with your API token

// Axios instance with authentication
const axiosInstance = axios.create({
  baseURL: KIMAI_API_URL,
  headers: {
    'Authorization': `Bearer ${API_TOKEN}`,
    'Content-Type': 'application/json',
  },
});

// Function to fetch timesheets from Kimai API
async function fetchTimesheets() {
  try {
    const response = await axiosInstance.get('/timesheets');
    console.log(response.data)
    return response.data;
  } catch (error) {
    console.error('Error fetching timesheets:', error.message);
    return [];
  }
}

// Function to create an Excel file per project
async function createTimesheetExcel(timesheets) {
  const workbook = new ExcelJS.Workbook();
  const projectMap = {};

  // Group timesheets by project name
  timesheets.forEach((timesheet) => {
    const projectName = timesheet.project;

    if (!projectMap[projectName]) {
      projectMap[projectName] = [];
    }

    // Push timesheet details for each project
    projectMap[projectName].push({
      user: timesheet.user ? timesheet.user.firstName + ' ' + timesheet.user.lastName : timesheet.user,
      activity: timesheet.activity ? timesheet.activity : 'N/A',
      time: timesheet.time,
      comment: timesheet.comment || '',
      start: timesheet.begin,
      end: timesheet.end,
    });
  });

  // Create a new sheet for each project
  for (const projectName in projectMap) {
    const worksheet = workbook.addWorksheet(projectName);

    worksheet.columns = [
      { header: 'User', key: 'user', width: 20 },
      { header: 'Activity', key: 'activity', width: 20 },
      { header: 'Time (in seconds)', key: 'time', width: 20 },
      { header: 'Comment', key: 'comment', width: 30 },
      { header: 'Start', key: 'start', width: 20 },
      { header: 'End', key: 'end', width: 20 },
    ];

    // Add rows of timesheet data
    projectMap[projectName].forEach((entry) => {
      worksheet.addRow(entry);
    });
  }

  // Save the workbook to a file
  await workbook.xlsx.writeFile('timesheets_by_project.xlsx');
  console.log('Excel file created successfully!');
}

// Fetch timesheets and generate Excel file
fetchTimesheets().then((timesheets) => {
  if (timesheets.length > 0) {
    createTimesheetExcel(timesheets);
  } else {
    console.log('No timesheets found.');
  }
});
