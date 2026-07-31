const XLSX = require('xlsx');
const path = require('path');

// Columns: EMAIL ADDRESS | TOPIC | BODY | ATTACHMENT PATH
const sampleData = [
  {
    "EMAIL ADDRESS": "john.doe@example.com, tech.lead@example.com; admin@example.com",
    "TOPIC": "Project Status Update & Monthly Report",
    "BODY": "Dear Team,\n\nPlease review the status update for this month. Let us know if you have any questions.\n\nBest regards,\nProject Manager",
    "ATTACHMENT PATH": ""
  },
  {
    "EMAIL ADDRESS": "alice.smith@example.com; manager@company.org",
    "TOPIC": "Welcome & Onboarding Instructions",
    "BODY": "Hello Alice,\n\nWelcome to the team! Please complete your onboarding steps by Friday.\n\nBest,\nHR Team",
    "ATTACHMENT PATH": "C:\\Users\\Public\\Documents\\onboarding_guide.pdf"
  },
  {
    "EMAIL ADDRESS": "support@testdomain.org, billing@testdomain.org",
    "TOPIC": "Subscription Renewal Notice",
    "BODY": "Hi Support Team,\n\nYour account subscription is scheduled for renewal next week. Thank you for choosing our service!\n\nWarm regards,\nBilling Department",
    "ATTACHMENT PATH": ""
  }
];

const worksheet = XLSX.utils.json_to_sheet(sampleData);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

const targetPath = path.join(__dirname, 'sample_contacts.xlsx');
XLSX.writeFile(workbook, targetPath);
console.log(`Saved sample Excel file: ${targetPath}`);

