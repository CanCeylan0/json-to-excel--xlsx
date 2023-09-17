const XLSX = require("xlsx");
const path = require("path");

// Sheet header names
const worksheetColumnName = [
    "First Name",
    "Email",
    "Gender"
];
const worksheetName = 'Users'; // work sheet name
const filePath = './users.xlsx';

// Sample data
const UserList = [
    {
        "fname": "Osman Can",
        "lname": "CEYLAN",
        "email": "test@gmail.com",
        "gender": "Male"
    },
    {
        "fname": "Furkan Atakan",
        "lname": "BOZKURT",
        "email": "deneme@gmail.com",
        "gender": "Male"
    }
];

const exportUsersToExcel = (UserList, workSheetColumnNames, worksheetName, filePath) => {
    const data = UserList.map(user => {
        return [user.fname, user.email, user.gender];
    });

    const workBook = XLSX.utils.book_new(); // Create a new workbook
    const WorkSheetData = [
        workSheetColumnNames,
        ...data
    ]
    const worksheet = XLSX.utils.aoa_to_sheet(WorkSheetData);
    XLSX.utils.book_append_sheet(workBook, worksheet, worksheetName);
    XLSX.writeFile(workBook, path.resolve(filePath));
    return true;
}

// Call the script
exportUsersToExcel(UserList, worksheetColumnName, worksheetName, filePath);
