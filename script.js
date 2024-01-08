let previousSelectedEmail = 'all';

// Function to check for updates and display data with filter
function checkForUpdatesAndDisplayWithFilter(selectedEmail) {
  console.log('Checking for updates and displaying data with filter...', new Date());
  fetch('ct_ops_workup_tracker.xlsx')
    .then(response => response.arrayBuffer())
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const tableData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const emailColumnIndex = 0;
      const uniqueEmails = [...new Set(tableData.slice(1).map(row => row[emailColumnIndex]))];
      
      // Display all data by default if 'All' is selected or no filter selected
      if (selectedEmail === 'all' || selectedEmail === null || selectedEmail === undefined) {
        displayDataInTable(tableData, tableData[0]); // Change: Display all data initially
      } else {
        let filteredData = [];
        filteredData.push(tableData[0]); // Include headers

        tableData.forEach(row => {
          if (row[emailColumnIndex] === selectedEmail) {
            // Your existing code for filtering data based on selectedEmail
            filteredData.push(row);
          }
        });
        console.log('stop2')
        // Display filtered data for the selected email
        console.log('displayed data:', displayDataInTable(filteredData, filteredData[0]));
      }

      populateDropdown(uniqueEmails, selectedEmail); // Pass selectedEmail to maintain dropdown state
    })
    .catch(error => {
      console.error('Error reading Excel file:', error);
    });
}

// Function to display data in an HTML table
function displayDataInTable(data, headers) {
  const tableContainer = document.getElementById('table-container');
  tableContainer.innerHTML = ''; // Clear previous content

  const table = document.createElement('table');
  table.classList.add('excel-table');

  const headerRow = document.createElement('tr');

  // Create table headers
  headers.forEach(headerText => {
    if (headerText === 'Workup by') {
    return;
  } else {
    const th = document.createElement('th');
    th.textContent = headerText;
    headerRow.appendChild(th);
  });

  table.appendChild(headerRow);

  // Find the index of the 'WCS Email' column
  // const emailColumnIndex = headers.indexOf('WCS Email');

  // Filter and display data based on selectedEmail
  data.forEach((rowData, index) => {
    if (index === 0) {
      // Skip the iteration to avoid displaying the header row as regular data
      return;
    }
    if (index === 0 || rowData[0] !== undefined) {
      const row = document.createElement('tr');
      headers.forEach(header => {
        const cellData = rowData[headers.indexOf(header)];
        const td = document.createElement('td');
        // Convert 'WCS Email' and 'Additional WCS to email' column to lowercase and skip the 'Workup by' column
        if (header === 'WCS Email' && cellData !== undefined) {
          const lowerCaseEmail = cellData.toLowerCase();
          td.textContent = lowerCaseEmail;
        } else if (header === 'Workup by'){
            return;
        } else if (header === 'Additional WCS to email' && cellData !== null && cellData !== undefined) {
          const lowerCaseAdEmail = cellData.toLowerCase();
          td.textContent = lowerCaseAdEmail;
        } else {
          // For other columns, perform date conversion or display data as needed
          const cellValue = cellData !== null ? parseCellData(cellData, header) : 'No Data'; // Placeholder for blank cells
          td.textContent = cellValue;
      }
      row.appendChild(td);
    });
    table.appendChild(row);
  }
});
  
  tableContainer.appendChild(table);
}

function parseCellData(cellData, header) {
  const dateColumns = ['Submission Date', 'Procedure Date'];

  if (dateColumns.includes(header)) {
    // Convert decimal numbers to integers before date conversion
    if (!isNaN(cellData)) {
      if (Number.isInteger(cellData)) {
        // Convert whole numbers to dates
        const excelSerialDate = cellData;
        const date = new Date((excelSerialDate - 25568) * 86400 * 1000); // Adjust for Excel's date origin

        if (!isNaN(date.getTime())) {
          const formattedDate = ("0" + (date.getMonth() + 1)).slice(-2) + '/' + ("0" + date.getDate()).slice(-2) + '/' + date.getFullYear().toString().slice(-2);
          return formattedDate;
        }
      } else {
        // Handle decimal numbers here, if necessary
        // For example, you might want to round the decimal to the nearest whole number
        const roundedDate = Math.round(cellData);
        return parseCellData(roundedDate, header); // Re-run the function with the rounded integer value
      }
    }
  }

  // For non-date columns or if conversion is not required, return the original cell data
  return cellData;
}

// Function to populate dropdown menu with unique lowercase email values
function populateDropdown(emails, selectedEmail) {
  const dropdown = document.getElementById('filter-dropdown');
  dropdown.innerHTML = ''; // Clear previous options

  const allOption = document.createElement('option');
  allOption.value = 'all';
  allOption.textContent = 'All';
  dropdown.appendChild(allOption);

  const uniqueLowerCaseEmails = new Set(); // Store unique lowercase emails

  emails.forEach(email => {
    const lowerCaseEmail = email.toLowerCase(); // Convert email to lowercase
    uniqueLowerCaseEmails.add(lowerCaseEmail); // Add lowercase email to the set
  });

  // Add unique lowercase emails to the dropdown
  uniqueLowerCaseEmails.forEach(email => {
    const option = document.createElement('option');
    option.value = email;
    option.textContent = email;
    dropdown.appendChild(option);
  });

  // Set the selectedEmail in the dropdown if it was previously selected, otherwise set it to 'all'
  const lowerCaseSelectedEmail = selectedEmail !== 'all' ? selectedEmail.toLowerCase() : 'all';
  dropdown.value = lowerCaseSelectedEmail;

  dropdown.addEventListener('change', function () {
    const selectedEmail = this.value;
    checkForUpdatesAndDisplayWithFilter(selectedEmail);
  });
}

function startUpdateCheck() {
  checkForUpdatesAndDisplayWithFilter('all');
  setInterval(() => {
    checkForUpdatesAndDisplayWithFilter(document.getElementById('filter-dropdown').value);
  }, 30 * 60 * 1000);
}


