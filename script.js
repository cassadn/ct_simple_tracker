let previousSelectedEmail = 'all';

// Function to check for updates and display data with filter
function checkForUpdatesAndDisplayWithFilter(selectedEmail) {
  console.log('Checking for updates and displaying data with filter...');
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
    if (index === 0 || rowData[0] === previousSelectedEmail || previousSelectedEmail === 'all') {
      const row = document.createElement('tr');
      headers.forEach(header => {
        const cellData = rowData[headers.indexOf(header)];
        const td = document.createElement('td');
        // Handle blank cells individually
        const cellValue = cellData !== null ? parseCellData(cellData) : 'No Data'; // Placeholder for blank cells
        td.textContent = cellValue;
        row.appendChild(td);
      });
      table.appendChild(row);
    }
  });
  
  tableContainer.appendChild(table);
}

// Function to parse and convert cell data if needed
function parseCellData(cellData) {
  // Perform data conversion here if necessary
  // Example conversion for Excel date serial number to formatted date
  if (!isNaN(Date.parse(cellData))) {
    const excelSerialDate = cellData;
    const date = new Date((excelSerialDate - 25568) * 86400 * 1000);

    if (!isNaN(date.getTime())) {
      const formattedDate = ("0" + (date.getMonth() + 1)).slice(-2) + '/' + ("0" + date.getDate()).slice(-2) + '/' + date.getFullYear().toString().slice(-2);
      return formattedDate;
    }
  }

  // If no conversion needed, return the original cell data
  return cellData ;
}

// Function to populate dropdown menu with unique email values
function populateDropdown(emails, selectedEmail) {
  const dropdown = document.getElementById('filter-dropdown');
  dropdown.innerHTML = ''; // Clear previous options

  const allOption = document.createElement('option');
  allOption.value = 'all';
  allOption.textContent = 'All';
  dropdown.appendChild(allOption);

  emails.forEach(email => {
    const option = document.createElement('option');
    option.value = email;
    option.textContent = email;
    dropdown.appendChild(option);
  });

  // Set the selectedEmail in the dropdown if it was previously selected, otherwise set it to 'All'
  dropdown.value = selectedEmail !== 'all' && emails.includes(selectedEmail) ? selectedEmail : 'all';

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


