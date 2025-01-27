       const SPREADSHEET_ID = '1OaLFbPpDBXgbcWjLsWv2v8OXvyNIpZMNhymIBK3tK4U';
       const API_KEY = 'AIzaSyCe-KL54kSlp_YyoFJOfx1jek1go24PUF8';
       const RANGE = 'Summary!A1:Z';
       const CLIENT_ID = '692240960089-tlq9d03incdkf0rco0uav7ac4udbvj94.apps.googleusercontent.com';

        // Home page navigation
        document.getElementById('adminBtn').addEventListener('click', () => {
            document.getElementById('home-container').style.display = 'none';
            document.getElementById('admin-content').style.display = 'block';
        });

        document.getElementById('guestBtn').addEventListener('click', () => {
            document.getElementById('home-container').style.display = 'none';
            document.getElementById('guest-content').style.display = 'block';
            fetchGuestSheetData();
        });

        document.getElementById('guestBackBtn').addEventListener('click', () => {
            document.getElementById('home-container').style.display = 'block';
            document.getElementById('guest-content').style.display = 'none';
        });

        document.getElementById('adminBackBtn').addEventListener('click', () => {
            document.getElementById('home-container').style.display = 'block';
            document.getElementById('admin-content').style.display = 'none';
            handleSignoutClick();
            // document.getElementById('signout_button').style.display = 'none'
        });

        // Guest View Data Fetching
        async function fetchGuestSheetData() {
            try {
                const response = await fetch(
                    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}?key=${API_KEY}`
                );
                
                if (!response.ok) {
                    throw new Error('Network response was not ok');
                }

                const data = await response.json();
                displayGuestData(data.values);
            } catch (error) {
                console.error('Error fetching data:', error);
                document.getElementById('guestSheetData').innerHTML = 
                    '<tr><td>Error loading data. Please try again.</td></tr>';
            }
        }

        function displayGuestData(values) {
            if (!values || values.length === 0) {
                document.getElementById('guestSheetData').innerHTML = 
                    '<tr><td>No data found.</td></tr>';
                return;
            }

            const table = document.getElementById('guestSheetData');
            table.innerHTML = '';

            // Create header row
            const headerRow = document.createElement('tr');
            values[0].forEach(header => {
                const th = document.createElement('th');
                th.textContent = header;
                headerRow.appendChild(th);
            });
            table.appendChild(headerRow);

            // Create data rows
            for (let i = 1; i < values.length; i++) {
                const row = document.createElement('tr');
                values[i].forEach(cellValue => {
                    const td = document.createElement('td');
                    td.textContent = cellValue || '';
                    row.appendChild(td);
                });
                table.appendChild(row);
            }
        }


        
        // const API_KEY = 'AIzaSyCe-KL54kSlp_YyoFJOfx1jek1go24PUF8';
        // const SPREADSHEET_ID = '1fDRj8NuZWFV16h8RV5gFotthRyRPdf4QBNo-AiVFCRw';
        // const RANGE = 'Summary!A1:Z';
        const DISCOVERY_DOC = 'https://sheets.googleapis.com/$discovery/rest?version=v4';
        const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

        let tokenClient;
        let gapiInited = false;
        let gisInited = false;
        let headers = [];
        let sheetId = null;

        // ... (keep existing initialization functions the same) ...
        document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('authorize_button').addEventListener('click', handleAuthClick);
            document.getElementById('signout_button').addEventListener('click', handleSignoutClick);

            // Add download button to guest view
            const guestDownloadButton = document.createElement('button');
            guestDownloadButton.textContent = 'Download Data';
            guestDownloadButton.className = 'download-button';
            guestDownloadButton.onclick = () => {
                const guestData = Array.from(
                    document.getElementById('guestSheetData').querySelectorAll('tr')
                ).map(row => 
                    Array.from(row.querySelectorAll('th, td')).map(cell => cell.textContent)
                );
                downloadData(guestData);
            };
            document.getElementById('guest-content').insertBefore(
                guestDownloadButton, 
                document.getElementById('guestSheetData')
            );

            // Add download button to admin view
            const adminDownloadButton = document.createElement('button');
            adminDownloadButton.textContent = 'Download Data';
            adminDownloadButton.className = 'download-button';
            adminDownloadButton.onclick = () => {
                if (filteredData) {
                    downloadData(filteredData);
                } else {
                    alert('No data available to download');
                }
            };
            document.getElementById('filter-container').appendChild(adminDownloadButton);
        });


        function downloadData(data) {
            // Convert data to worksheet
            const worksheet = XLSX.utils.aoa_to_sheet(data);
            
            // Create workbook
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet Data");
            
            // Generate file name with current date
            const date = new Date();
            const fileName = `sheet_data_${date.toISOString().split('T')[0]}.xlsx`;
            
            // Trigger download
            XLSX.writeFile(workbook, fileName);
        }
        

        function loadGapiClient() {
            gapi.load('client', initializeGapiClient);
        }

        async function initializeGapiClient() {
            try {
                await gapi.client.init({
                    apiKey: API_KEY,
                    discoveryDocs: [DISCOVERY_DOC],
                });
                gapiInited = true;
                maybeEnableButtons();
            } catch (err) {
                console.error(err);
                alert('Error initializing GAPI client');
            }
        }

        function gisLoaded() {
            tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: CLIENT_ID,
                scope: SCOPES,
                callback: '', // defined later
            });
            gisInited = true;
            maybeEnableButtons();
        }

        function maybeEnableButtons() {
            if (gapiInited && gisInited) {
                document.getElementById('authorize_button').style.display = 'inline-block';
            }
        }

        function handleAuthClick() {
            tokenClient.callback = async (resp) => {
                if (resp.error !== undefined) {
                    throw (resp);
                }
                document.getElementById('signout_button').style.display = 'inline-block';
                document.getElementById('authorize_button').style.display = 'none';
                await getSheetId();
                await loadSheetData();
            };

            if (gapi.client.getToken() === null) {
                tokenClient.requestAccessToken({prompt: 'consent'});
            } else {
                tokenClient.requestAccessToken({prompt: ''});
            }
        }

        async function getSheetId() {
            try {
                const response = await gapi.client.sheets.spreadsheets.get({
                    spreadsheetId: SPREADSHEET_ID
                });
                
                const sheet = response.result.sheets.find(s => s.properties.title === 'Summary');
                if (sheet) {
                    sheetId = sheet.properties.sheetId;
                } else {
                    throw new Error('Summary sheet not found');
                }
            } catch (err) {
                console.error('Error getting sheet ID:', err);
                alert('Error getting sheet ID');
            }
        }

        async function loadSheetData() {
            try {
                const response = await gapi.client.sheets.spreadsheets.values.get({
                    spreadsheetId: SPREADSHEET_ID,
                    range: RANGE,
                });

                const range = response.result;
                if (!range || !range.values || range.values.length === 0) {
                    document.getElementById('content').innerHTML = 'No data found.';
                    return;
                }

                headers = range.values[0];
                originalData = range.values;
                filteredData = range.values;
                
                displayData(range.values);
                setupAddForm();
                setupFilterForm();
                document.getElementById('addForm').style.display = 'block';
                document.getElementById('filter-container').style.display = 'block';
                document.getElementById('add-column-container').style.display = 'block';
            } catch (err) {
                console.error(err);
                alert('Error loading sheet data');
            }
        }

        async function addNewColumn() {
            const newColumnName = document.getElementById('newColumnName').value;
            const defaultValue = document.getElementById('newColumnDefaultValue').value;

            if (!newColumnName) {
                alert('Please enter a column name');
                return;
            }

            try {
                // Add column to headers
                await gapi.client.sheets.spreadsheets.batchUpdate({
                    spreadsheetId: SPREADSHEET_ID,
                    resource: {
                        requests: [{
                            appendDimension: {
                                sheetId: sheetId,
                                dimension: 'COLUMNS',
                                length: 1
                            }
                        }]
                    }
                });

                // Update header row
                await gapi.client.sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: `Summary!${String.fromCharCode(65 + headers.length)}1`,
                    valueInputOption: 'RAW',
                    resource: {
                        values: [[newColumnName]]
                    }
                });

                // If default value is provided, fill the column
                if (defaultValue) {
                    await gapi.client.sheets.spreadsheets.values.update({
                        spreadsheetId: SPREADSHEET_ID,
                        range: `Summary!${String.fromCharCode(65 + headers.length)}2:${String.fromCharCode(65 + headers.length)}`,
                        valueInputOption: 'RAW',
                        resource: {
                            values: Array(originalData.length - 1).fill([defaultValue])
                        }
                    });
                }

                // Reload sheet data
                await loadSheetData();

                // Clear inputs
                document.getElementById('newColumnName').value = '';
                document.getElementById('newColumnDefaultValue').value = '';
            } catch (err) {
                console.error(err);
                alert('Error adding column');
            }
        }


        function setupFilterForm() {
            const filterFields = document.getElementById('filterFields');
            filterFields.innerHTML = ''; // Clear existing filter fields

            // Add first filter row
            addFilterRow(filterFields);

            // Add Filter Button Event
            document.getElementById('addFilterButton').onclick = () => addFilterRow(filterFields);
            
            // Apply Filter Button Event
            document.getElementById('filterButton').onclick = applyFilters;
            
            // Clear Filter Button Event
            document.getElementById('clearFilterButton').onclick = clearFilters;
        }

        function addFilterRow(filterFields) {
            const filterRow = document.createElement('div');
            filterRow.className = 'filter-row';

            // Column Select Dropdown
            const columnSelect = document.createElement('select');
            headers.forEach((header, index) => {
                if (index > 0) { // Skip first column (actions)
                    const option = document.createElement('option');
                    option.value = index;
                    option.textContent = header;
                    columnSelect.appendChild(option);
                }
            });

            // Condition Select Dropdown
            const conditionSelect = document.createElement('select');
            const conditions = [
                'Contains', 
                'Equals', 
                'Starts With', 
                'Ends With', 
                'Greater Than', 
                'Less Than'
            ];
            conditions.forEach(condition => {
                const option = document.createElement('option');
                option.value = condition;
                option.textContent = condition;
                conditionSelect.appendChild(option);
            });

            // Value Input
            const valueInput = document.createElement('input');
            valueInput.type = 'text';
            valueInput.placeholder = 'Filter Value';

            // Remove Filter Button
            const removeButton = document.createElement('button');
            removeButton.textContent = '✖';
            removeButton.className = 'button delete-button';
            removeButton.onclick = () => filterFields.removeChild(filterRow);

            filterRow.appendChild(columnSelect);
            filterRow.appendChild(conditionSelect);
            filterRow.appendChild(valueInput);
            filterRow.appendChild(removeButton);

            filterFields.appendChild(filterRow);
        }

        function compareDates(date1Str, date2Str, condition) {
    // Parse DD/MM/YYYY format
    const parseDate = (dateStr) => {
        const [day, month, year] = dateStr.split('/').map(Number);
        return new Date(year, month - 1, day);
    };

    try {
        const date1 = parseDate(date1Str);
        const date2 = parseDate(date2Str);

        switch(condition) {
            case 'Greater Than':
                return date1 > date2;
            case 'Less Than':
                return date1 < date2;
            default:
                return false;
        }
    } catch (error) {
        console.error('Invalid date format', error);
        return false;
    }
}

// Modify the applyFilters function to use compareDates for date columns
function parseComparableValue(value) {
    // Remove commas for numeric parsing
    const numericValue = parseFloat(value.replace(/,/g, ''));
    
    // Check if it's a date in DD/MM/YYYY format
    const datePattern = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    const dateMatch = value.match(datePattern);
    
    if (dateMatch) {
        // Convert to Date object for comparison
        const [, day, month, year] = dateMatch;
        return new Date(year, month - 1, day);
    }
    
    // Return numeric value if it's a valid number
    return isNaN(numericValue) ? value.toLowerCase() : numericValue;
}

function applyFilters() {
    const filterFields = document.getElementById('filterFields');
    const filterRows = filterFields.getElementsByClassName('filter-row');
    
    let result = originalData.slice(1); // Skip headers

    for (let row of filterRows) {
        const columnIndex = parseInt(row.getElementsByTagName('select')[0].value);
        const condition = row.getElementsByTagName('select')[1].value;
        const filterValue = row.getElementsByTagName('input')[0].value;

        result = result.filter(dataRow => {
            const cellValue = dataRow[columnIndex] || '';
            
            switch(condition) {
                case 'Contains':
                    return cellValue.toLowerCase().includes(filterValue.toLowerCase());
                case 'Equals':
                    return cellValue.toLowerCase() === filterValue.toLowerCase();
                case 'Starts With':
                    return cellValue.toLowerCase().startsWith(filterValue.toLowerCase());
                case 'Ends With':
                    return cellValue.toLowerCase().endsWith(filterValue.toLowerCase());
                case 'Greater Than':
                case 'Less Than':
                    const parsedCellValue = parseComparableValue(cellValue);
                    const parsedFilterValue = parseComparableValue(filterValue);
                    
                    // Only compare if both values are comparable (same type)
                    if (typeof parsedCellValue === typeof parsedFilterValue) {
                        if (condition === 'Greater Than') {
                            return parsedCellValue > parsedFilterValue;
                        } else {
                            return parsedCellValue < parsedFilterValue;
                        }
                    }
                    return false;
                default:
                    return true;
            }
        });
    }

    // Combine headers with filtered results
    filteredData = [originalData[0], ...result];
    displayData(filteredData);
}

        function clearFilters() {
            // Reset to original data
            filteredData = originalData;
            displayData(originalData);
            
            // Clear filter inputs
            const filterFields = document.getElementById('filterFields');
            filterFields.innerHTML = '';
            addFilterRow(filterFields);
        }

        function formatDate(dateStr) {
    if (!dateStr || typeof dateStr !== 'string' || !dateStr.includes('/')) {
        return dateStr;
    }

    const [month, day, year] = dateStr.split('/').map(Number);
    return `${day.toString().padStart(2, '0')}/${month.toString().padStart(2, '0')}/${year}`;
}

function displayData(values) {
    const table = document.getElementById('sheetData');
    table.innerHTML = '';

    // Create header row
    const headerRow = document.createElement('tr');
    
    // Add Actions header first
    const actionsTh = document.createElement('th');
    actionsTh.textContent = 'Actions';
    actionsTh.className = 'actions-cell';
    headerRow.appendChild(actionsTh);

    // Add other headers
    values[0].forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create data rows
    for (let i = 1; i < values.length; i++) {
        const row = document.createElement('tr');
        row.dataset.rowIndex = i;

        // Add action buttons first
        const actionsTd = document.createElement('td');
        actionsTd.className = 'actions-cell';
        
        const editButton = document.createElement('button');
        editButton.textContent = 'Edit';
        editButton.className = 'button edit-button';
        editButton.onclick = () => makeRowEditable(row);
        
        const deleteButton = document.createElement('button');
        deleteButton.textContent = 'Delete';
        deleteButton.className = 'button delete-button';
        deleteButton.onclick = () => deleteRow(i);

        actionsTd.appendChild(editButton);
        actionsTd.appendChild(deleteButton);
        row.appendChild(actionsTd);

        // Add data cells
        const rowData = values[i];
        for (let j = 0; j < headers.length; j++) {
            const td = document.createElement('td');
            const cellValue = j < rowData.length ? rowData[j] : '';
            
            // Check if the cell contains a date-like string
            const formattedValue = cellValue && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cellValue) 
                ? formatDate(cellValue) 
                : cellValue;

            if (formattedValue === '' || formattedValue === null || formattedValue === undefined) {
                td.textContent = '';
                td.className = 'empty-cell';
            } else {
                td.textContent = formattedValue;
            }
            row.appendChild(td);
        }

        table.appendChild(row);
    }
}
        

    

        

        function makeRowEditable(row) {
            const cells = row.getElementsByTagName('td');
            const rowData = [];

            // Skip the first cell (actions cell)
            for (let i = 1; i < cells.length; i++) {
                const cell = cells[i];
                rowData.push(cell.textContent === '' ? '' : cell.textContent);
                const input = document.createElement('input');
                input.type = 'text';
                input.value = cell.textContent === '' ? '' : cell.textContent;
                cell.textContent = '';
                cell.appendChild(input);
            }

            // Replace edit/delete buttons with save button
            const actionsTd = cells[0];
            actionsTd.innerHTML = '';
            const saveButton = document.createElement('button');
            saveButton.textContent = 'Save';
            saveButton.className = 'button save-button';
            saveButton.onclick = () => saveRow(row, rowData);
            actionsTd.appendChild(saveButton);
        }

        async function saveRow(row, oldData) {
            const inputs = row.getElementsByTagName('input');
            const newData = Array.from(inputs).map(input => input.value || ''); // Convert empty inputs to empty strings
            const rowIndex = parseInt(row.dataset.rowIndex) + 1; // +1 for header row

            try {
                await gapi.client.sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: `Summary!A${rowIndex}:${String.fromCharCode(65 + newData.length - 1)}${rowIndex}`,
                    valueInputOption: 'RAW',
                    resource: {
                        values: [newData]
                    }
                });

                await loadSheetData(); // Refresh the data
            } catch (err) {
                console.error(err);
                alert('Error saving data');
                // Revert to old data
                displayData([headers, ...oldData]);
            }
        }

        async function deleteRow(rowIndex) {
            const originalRowIndex = originalData.findIndex(row => 
                JSON.stringify(row) === JSON.stringify(filteredData[rowIndex])
            );

            if (!confirm('Are you sure you want to delete this row?')) {
                return;
            }

            if (!sheetId) {
                alert('Sheet ID not found. Please refresh and try again.');
                return;
            }

            try {
                await gapi.client.sheets.spreadsheets.batchUpdate({
                    spreadsheetId: SPREADSHEET_ID,
                    resource: {
                        requests: [{
                            deleteDimension: {
                                range: {
                                    sheetId: sheetId,
                                    dimension: 'ROWS',
                                    startIndex: originalRowIndex,
                                    endIndex: originalRowIndex + 1
                                }
                            }
                        }]
                    }
                });

                await loadSheetData(); // Refresh the data
            } catch (err) {
                console.error(err);
                alert('Error deleting row');
            }
        }

        

        function setupAddForm() {
            const formFields = document.getElementById('addFormFields');
            formFields.innerHTML = '';

            headers.forEach(header => {
                const input = document.createElement('input');
                input.type = 'text';
                input.placeholder = header;
                input.className = 'add-input';
                formFields.appendChild(input);
            });
        }

        async function addNewRow() {
            const inputs = document.getElementsByClassName('add-input');
            const newData = Array.from(inputs).map(input => input.value || ''); // Convert empty inputs to empty strings

            try {
                await gapi.client.sheets.spreadsheets.values.append({
                    spreadsheetId: SPREADSHEET_ID,
                    range: 'Summary!A1',
                    valueInputOption: 'RAW',
                    insertDataOption: 'INSERT_ROWS',
                    resource: {
                        values: [newData]
                    }
                });

                // Clear form
                Array.from(inputs).forEach(input => input.value = '');
                
                await loadSheetData(); // Refresh the data
            } catch (err) {
                console.error(err);
                alert('Error adding new row');
            }
        }

        function handleSignoutClick() {
    const token = gapi.client.getToken();
    if (token !== null) {
        // Revoke the access token
        google.accounts.oauth2.revoke(token.access_token);
        // Clear the locally stored token
        gapi.client.setToken(null);
        
        // Hide the sign-out button and show the authorize button
        document.getElementById('signout_button').style.display = 'none';
        document.getElementById('authorize_button').style.display = 'inline-block';
        
        // Clear the content
        document.getElementById('sheetData').innerHTML = '';
        document.getElementById('addForm').style.display = 'none';
        document.getElementById('filter-container').style.display = 'none';
        document.getElementById('add-column-container').style.display = 'none';
        
        // Reset global state
        headers = [];
        sheetId = null;
    }
}
