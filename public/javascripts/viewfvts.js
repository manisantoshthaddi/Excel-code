
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Add Favorite Button
    document.getElementById("Add").blur();
    document.getElementById("Add").onclick = AddFav;
  }
});

// Function to display table 
function displayTable(tbody, serialNo, queryTag, timestamp, acceptSpan, deleteSpan) {
  let row = tbody.insertRow();
  let cell1 = row.insertCell(0);
  let cell2 = row.insertCell(1);
  let cell3 = row.insertCell(2);
  let cell4 = row.insertCell(3);
  let cell5 = row.insertCell(4);

  // Set First Cell Value
  cell1.innerText = serialNo;

  // Set Second Cell
  cell2.appendChild(queryTag);

  // Set Third Cell Value
  cell3.innerText = timestamp;

  // Set fourth cell value
  cell4.appendChild(acceptSpan);

  // Set fifth cell value
  cell5.appendChild(deleteSpan);
}

//Function to display UI
function generateTable() {
  // Display table
  document.querySelector('.table').style.display = 'table';

  // Select the existing table element
  let tbody = document.querySelector('.table_body');

  // Set empty table body
  let newTbody = document.createElement('tbody');
  newTbody.classList.add('table_body');

  // Initialize localStorage if empty
  if (localStorage.getItem('queryList') === null) {
    localStorage.setItem('queryList', JSON.stringify([]));
    //console.log(localStorage.queryList);
  }

  // If there are no queries saved make the table display as none
  if (JSON.parse(localStorage.getItem('queryList')).length === 0)
    // Hide table columns
    document.querySelector('.table').style.display = 'none';

  // Set rows
  JSON.parse(localStorage.getItem('queryList')).forEach(([query, timestamp], seno) => {
    //console.log(query, timestamp);

    // Set Query Field
    let preTag = document.createElement("pre");
    preTag.style.marginBottom = '0px';
    //console.log(sqlFormatter.format(query, { language: 'mysql' }));

    // Format query tag
    let code = document.createElement('code');
    code.setAttribute('query-content', `${query}`);
    code.textContent = `${sqlFormatter.format(query, { language: 'mysql', "tabWidth": 3, "keywordCase": "upper" })}`;
    preTag.appendChild(code);
    //preTag.innerHTML = `<code query-content="${query}">${format(query, { language: 'mysql' })}</code>`;

    // Set Accept Button
    let acceptSpan = document.createElement('span');
    acceptSpan.setAttribute('class', 'open');
    acceptSpan.setAttribute('title', 'Select Query');
    acceptSpan.innerHTML = `<img width="16" height="16" src="/images/accept.png" alt="Infor" />`;

    // Add Event Listener for select span button
    acceptSpan.onclick = function () {
      setButtonEventListener(sqlFormatter.format(query, { language: 'mysql', "keywordCase": "upper" }).replace(/\n/g, " "));;
    };

    // Set Delete Button
    let deleteSpan = document.createElement('span');
    deleteSpan.setAttribute('class', 'close');
    deleteSpan.setAttribute('title', 'Delete Query');
    deleteSpan.innerHTML = `<img width="16" height="16" src="/images/remove.png" alt="Infor" />`;
    // Add Event Listener for delete span button
    deleteSpan.onclick = function () {
      setCloseEventListener(this);
    };

    // Update Table UI
    displayTable(newTbody, seno + 1, preTag, timestamp, acceptSpan, deleteSpan);

  });

  // Replace old tBody with new one
  //console.log(newTbody);
  tbody.replaceWith(newTbody);

}

window.addEventListener('load', () => {
  generateTable();
});

function setButtonEventListener(opt) {
  //console.log(opt);
  var messageObject = { messageType: "UserQuery", query: opt };
  var jsonMessage = JSON.stringify(messageObject);
  Office.context.ui.messageParent(jsonMessage);

}

//Set Event Listener to Remove ION API File from dropdown
function setCloseEventListener(row) {
  //console.log(row);

  // select element with queryText
  let i = row.parentNode.parentNode.rowIndex;
  //console.log(i);
  let tr = document.getElementsByTagName("tr")[i];

  let name = tr.cells[1].firstChild.firstChild.getAttribute('query-content');
  //console.log(name);
  let queryList = JSON.parse(localStorage.getItem('queryList'));

  //console.log(queryList.filter(([query]) => query !== name));
  localStorage.setItem('queryList', JSON.stringify(queryList.filter(([query]) => query !== name)));

  generateTable();
}

// Function to add favorite element
function AddFav() {
  // Blur the button
  document.getElementById("Add").blur();

  let tbody = document.querySelector('.table_body');

  // Read Query
  let query = document.getElementById('UserQuery').value.trim();

  // Original unFormatted query for message
  //let originalQuery = query;

  // format query using library to avoid duplication
  query = sqlFormatter.format(query, { language: 'mysql', "tabWidth": 3, "keywordCase": "upper" });
  //console.log(query);

  // Check if empty query is entered
  if (query === "") {
    let myModal = new bootstrap.Modal(document.getElementById("myModal"));
    document.getElementById("modalHeading").innerHTML = "Empty Query";
    document.getElementById("modalText").innerHTML = "Please enter a Query";
    myModal.show();
  }

  // Check if Query already exists;

  else if (localStorage.getItem('queryList') && JSON.parse(localStorage.queryList).map(([query]) => query).includes(query)) {
    //console.log("Same query");
    let myModal = new bootstrap.Modal(document.getElementById("myModal"));
    document.getElementById("modalHeading").innerHTML = "Duplicate Query";
    document.getElementById("modalText").innerHTML = "The entered query matches an existing query, please verify.";
    myModal.show();
  }

  else {
    // Display table columns
    document.querySelector('.table').style.display = 'table';

    // Check if queryList property exists in localstorage
    if (localStorage.getItem('queryList') === null) {
      localStorage.setItem('queryList', JSON.stringify([]));
      //console.log(localStorage.queryList);
    }

    // Check if 20 queries are already saved
    else if (JSON.parse(localStorage.getItem('queryList')).length === 20) {
      //console.log("Limit reached");
      let myModal = new bootstrap.Modal(document.getElementById("myModal"));
      document.getElementById("modalHeading").innerHTML = "Query Limit";
      document.getElementById("modalText").innerHTML = "Saved Query Limit is 20, please verify.";
      myModal.show();
    }

    // Valid query. Add to localstorage
    else {
      // Add query and timestamp to queryList
      let timestamp = new Date().toLocaleString();

      // Set queryList to localStorage
      let queryList = JSON.parse(localStorage.getItem('queryList'));
      //console.log(queryList);
      queryList.push([query, timestamp]);
      localStorage.setItem('queryList', JSON.stringify(queryList));

      // Make query textbox as null
      document.getElementById('UserQuery').value = "";

      // Set Query Field
      let preTag = document.createElement("pre");
      preTag.style.marginBottom = '0px';
      preTag.classList.add('queryOption');
      //console.log(sqlFormatter.format(query, { language: 'mysql' }));

      // Format query tag
      let code = document.createElement('code');
      code.setAttribute('query-content', `${query}`);
      code.textContent = `${sqlFormatter.format(query, { language: 'mysql', "tabWidth": 3, "keywordCase": "upper" })}`;
      preTag.appendChild(code);
      //preTag.innerHTML = `<code query-content="${query}">${format(query, { language: 'mysql' })}</code>`;

      // Set Accept Button
      let acceptSpan = document.createElement('span');
      acceptSpan.setAttribute('class', 'open');
      acceptSpan.setAttribute('title', 'Select Query');
      acceptSpan.innerHTML = `<img width="16" height="16" src="/images/accept.png" alt="Infor" />`;

      // Add Event Listener for select span button
      acceptSpan.onclick = function () {
        setButtonEventListener(sqlFormatter.format(query, { language: 'mysql', "keywordCase": "upper" }).replace(/\n/g, " "));
      };

      // Set Delete Button
      let deleteSpan = document.createElement('span');
      deleteSpan.setAttribute('class', 'close');
      deleteSpan.setAttribute('title', 'Delete Query');
      deleteSpan.innerHTML = `<img width="16" height="16" src="/images/remove.png" alt="Infor" />`;
      // Add Event Listener for delete span button
      deleteSpan.onclick = function () {
        setCloseEventListener(this);
      };

      // Update Table UI
      displayTable(tbody, queryList.length, preTag, timestamp, acceptSpan, deleteSpan);
    }

  }
}