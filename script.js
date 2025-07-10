let clusters = [];
let currentClusterIndex = 0;
let approvedClusters = [];
let rejectedClusters = [];
let originalRows = [];
let headers = [];

document.getElementById('uploadBtn').addEventListener('click', () => {
  const fileInput = document.getElementById('csvFile');
  const file = fileInput.files[0];

  if (!file) {
    alert('Please select an XLSX file.');
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const dataArray = new Uint8Array(e.target.result);
    const workbook = XLSX.read(dataArray, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (jsonData.length < 2) {
      alert('No data found in the Excel sheet.');
      return;
    }

    headers = jsonData[0];

    // Parse all rows
    originalRows = jsonData.slice(1).map(row => {
      const obj = {};
      row.forEach((value, idx) => {
        obj[headers[idx].trim()] = value?.toString().trim();
      });
      return obj;
    });

    // Filter rows not marked as 'Done'
    const activeRows = originalRows.filter(row => row.Actions?.toLowerCase() !== 'done');

    // Group rows by product_id
    const grouped = {};
    activeRows.forEach(row => {
      const id = row.product_id;
      if (!id) return;

      if (!grouped[id]) {
        grouped[id] = {
          product_id: id,
          ai_web_title: row.ai_web_title || '',
          ai_amzn_title: row.ai_amzn_title || '',
          ai_desc: row.ai_desc || '',
          items: [],
          rows: []
        };
      }

      grouped[id].items.push({
        item_id: row.item_id || '',
        variation: row.Variation || '',
        image: row.Images || ''
      });

      grouped[id].rows.push(row);
    });

    clusters = Object.values(grouped);
    currentClusterIndex = 0;

    if (clusters.length > 0) {
      document.getElementById('uploadSection').style.display = 'none';
      showCluster();
    } else {
      document.getElementById('output').textContent = 'No valid clusters found (all rows marked as Done).';
    }
  };

  reader.readAsArrayBuffer(file);
});

function showCluster() {
  const cluster = clusters[currentClusterIndex];

  let detailsHTML = `
    <strong>Product ID:</strong> ${cluster.product_id}<br><br>
    <strong>AI Web Title:</strong> ${cluster.ai_web_title}<br><br>
    <strong>AI Amazon Title:</strong> ${cluster.ai_amzn_title}<br><br>
    <strong>Description:</strong>
    <div class="clamp-text">${cluster.ai_desc}</div><br><br>
    <strong>Items:</strong><br>
  `;

  cluster.items.forEach(item => {
    detailsHTML += `
      <div class="item-block">
        <strong>Item ID:</strong> ${item.item_id}<br>
        <strong>Variation:</strong> ${item.variation}<br>
        ${item.image ? `<img src="${item.image}" alt="Image for ${item.item_id}">` : ''}
      </div>
    `;
  });

  document.getElementById('clusterDetails').innerHTML = detailsHTML;

  document.getElementById('acceptBtn').style.display = 'inline-block';
  document.getElementById('rejectBtn').style.display = 'inline-block';
  document.getElementById('closeBtn').style.display = 'inline-block';
}

document.getElementById('acceptBtn').addEventListener('click', () => {
  approvedClusters.push(clusters[currentClusterIndex]);
  markRowsAsDone(clusters[currentClusterIndex]);
  nextCluster();
});

document.getElementById('rejectBtn').addEventListener('click', () => {
  rejectedClusters.push(clusters[currentClusterIndex]);
  markRowsAsDone(clusters[currentClusterIndex]);
  nextCluster();
});

function markRowsAsDone(cluster) {
  cluster.rows.forEach(row => {
    row.Actions = 'Done';
  });
}

function nextCluster() {
  currentClusterIndex++;
  if (currentClusterIndex < clusters.length) {
    showCluster();
  } else {
    document.getElementById('output').textContent = 'All clusters processed!';
    document.getElementById('acceptBtn').style.display = 'none';
    document.getElementById('rejectBtn').style.display = 'none';
    document.getElementById('clusterDetails').innerHTML = '';
  }
}

document.getElementById('closeBtn').addEventListener('click', () => {
  // 1. Update original file with "Done"
  const updatedWorkbook = XLSX.utils.book_new();
  const updatedSheet = XLSX.utils.json_to_sheet(originalRows, { header: headers });
  XLSX.utils.book_append_sheet(updatedWorkbook, updatedSheet, 'Updated Data');
  XLSX.writeFile(updatedWorkbook, 'Updated_Input_With_Actions.xlsx');

  // 2. Generate Approved & Rejected sheets
  const flattenClusters = (clusterList) => {
    const flat = [];
    clusterList.forEach(cluster => {
      cluster.items.forEach(item => {
        flat.push({
          product_id: cluster.product_id,
          ai_web_title: cluster.ai_web_title,
          ai_amzn_title: cluster.ai_amzn_title,
          ai_desc: cluster.ai_desc,
          item_id: item.item_id,
          variation: item.variation,
          image: item.image
        });
      });
    });
    return flat;
  };

  const outputWorkbook = XLSX.utils.book_new();
  const approvedSheet = XLSX.utils.json_to_sheet(flattenClusters(approvedClusters));
  const rejectedSheet = XLSX.utils.json_to_sheet(flattenClusters(rejectedClusters));
  XLSX.utils.book_append_sheet(outputWorkbook, approvedSheet, 'Approved');
  XLSX.utils.book_append_sheet(outputWorkbook, rejectedSheet, 'Rejected');
  XLSX.writeFile(outputWorkbook, 'Approved_and_Rejected.xlsx');

  // 3. Close the browser window after short delay
  setTimeout(() => {
    window.close();
  }, 2000);
});
