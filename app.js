```javascript id="4e6q1j"
const EXCEL_FILE = "master data final.xlsx";
const SHEET_NAME = "Master Data";

async function loadExcel() {

  try {

    const response = await fetch(EXCEL_FILE);

    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array"
    });

    const worksheet = workbook.Sheets[SHEET_NAME];

    const data = XLSX.utils.sheet_to_json(worksheet, {
      defval: ""
    });

    buildDashboard(data);

  } catch(error) {

    console.error(error);

    document.getElementById("loading").innerHTML =
      "Excel loading failed";
  }
}

function buildDashboard(data) {

  try {

    const totalSales = sumColumn(
      data,
      "Total Sales Qty"
    );

    const totalRevenue = sumColumn(
      data,
      "Revenue"
    );

    const totalReturns = sumColumn(
      data,
      "Total Return Qty"
    );

    document.getElementById("salesQty").innerText =
      formatNumber(totalSales);

    document.getElementById("revenue").innerText =
      formatCurrency(totalRevenue);

    document.getElementById("returns").innerText =
      formatNumber(totalReturns);

    const returnRate =
      totalSales > 0
        ? ((totalReturns / totalSales) * 100).toFixed(1)
        : 0;

    document.getElementById("returnRate").innerText =
      `${returnRate}%`;

    buildTable(data);

    document.getElementById("loading").style.display =
      "none";

    document.getElementById("dashboard").style.display =
      "block";

  } catch(error) {

    console.error(error);

    document.getElementById("loading").innerHTML =
      "Dashboard failed";
  }
}

function buildTable(data) {

  const tbody =
    document.getElementById("stylesTableBody");

  tbody.innerHTML = "";

  data.slice(0, 100).forEach(row => {

    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${row["Brand"] || ""}</td>
      <td>${row["Style id"] || ""}</td>
      <td>${row["Article Type"] || ""}</td>
      <td>${cleanNumber(row["Total Sales Qty"])}</td>
      <td>${cleanNumber(row["Revenue"])}</td>
      <td>${cleanNumber(row["Total Return Qty"])}</td>
      <td>${row["Return %"] || ""}</td>
      <td>${row["ROS"] || ""}</td>
    `;

    tbody.appendChild(tr);
  });
}

function sumColumn(data, columnName) {

  return data.reduce((sum, row) => {

    return sum + cleanNumber(row[columnName]);

  }, 0);
}

function cleanNumber(value) {

  if (!value) return 0;

  const cleaned =
    String(value)
      .replace(/,/g, "")
      .replace(/%/g, "");

  const number = Number(cleaned);

  return isNaN(number) ? 0 : number;
}

function formatNumber(value) {

  return Number(value).toLocaleString("en-IN");
}

function formatCurrency(value) {

  return new Intl.NumberFormat("en-IN", {

    style: "currency",

    currency: "INR",

    maximumFractionDigits: 0

  }).format(value);
}

loadExcel();
```
