```javascript
const EXCEL_FILE = "master data final.xlsx";
const SHEET_NAME = "Master Data";

let dashboardData = [];

async function loadExcel() {

  try {

    const response = await fetch(EXCEL_FILE);

    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array"
    });

    const worksheet = workbook.Sheets[SHEET_NAME];

    dashboardData = XLSX.utils.sheet_to_json(worksheet, {
      defval: ""
    });

    initializeDashboard(dashboardData);

  } catch (error) {

    console.error(error);

    document.getElementById("loading").innerHTML =
      "Failed to load Excel file";
  }
}

function initializeDashboard(data) {

  try {

    buildKPIs(data);

    buildCharts(data);

    buildTable(data);

    setupSearch(data);

    document.getElementById("loading").style.display = "none";

    document.getElementById("dashboard").style.display = "block";

  } catch(error) {

    console.error(error);

    document.getElementById("loading").innerHTML =
      "Dashboard rendering failed";
  }
}

function buildKPIs(data) {

  const totalSales =
    sumColumn(data, "Total Sales Qty");

  const totalRevenue =
    sumColumn(data, "Revenue");

  const totalReturns =
    sumColumn(data, "Total Return Qty");

  const returnRate =
    totalSales > 0
      ? ((totalReturns / totalSales) * 100).toFixed(1)
      : 0;

  document.getElementById("salesQty").innerText =
    formatNumber(totalSales);

  document.getElementById("revenue").innerText =
    formatCurrency(totalRevenue);

  document.getElementById("returns").innerText =
    formatNumber(totalReturns);

  document.getElementById("returnRate").innerText =
    `${returnRate}%`;
}

function buildCharts(data) {

  const monthlyMap = {};

  data.forEach(row => {

    const month =
      String(row["Month"] || "Unknown");

    if (!monthlyMap[month]) {

      monthlyMap[month] = {
        sales: 0,
        revenue: 0
      };
    }

    monthlyMap[month].sales +=
      cleanNumber(row["Total Sales Qty"]);

    monthlyMap[month].revenue +=
      cleanNumber(row["Revenue"]);
  });

  const labels = Object.keys(monthlyMap);

  const salesValues =
    labels.map(month => monthlyMap[month].sales);

  const revenueValues =
    labels.map(month => monthlyMap[month].revenue);

  new Chart(
    document.getElementById("monthlySalesChart"),
    {
      type: "bar",

      data: {
        labels,

        datasets: [{
          label: "Sales Qty",
          data: salesValues
        }]
      },

      options: {
        responsive: true
      }
    }
  );

  new Chart(
    document.getElementById("monthlyRevenueChart"),
    {
      type: "line",

      data: {
        labels,

        datasets: [{
          label: "Revenue",
          data: revenueValues
        }]
      },

      options: {
        responsive: true
      }
    }
  );
}

function buildTable(data) {

  const tbody =
    document.getElementById("stylesTableBody");

  tbody.innerHTML = "";

  data.forEach(row => {

    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${row["Brand"] || ""}</td>

      <td>${row["Style id"] || ""}</td>

      <td>${row["Article Type"] || ""}</td>

      <td>${formatNumber(cleanNumber(row["Total Sales Qty"]))}</td>

      <td>${formatCurrency(cleanNumber(row["Revenue"]))}</td>

      <td>${formatNumber(cleanNumber(row["Total Return Qty"]))}</td>

      <td>${row["Return %"] || 0}</td>

      <td>${row["ROS"] || 0}</td>
    `;

    tbody.appendChild(tr);
  });
}

function setupSearch(data) {

  const input =
    document.getElementById("searchInput");

  input.addEventListener("input", function() {

    const value = this.value.toLowerCase();

    const filtered = data.filter(row => {

      return (

        String(row["Style id"] || "")
          .toLowerCase()
          .includes(value)

        ||

        String(row["Article Type"] || "")
          .toLowerCase()
          .includes(value)
      );
    });

    buildTable(filtered);
  });
}

function sumColumn(data, columnName) {

  return data.reduce((sum, row) => {

    return sum + cleanNumber(row[columnName]);

  }, 0);
}

function cleanNumber(value) {

  if (value === null || value === undefined || value === "") {
    return 0;
  }

  const cleaned =
    String(value)
      .replace(/,/g, "")
      .replace(/%/g, "")
      .trim();

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
