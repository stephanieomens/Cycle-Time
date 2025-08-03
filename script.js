let cycleChart;

document.getElementById("excelUpload").addEventListener("change", handleExcelUpload);

function handleExcelUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    let jsonData = XLSX.utils.sheet_to_json(sheet);

    // Calculate durations
    jsonData = jsonData.map(row => {
      const parse = str => (str ? new Date(str) : null);
      const daysBetween = (start, end) => {
        if (!start || !end) return null;
        const diff = (end - start) / (1000 * 60 * 60 * 24);
        return Math.round(diff);
      };

      const disb = parse(row.DisbursementDate);
      const grn = parse(row.GRNDate);
      const dispatch = parse(row.DispatchDate);
      const port = parse(row.PortArrivalDate);
      const offload = parse(row.OffloadDate);

      return {
        ...row,
        FundDurationDays: daysBetween(disb, grn),
        GRNToDispatchDays: daysBetween(grn, dispatch),
        DispatchToPortDays: daysBetween(dispatch, port),
        PortToOffloadDays: daysBetween(port, offload),
      };
    });

    renderMetrics(jsonData);
    renderChart(jsonData);
    enableReupload();
  };
  reader.readAsArrayBuffer(file);
}

function renderMetrics(data) {
  const sum = (arr, key) =>
    arr.reduce((acc, row) => {
      const val = parseFloat(row[key]);
      return acc + (isNaN(val) ? 0 : val);
    }, 0);

  const avg = (arr, key) => {
    const validValues = arr.map(row => parseFloat(row[key])).filter(val => !isNaN(val));
    return validValues.length ? (sum(arr, key) / validValues.length).toFixed(1) : "--";
  };

  document.getElementById("fundDuration").textContent = `Fund Duration: ${avg(data, "FundDurationDays")} days`;
  document.getElementById("grnToDispatch").textContent = `GRN to Dispatch: ${avg(data, "GRNToDispatchDays")} days`;
  document.getElementById("dispatchToPort").textContent = `Dispatch to Port: ${avg(data, "DispatchToPortDays")} days`;
  document.getElementById("portToOffload").textContent = `Port to Offload: ${avg(data, "PortToOffloadDays")} days`;
}

function renderChart(data) {
  const buyers = data.map(row => row.Buyer || "Unknown");
  const fundDur = data.map(row => parseFloat(row.FundDurationDays) || 0);
  const grn = data.map(row => parseFloat(row.GRNToDispatchDays) || 0);
  const port = data.map(row => parseFloat(row.DispatchToPortDays) || 0);
  const offload = data.map(row => parseFloat(row.PortToOffloadDays) || 0);

  if (cycleChart) cycleChart.destroy();

  cycleChart = new Chart(document.getElementById("cycleChart"), {
    type: "bar",
    data: {
      labels: buyers,
      datasets: [
        {
          label: "Fund Duration",
          backgroundColor: "#a83232",
          data: fundDur,
        },
        {
          label: "GRN to Dispatch",
          backgroundColor: "#f47c3c",
          data: grn,
        },
        {
          label: "Dispatch to Port",
          backgroundColor: "#b8923c",
          data: port,
        },
        {
          label: "Port to Offload",
          backgroundColor: "#1e4d2b",
          data: offload,
        }
      ]
    },
    options: {
      responsive: true,
      scales: {
        x: {
          stacked: true,
          title: {
            display: true,
            text: "Buyers",
            font: { weight: "bold" }
          }
        },
        y: {
          stacked: false,
          title: {
            display: true,
            text: "Days",
            font: { weight: "bold" }
          },
          beginAtZero: true
        }
      },
      plugins: {
        legend: { position: "top" },
        tooltip: {
          mode: "index",
          intersect: false,
        }
      }
    }
  });
}

function downloadExcel() {
  if (!cycleChart) {
    alert("No data available to download.");
    return;
  }

  const table = [
    ["Buyer", "FundDurationDays", "GRNToDispatchDays", "DispatchToPortDays", "PortToOffloadDays"],
    ...cycleChart.data.labels.map((buyer, i) => [
      buyer,
      cycleChart.data.datasets[0].data[i],
      cycleChart.data.datasets[1].data[i],
      cycleChart.data.datasets[2].data[i],
      cycleChart.data.datasets[3].data[i],
    ])
  ];

  const ws = XLSX.utils.aoa_to_sheet(table);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Cycle Times");

  const now = new Date();
  const filename = `Cycle_Report_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}.xlsx`;
  XLSX.writeFile(wb, filename);
}

function enableReupload() {
  if (!document.getElementById("reuploadBtn")) {
    const btn = document.createElement("button");
    btn.textContent = "Upload Another File";
    btn.id = "reuploadBtn";
    btn.style.marginTop = "10px";
    btn.className = "reupload";
    btn.onclick = () => {
      document.getElementById("excelUpload").value = "";
      document.getElementById("excelUpload").click();
    };
    document.querySelector(".upload-section").appendChild(btn);
  }
}
