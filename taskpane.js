Office.initialize = () => {
  console.log("Office.js initialized");
};

function validateLicense() {
  const company = document.getElementById("company").value;
  const user = document.getElementById("user").value;
  const key = document.getElementById("key").value;
  const status = document.getElementById("status");

  fetch("https://your-api-domain.com/validate-license", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ company, user, licenseKey: key })
  })
    .then(res => res.json())
    .then(data => {
      if (data.status === "valid") {
        status.innerText = "✅ License Valid!";
        Office.context.roamingSettings.set("licenseValid", true);
        Office.context.roamingSettings.saveAsync();
      } else {
        status.innerText = "❌ Invalid License";
      }
    })
    .catch(error => {
      status.innerText = "⚠️ License validation failed.";
      console.error("Validation error:", error);
    });
}

function collectData() {
  Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const summary = sheets.add("Summary");
    let i = 0;
    for (const sheet of sheets.items) {
      summary.getRange(`A${i + 1}`).values = [[sheet.name]];
      i++;
    }
    await context.sync();
  });
}
