async function generateUTM() {
    await Excel.run(async (context) => {
      const url = document.getElementById("urlText").value;
      const campaignId = document.getElementById("campaignId").value;
      const campaignSource = document.getElementById("campaignSourceList").value;
      const campaignMedium = document.getElementById("campaignMediumList").value;
      const campaignName = document.getElementById("campaignNameList").value;
      const campaignTerm = document.getElementById("campaignTermText").value;
  
      if (!url || !campaignId || !campaignSource || !campaignMedium || !campaignName) {
        alert("All fields must be filled out.");
        return;
      }
  
      const utm = `${url}?utm_campaign=${campaignId.replace(/ /g, "+")}&utm_source=${campaignSource.replace(
        / /g,
        "+"
      )}&utm_medium=${campaignMedium.replace(/ /g, "+")}&utm_term=${campaignTerm.replace(
        / /g,
        "+"
      )}&utm_name=${campaignName.replace(/ /g, "+")}`;
  
      document.getElementById("generatedUTMText").value = utm;
  
      const sheet = context.workbook.worksheets.getItem("Data");
      const nextRow = sheet.getUsedRange().getRowCount() + 1;
      sheet.getRange(`F${nextRow}`).values = [[utm]];
      sheet.getRange(`G${nextRow}`).values = [[campaignId]];
      sheet.getRange(`H${nextRow}`).values = [[campaignName]];
      await context.sync();
    });
  }
  
  /** Default helper for invoking an action and handling errors. */
  async function tryCatch(callback) {
    try {
      await callback();
    } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    }
  }