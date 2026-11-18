
/**
 * CCMS SOURCE EXPORTER
 *
 * READ-ONLY for both existing CCMS Apps Script projects.
 * Creates ONE ZIP in the account's Google Drive.
 *
 * Does not modify source projects, data or deployments.
 */

function exportCCMSSource() {

  const projects = [
    {
      name: "Original_CCMS",
      id: "1d_dO1Oh8fpWIXfJZDx8LHw4etJxSrGLJTCPdb-CpLbf0tB2VTbj7buY_"
    },
    {
      name: "Clinical_Instructor_Portal",
      id: "1HRlDo-91xilN1G1jZO854htjvBfdvlkTdcm5Djq_XPnBIDvT9uPFPS9_"
    }
  ];

  const token = ScriptApp.getOAuthToken();
  const blobs = [];
  const report = {
    exportedAt: new Date().toISOString(),
    sourceType: "LATEST_SAVED_HEAD",
    deploymentVersionsVerified: false,
    projects: []
  };

  for (const project of projects) {

    const url =
      "https://script.googleapis.com/v1/projects/" +
      encodeURIComponent(project.id) +
      "/content";

    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        Authorization: "Bearer " + token
      },
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();

    if (status !== 200) {
  const errorDetails = response.getContentText();

  console.error(
    "Export failed for " + project.name +
    " (HTTP " + status + "): " +
    errorDetails.slice(0, 3000)
  );

  throw new Error(
    "Export failed. Check the execution log " +
    "for Google's detailed error message."
  );
}

    const data = JSON.parse(response.getContentText());
    const files = data.files || [];

    if (!files.length) {
      throw new Error(
        "No source files returned for " + project.name
      );
    }

    const fileReport = [];

    for (const file of files) {

      const extension = {
        SERVER_JS: ".gs",
        HTML: ".html",
        JSON: ".json"
      }[file.type];

      if (!extension) {
        throw new Error(
          "Unsupported source file type in " +
          project.name + ": " + file.type
        );
      }

      const safeName = file.name.replace(
        /[\/\\:*?"<>|]/g, "_"
      );

      const filePath =
        project.name + "/" +
        safeName + extension;

      blobs.push(
        Utilities.newBlob(
          file.source || "",
          "text/plain",
          filePath
        )
      );

      fileReport.push({
        name: file.name,
        type: file.type
      });
    }

    report.projects.push({
      name: project.name,
      scriptId: project.id,
      fileCount: files.length,
      files: fileReport
    });

    console.log(
      project.name + ": " +
      files.length + " files retrieved."
    );
  }

  blobs.push(
    Utilities.newBlob(
      JSON.stringify(report, null, 2),
      "application/json",
      "EXPORT_MANIFEST.json"
    )
  );

  const timestamp = Utilities.formatDate(
    new Date(),
    "America/New_York",
    "yyyyMMdd_HHmmss"
  );

  const zipName =
    "CCMS_Source_Export_" +
    timestamp + ".zip";

  const zip = Utilities.zip(blobs, zipName);

  const savedFile = DriveApp.createFile(zip);

  console.log("EXPORT COMPLETE");
  console.log("ZIP: " + savedFile.getName());
  console.log("Download: " + savedFile.getUrl());

  return savedFile.getUrl();
}
