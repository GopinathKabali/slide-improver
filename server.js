const https = require("https");
const path = require("path");
const express = require("express");
const devCerts = require("office-addin-dev-certs");

const app = express();
const port = Number(process.env.PORT || 3000);

app.get(["/", "/taskpane.html"], (_request, response) => response.sendFile(path.join(__dirname, "taskpane.html")));
app.get("/taskpane.css", (_request, response) => response.sendFile(path.join(__dirname, "taskpane.css")));
app.get("/taskpane.js", (_request, response) => response.sendFile(path.join(__dirname, "taskpane.js")));

async function start() {
  try {
    const certificateOptions = await devCerts.getHttpsServerOptions();
    https.createServer(certificateOptions, app).listen(port, () => {
      console.log(`Slide Improver is running at https://localhost:${port}`);
    });
  } catch (_error) {
    console.error("Could not start the local HTTPS server. Run npm run install-dev-cert first.");
    process.exitCode = 1;
  }
}

start();
