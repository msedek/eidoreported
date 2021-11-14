const express = require("express");
const http = require("http");
const app = express();
const bodyParser = require("body-parser");
const cors = require("cors");
const port = process.env.PORT || 3001;

//SUDO APT INSTALL DEFAULT-JRE

const server = http.createServer(app);
app.use(cors());

app.use(bodyParser.json({ limit: "100mb", extended: true }));

const pdf = require("./routes/pdfRoutes");
const excel = require("./routes/excelRoutes");

app.use(pdf);
app.use(excel);

server.listen(port, () => {
  console.log(`Running on port ${port}`);
});
