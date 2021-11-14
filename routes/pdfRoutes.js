const express = require("express");
const router = express.Router();
const Excel = require("exceljs");
const converter = require("office-converter")();
const fs = require("fs");
const path = require("path");
const _ = require("underscore");
const merge = require("easy-pdf-merge");
const nodemailer = require("nodemailer");
// const PDFDocument = require('pdfkit')
// const xoauth2 = require("xoauth2");
const smtpTransport = require("nodemailer-smtp-transport");

const configs = require("../configs/configs");
const MAIL = configs.mailing;

const SERVICE = configs.service;
const HOST = configs.host;
const USER = configs.user;
const PASS = configs.pass;
const FROM = configs.from;

const maillist = MAIL;

const clean = fileNames => {
  return new Promise((resolve, reject) => {
    fs.readdir("./reports/", (err, files) => {
      if (err) throw err;
      for (const file of files) {
        for (let i = 1; i < fileNames.length; i++) {
          const myFile = fileNames[i];
          if (file === `${myFile}.pdf` || file === `${myFile}.xlsx`) {
            fs.unlink(path.join("./reports/", file), err => {
              if (err) throw err;
            });
          }
        }
      }
      resolve("End");
    });
  });
};

const mail = fileNames => {
  return new Promise((resolve, reject) => {
    let matrix = fileNames[0].split(".");
    let thePDF = `${matrix[0]}.pdf`;

    const transporter = nodemailer.createTransport(
      smtpTransport({
        service: SERVICE,
        host: HOST,
        auth: {
          user: USER,
          pass: PASS
        }
      })
    );

    // const transporter = nodemailer.createTransport({
    //     service: 'gmail',
    //     auth: {
    //         xoauth2: xoauth2.createXOAuth2Generator({
    //             user: 'miguelsedek@gmail.com',
    //             clientId: '539656676705-gotkj2cmlrin0ie4emgfa0datkmu4ptg.apps.googleusercontent.com',
    //             clientSecret: '9roMtgxg4E9RcCrNXoa02Q9A',
    //             refreshToken: '1/RymuZBeCVdyw5YyEag7TFc3gKBTH6Ie3TLPzJ-0LTxUvsdxIV1y0_XM1uV_qz4yX'
    //         })
    //     }
    // })

    const mailOptions = {
      from: `EIDOTAB ${FROM}`,
      to: maillist,
      subject: matrix[0],
      text: `Reporte de cierre de caja ${matrix[0]}`,
      attachments: [
        {
          // file on disk as an attachment
          filename: fileNames[0],
          path: `./reports/${fileNames[0]}` // stream this file
        },
        {
          // file on disk as an attachment
          filename: thePDF,
          path: `./reports/${thePDF}` // stream this file
        }
      ]
    };

    transporter.sendMail(mailOptions, function(err, res) {
      if (err) {
        console.log(err);
      } else {
        resolve("Email Sent");
      }
    });
  });
};

const goMerge = fileNames => {
  return new Promise((resolve, reject) => {
    let files = [];
    let result = "";
    for (let i = 0; i < fileNames.length; i++) {
      let file = `${fileNames[i]}`;
      if (i > 0) {
        file = `${file}.pdf`;
      } else {
        let matrix = file.split(".");
        file = `${matrix[0]}.pdf`;
        result = `./reports/${matrix[0]}.pdf`;
      }
      files.push(`./reports/${file}`);
    }

    merge(files, result, function(err) {
      if (err) return console.log(err);
      resolve("Success");
    });
  });
};

const generatePdf = file => {
  return new Promise((resolve, reject) => {
    converter.generatePdf(`./reports/${file}`, function(err, result) {
      if (result.status === 0) {
        resolve("next");
      }
    });
  });
};

const genMatrix = fileNames => {
  return new Promise(async (resolve, reject) => {
    for (let i = 0; i < fileNames.length; i++) {
      let file = `${fileNames[i]}`;
      if (i > 0) {
        file = `${file}.xlsx`;
      }
      const genPDF = await generatePdf(file).catch(err => {
        console.log(err);
      });
    }
    resolve("goMerge");
  });
};

router.post("/api/getpdf", async (req, res) => {
  let fileNames = req.body.fileNames;
  for (let i = 1; i < fileNames.length; i++) {
    const file = fileNames[i];
    const wb = new Excel.Workbook();
    wb.xlsx.readFile(`./reports/${file}.xlsx`).then(() => {
      wb.eachSheet(function(worksheet, sheetId) {
        const idToKeep = `Ar${i + 1}`;
        if (worksheet.name !== idToKeep) {
          wb.removeWorksheet(sheetId);
        } else {
          wb.xlsx.writeFile(`./reports/${file}.xlsx`);
        }
      });
    });
  }

  const pdfMatrix = await genMatrix(fileNames).catch(err => console.log(err));
  if (fileNames.length > 1) {
    const createMerge = await goMerge(fileNames).catch(err => console.log(err));
    const goClean = await clean(fileNames).catch(err => console.log(err));
  }
  const sendTheMail = await mail(fileNames).catch(err => console.log(err));

  const forSplit = fileNames.pop().split(".");
  const myPdf = `./reports/${forSplit[0]}`;

  res.send(myPdf);
});

module.exports = router;
