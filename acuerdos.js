require('dotenv').config();  // Cargar variables de entorno

const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
const ExcelJS = require("exceljs");

// Accede a las variables de entorno
const { CLIENT_ID, CLIENT_SECRET, REDIRECT_URI, REFRESH_TOKEN } = process.env;

const app = express();

// Autenticación de Gmail con las credenciales de las variables de entorno
async function authenticate() {
  const oAuth2Client = new google.auth.OAuth2(
    CLIENT_ID,
    CLIENT_SECRET,
    REDIRECT_URI
  );

  // Usar el refresh token para obtener nuevas credenciales
  oAuth2Client.setCredentials({
    refresh_token: REFRESH_TOKEN
  });

  return oAuth2Client;
}


// Enviar correo con el archivo adjunto
async function sendEmail(filePath, fileName) {
  const auth = await authenticate();
  const gmail = google.gmail({ version: "v1", auth });

  const attachment = fs.readFileSync(filePath).toString("base64");

  const rawMessage = [
    "From: carolainsilva1@gmail.com",
    "To: carolainsilva1@gmail.com",
    "Subject: Acuerdos Credito Directos Capta",
    "Content-Type: multipart/mixed; boundary=boundary_string",
    "",
    "--boundary_string",
    "Content-Type: text/plain; charset=UTF-8",
    "",
    "Estimados/as, espero que se encuentren bien, les envío los acuerdos generados en el día de hoy. ¡Saludos!, Capta.",
    "",
    "--boundary_string",
    `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="${fileName}"`,
    `Content-Disposition: attachment; filename="${fileName}"`,
    "Content-Transfer-Encoding: base64",
    "",
    attachment,
    "",
    "--boundary_string--",
  ].join("\r\n");

  const encodedMessage = Buffer.from(rawMessage)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");

  const response = await gmail.users.messages.send({
    userId: "me",
    requestBody: {
      raw: encodedMessage,
    },
  });

  if (response.data.id) {
    console.log("Correo enviado con éxito:", response.data);
  } else {
    console.error("Error al enviar el correo:", response);
  }
}

// Generar y enviar el archivo Excel
async function generateAndSendEmail() {
  const fecha = new Date();
  const dia = String(fecha.getDate()).padStart(2, "0");
  const mes = String(fecha.getMonth() + 1).padStart(2, "0");
  const processedFilePath = path.join(
    __dirname,
    "uploads",
    `acuerdos-${dia}-${mes}.xlsx`
  );

  // Datos proporcionados
  const newdata = [
    {
      "CODIGO": 13,
      "DOCUMENTO": 22222,
      "NOMBRE": "FERNANDO DA SILVA",
      "ENTREGA": 0,
      "MONTOcuota": "1635",
      "CANTIDADcuotas": "3",
      "FECHAgestion": "21/01/2025",
      "MONTOtotal": "57427.00",
      "FECHApago": "24/01/2025",
      "SUCURSAL": "SUCURSAL CERRO",
      "PRD": ""
    }
  ];

  try {
    // Crear un archivo Excel usando ExcelJS
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Acuerdos");

    // Agregar encabezados
    const headers = [
      "CODIGO",
      "DOCUMENTO",
      "NOMBRE",
      "ENTREGA",
      "MONTOcuota",
      "CANTIDADcuotas",
      "FECHAgestion",
      "MONTOtotal",
      "FECHApago",
      "SUCURSAL",
      "PRD",
    ];
    worksheet.addRow(headers);

    // Agregar los datos de newdata al archivo Excel
    newdata.forEach((row) => {
      const rowData = headers.map((key) => row[key] || "");
      worksheet.addRow(rowData);
    });

    // Escribir el archivo Excel
    await workbook.xlsx.writeFile(processedFilePath);

    // Enviar el archivo Excel por correo
    await sendEmail(processedFilePath, path.basename(processedFilePath));

    console.log("Correo enviado con éxito");
  } catch (error) {
    console.error("Error al procesar y enviar el archivo:", error);
  } finally {
    // Eliminar el archivo generado
    if (fs.existsSync(processedFilePath)) {
      fs.unlinkSync(processedFilePath);
    }
  }
}

// Ejecutar la función
generateAndSendEmail();

