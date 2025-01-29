require('dotenv').config(); // Cargar variables de entorno

const express = require("express");
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
  oAuth2Client.setCredentials({ refresh_token: REFRESH_TOKEN });
  return oAuth2Client;
}

// Enviar correo con el archivo adjunto
async function sendEmail(filePath, fileName) {
  try {
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

    await gmail.users.messages.send({
      userId: "me",
      requestBody: { raw: encodedMessage },
    });
    console.log("Correo enviado con éxito");
  } catch (error) {
    console.error("Error al enviar el correo:", error);
  }
}

// Función para descomponer la descripción
function descomponerDescripcion(texto) {
  if (!texto) return { FECHApago: "", MONTOcuota: "", CANTIDADcuotas: "", LUGARpago: "", MONTOtotal: "" };

  const regexFechaPago = /Primer vencimiento:\s*(\d{2}\/\d{2}\/\d{4})/;
  const regexMontoCuota = /Monto de cuota:\s*(\d+(?:\.\d{1,2})?)/;
  const regexCantidadCuotas = /Cantidad de cuota\/s:\s*(\d+)/;
  const regexLugarPago = /Lugar de pago:\s*([\w\s]+)/;
  const regexMontoTotal = /Deuda mÃ¡xima total:\s*(\d+(?:\.\d{1,2})?)/;

  return {
    FECHApago: texto.match(regexFechaPago)?.[1] || "",
    MONTOcuota: texto.match(regexMontoCuota)?.[1] || "",
    CANTIDADcuotas: texto.match(regexCantidadCuotas)?.[1] || "",
    LUGARpago: texto.match(regexLugarPago)?.[1]?.trim() || "",
    MONTOtotal: texto.match(regexMontoTotal)?.[1] || "",
  };
}

// Generar y enviar el archivo Excel
async function generateAndSendEmail() {
  try {
    const fecha = new Date();
    const dia = String(fecha.getDate()).padStart(2, "0");
    const mes = String(fecha.getMonth() + 1).padStart(2, "0");
    const processedFilePath = path.join(__dirname, "uploads", `acuerdos-${dia}-${mes}.xlsx`);

    if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");

    // Datos de ejemplo
    const newdata1 = [
      {
        "ctacteca_dtAlta": "14/01/2025 16:35:42",
        "personas_nNumeDocu": "5666666",
        "personas_cContacto": "Maria ferreira",
        "descripcion": "Deuda adquirida: FUCEREP TARDIA - Deuda mínima total: 3800.00 - Deuda máxima total: 3800.00 - Cantidad de cuota/s: 1 - Monto de cuota: 1934 - Primer vencimiento: 27/01/2025 - Observación corta: 1*1934-abitab-27/01 - Lugar de pago: abitab",
      }
    ];

    const newdata = newdata1.map(row => {
      const desc = descomponerDescripcion(row.descripcion);
      return {
        CODIGO: 13,
        DOCUMENTO: row.personas_nNumeDocu || "",
        NOMBRE: row.personas_cContacto || "",
        ENTREGA: 0,
        "MONTO CUOTA": desc.MONTOcuota || "",
        "CANTIDAD DE CUOTAS": desc.CANTIDADcuotas || "",
        "FECHA GESTION": "21/01/2025",
        "MONTO TOTAL": desc.MONTOtotal || "",
        "FECHA PAGO": desc.FECHApago || "",
        SUCURSAL: desc.LUGARpago || "",
        PRD: "",
      };
    });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Acuerdos");

    const headers = Object.keys(newdata[0]);
    worksheet.addRow(headers);
    newdata.forEach(row => worksheet.addRow(headers.map(key => row[key] || "")));

    const columnWidths = [15, 30, 15, 20, 20, 25, 30, 20, 20, 15, 15];
    columnWidths.forEach((width, i) => worksheet.getColumn(i + 1).width = width);

    await workbook.xlsx.writeFile(processedFilePath);
    await sendEmail(processedFilePath, path.basename(processedFilePath));
    console.log("Correo enviado con éxito");
  } catch (error) {
    console.error("Error al procesar y enviar el archivo:", error);
  } finally {
    if (fs.existsSync(processedFilePath)) fs.unlinkSync(processedFilePath);
  }
}

// Ejecutar la función
generateAndSendEmail();
