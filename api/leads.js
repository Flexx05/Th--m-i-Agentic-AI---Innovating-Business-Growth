const { google } = require("googleapis");

const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
  },
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({
      success: false,
      message: "Method not allowed",
    });
  }

  // Kiểm tra nguồn request
  const origin = req.headers.origin;

  if (origin && origin !== process.env.SITE_ORIGIN) {
    return res.status(403).json({
      success: false,
      message: "Invalid origin",
    });
  }

  try {
    const { fullName, phone, email, website } = req.body || {};

    if (website) {
      return res.status(200).json({
        success: true,
        message: "Submitted",
      });
    }

    // Validate dữ liệu
    if (
      typeof fullName !== "string" ||
      typeof phone !== "string" ||
      fullName.trim().length < 2 ||
      fullName.trim().length > 100 ||
      !/^[0-9+().\s-]{8,20}$/.test(phone.trim())
    ) {
      return res.status(400).json({
        success: false,
        message: "Invalid customer information",
      });
    }

    const emailValue = typeof email === "string" ? email.trim() : "";

    if (emailValue && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailValue)) {
      return res.status(400).json({
        success: false,
        message: "Invalid email",
      });
    }

    const sheets = google.sheets({
      version: "v4",
      auth: await auth.getClient(),
    });

    const existingRows = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Leads!A:A",
    });

    const rows = existingRows.data.values || [];
    const lastSerial = Number(rows[rows.length - 1]?.[0]);
    const serialNumber = Number.isSafeInteger(lastSerial) && lastSerial > 0
      ? lastSerial + 1
      : Math.max(rows.length, 1);
    const values = [
      [
        serialNumber,
        fullName.trim(),
        phone.trim(),
        emailValue,
        new Date().toISOString(),
      ],
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Leads!A:E",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values,
      },
    });

    return res.status(200).json({
      success: true,
      message: "Customer information saved successfully",
    });
  } catch (error) {
    console.error("Submit lead error:", error);

    return res.status(500).json({
      success: false,
      message: "Unable to save customer information",
    });
  }
};
