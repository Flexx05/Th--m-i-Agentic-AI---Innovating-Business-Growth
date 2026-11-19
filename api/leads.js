const { google } = require("googleapis");

function createAuth() {
  return new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
}

function escapeSheetTitle(title) {
  return `'${String(title).replace(/'/g, "''")}'`;
}

async function resolveSheetTitle(sheets, spreadsheetId) {
  const configuredTab = process.env.GOOGLE_SHEET_TAB?.trim();
  if (configuredTab) return configuredTab;

  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties.title",
  });
  const titles = (meta.data.sheets || [])
    .map((sheet) => sheet.properties?.title)
    .filter(Boolean);

  if (titles.includes("Leads")) return "Leads";
  if (titles.length > 0) return titles[0];
  return "Leads";
}

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({
      success: false,
      message: "Method not allowed",
    });
  }

  // Kiểm tra nguồn request
  const origin = req.headers.origin;
  const allowedOrigins = (process.env.SITE_ORIGIN || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  if (origin && allowedOrigins.length > 0 && !allowedOrigins.includes(origin)) {
    return res.status(403).json({
      success: false,
      message: "Invalid origin",
    });
  }

  try {
    const { fullName, phone, email, company, position, website } = req.body || {};

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
    const companyValue = typeof company === "string" ? company.trim().slice(0, 100) : "";
    const positionValue = typeof position === "string" ? position.trim().slice(0, 100) : "";

    if (emailValue && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailValue)) {
      return res.status(400).json({
        success: false,
        message: "Invalid email",
      });
    }

    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    const sheets = google.sheets({
      version: "v4",
      auth: await createAuth().getClient(),
    });

    const sheetTitle = await resolveSheetTitle(sheets, spreadsheetId);
    const sheetRef = escapeSheetTitle(sheetTitle);

    const existingRows = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetRef}!A:A`,
    });

    const rows = existingRows.data.values || [];
    const lastSerial = Number(rows[rows.length - 1]?.[0]);
    const serialNumber = Number.isSafeInteger(lastSerial) && lastSerial > 0
      ? lastSerial + 1
      : Math.max(rows.length, 1);

    const values = [];
    if (rows.length === 0) {
      values.push([
        "STT",
        "Họ và tên",
        "Số điện thoại",
        "Email",
        "Công ty",
        "Chức vụ",
        "Thời gian đăng ký",
      ]);
    }
    values.push([
      serialNumber,
      fullName.trim(),
      phone.trim(),
      emailValue,
      companyValue,
      positionValue,
      new Date().toISOString(),
    ]);

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetRef}!A:G`,
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
