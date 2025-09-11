/*

const XLSX = require("xlsx");
const fs = require("fs");
const cheerio = require("cheerio"); // 📌 برای پردازش HTML
const fetch = (...args) =>
    import("node-fetch").then(({ default: fetch }) => fetch(...args));

const BASE_URL =
    "https://search.codal.ir/api/search/v2/q?&Category=-1&Childs=true&CompanyState=-1&CompanyType=-1&Consolidatable=true&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&search=true";

// 🛠 نرمال‌سازی عدد (فارسی → انگلیسی)
function normalizeNumber(str) {
    if (!str) return null;
    let normalized = str
        .toString()
        .replace(/[۰-۹]/g, (d) => "0123456789"["۰۱۲۳۴۵۶۷۸۹".indexOf(d)]) // اعداد فارسی
        .replace(/[^\d.-]/g, "") // حذف همه چیز غیر عدد
        .trim();
    return isNaN(normalized) || normalized === "" ? null : Number(normalized);
}

// 🛠 گرفتن گزارش‌ها از API
async function getReports(symbol, page = 1) {
    const url = `${BASE_URL}&Symbol=${encodeURIComponent(symbol)}&PageNumber=${page}`;
    console.log(`📡 Fetching page ${page}: ${url}`);
    const res = await fetch(url, {
        headers: {
            Accept: "application/json, text/plain, *!/!*",
            "User-Agent":
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/139 Safari/537.36",
        },
    });
    if (!res.ok) throw new Error(`❌ HTTP error! ${res.status}`);
    return res.json();
}

// 🛠 گرفتن همه صفحات
async function getAllReports(symbol) {
    let page = 1;
    let allReports = [];
    let totalPages = 1;
    do {
        const data = await getReports(symbol, page);
        if (data.Letters && data.Letters.length > 0) {
            allReports = allReports.concat(data.Letters);
        }
        if (page === 1) {
            totalPages = data.Page || 1;
            console.log(
                `✅ Symbol: ${symbol}, Total: ${data.Total}, Pages: ${totalPages}`
            );
        }
        page++;
    } while (page <= totalPages);
    return allReports;
}

// 🛠 پردازش اکسل
function extractFromWorkbook(wb, title) {
    let results = [];
    for (const sheetName of wb.SheetNames) {
        const sheet = wb.Sheets[sheetName];
        let foundSoodVaZian = false;

        for (const cellAddr in sheet) {
            const cell = sheet[cellAddr];
            if (!cell || !cell.v) continue;
            const val = String(cell.v).trim();

            if (val.includes("صورت سود") || val.includes("سود و زیان")) {
                foundSoodVaZian = true;
            }

            if (foundSoodVaZian && (val.includes("سرمایه") || val.includes("جمع"))) {
                const row = cellAddr.replace(/[A-Z]/g, "");
                let number = null;

                for (let colCode = 65; colCode <= 90; colCode++) {
                    const addr = String.fromCharCode(colCode) + row;
                    if (sheet[addr] && sheet[addr].v) {
                        const num = normalizeNumber(sheet[addr].v);
                        if (num !== null) {
                            number = num;
                            break;
                        }
                    }
                }

                if (number !== null && number !== 0) {
                    results.push({ title, label: val, value: number });
                    console.log(`✅ Found in [${title}] [${sheetName}]: ${val} = ${number}`);
                }
            }
        }
    }
    return results;
}

// 🛠 پردازش HTML
function extractFromHtml(html, title) {
    const $ = cheerio.load(html);
    let results = [];

    $("table tr").each((i, tr) => {
        const tds = $(tr).find("td").map((j, td) => $(td).text().trim()).get();
        if (tds.length >= 2) {
            const label = tds[0];
            const number = normalizeNumber(tds[1]);
            if (
                number !== null &&
                number !== 0 &&
                (label.includes("سرمایه") || label.includes("جمع"))
            ) {
                results.push({ title, label, value: number });
                console.log(`✅ Found in [${title}] [HTML]: ${label} = ${number}`);
            }
        }
    });

    if (results.length === 0) {
        console.log(`⚠️ No سرمایه/جمع found in HTML [${title}]`);
    }
    return results;
}

// 🛠 تشخیص نوع و پردازش
async function processFile(url, title) {
    console.log(`📥 Downloading: ${url}`);
    const res = await fetch(url);
    if (!res.ok) {
        console.log(`❌ Failed to download ${url}`);
        return [];
    }

    const buffer = await res.arrayBuffer();
    const contentType = res.headers.get("content-type") || "";

    // اکسل
    if (
        contentType.includes("spreadsheet") ||
        contentType.includes("excel") ||
        url.endsWith(".xls") ||
        url.endsWith(".xlsx")
    ) {
        try {
            const wb = XLSX.read(buffer, { type: "buffer" });
            return extractFromWorkbook(wb, title);
        } catch (e) {
            console.log(`⚠️ Excel parse failed, fallback to HTML`);
        }
    }

    // HTML
    const text = Buffer.from(buffer).toString("utf8");
    if (text.includes("<table")) {
        return extractFromHtml(text, title);
    }

    console.log(`⚠️ Unknown format for ${title}`);
    return [];
}

// 🛠 اجرای اصلی
async function main() {
    const symbol = "غکورش"; // نماد
    const reports = await getAllReports(symbol);

    const financials = reports.filter((r) => r.Title.includes("صورت‌های مالی"));
    console.log(`📌 گزارش‌های صورت‌های مالی: ${financials.length}`);

    let annualData = [];
    let interimData = [];

    for (const r of financials) {
        const fileUrl = r.ExcelUrl
            ? r.ExcelUrl
            : `https://excel.codal.ir/service/Excel/GetAll/${r.TracingNo}/0`;

        console.log(`\n📄 ${r.Title}`);
        console.log(`   📅 Date: ${r.PublishDateTime}`);

        const rows = await processFile(fileUrl, r.Title);

        if (r.Title.includes("سال مالی منتهی")) {
            annualData = annualData.concat(rows);
        } else if (r.Title.includes("میاندوره‌ای")) {
            interimData = interimData.concat(rows);
        }
    }

    // خروجی به اکسل
    const wb = XLSX.utils.book_new();

    if (annualData.length > 0) {
        const wsAnnual = XLSX.utils.json_to_sheet(annualData);
        XLSX.utils.book_append_sheet(wb, wsAnnual, "صورت‌های مالی سالانه");
    }

    if (interimData.length > 0) {
        const wsInterim = XLSX.utils.json_to_sheet(interimData);
        XLSX.utils.book_append_sheet(wb, wsInterim, "میاندوره‌ای");
    }

    const outFile = `${symbol}-12month.xlsx`;
    XLSX.writeFile(wb, outFile);
    console.log(`\n✅ فایل خروجی ساخته شد: ${outFile}`);
}

main().catch((err) => console.error("❌ Error in main:", err));*/


const XLSX = require("xlsx");
const fs = require("fs");
const cheerio = require("cheerio"); // 📌 برای پردازش HTML
const readline = require("readline"); // 📌 برای گرفتن ورودی کاربر
const fetch = (...args) =>
    import("node-fetch").then(({ default: fetch }) => fetch(...args));

const BASE_URL =
    "https://search.codal.ir/api/search/v2/q?&Category=-1&Childs=true&CompanyState=-1&CompanyType=-1&Consolidatable=true&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&search=true";

// 🛠 نرمال‌سازی عدد (فارسی → انگلیسی)
function normalizeNumber(str) {
    if (!str) return null;
    let normalized = str
        .toString()
        .replace(/[۰-۹]/g, (d) => "0123456789"["۰۱۲۳۴۵۶۷۸۹".indexOf(d)])
        .replace(/[^\d.-]/g, "")
        .trim();
    return isNaN(normalized) || normalized === "" ? null : Number(normalized);
}

// 🛠 گرفتن گزارش‌ها از API
async function getReports(symbol, page = 1) {
    const url = `${BASE_URL}&Symbol=${encodeURIComponent(symbol)}&PageNumber=${page}`;
    console.log(`📡 Fetching page ${page}: ${url}`);
    const res = await fetch(url, {
        headers: {
            Accept: "application/json, text/plain, */*",
            "User-Agent":
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/139 Safari/537.36",
        },
    });
    if (!res.ok) throw new Error(`❌ HTTP error! ${res.status}`);
    return res.json();
}

// 🛠 گرفتن همه صفحات
async function getAllReports(symbol) {
    let page = 1;
    let allReports = [];
    let totalPages = 1;
    do {
        const data = await getReports(symbol, page);
        if (data.Letters && data.Letters.length > 0) {
            allReports = allReports.concat(data.Letters);
        }
        if (page === 1) {
            totalPages = data.Page || 1;
            console.log(
                `✅ Symbol: ${symbol}, Total: ${data.Total}, Pages: ${totalPages}`
            );
        }
        page++;
    } while (page <= totalPages);
    return allReports;
}

// 🛠 پردازش اکسل
function extractFromWorkbook(wb, title) {
    let results = [];
    for (const sheetName of wb.SheetNames) {
        const sheet = wb.Sheets[sheetName];
        let foundSoodVaZian = false;

        for (const cellAddr in sheet) {
            const cell = sheet[cellAddr];
            if (!cell || !cell.v) continue;
            const val = String(cell.v).trim();

            if (val.includes("صورت سود") || val.includes("سود و زیان")) {
                foundSoodVaZian = true;
            }

            if (foundSoodVaZian && (val.includes("سرمایه") || val.includes("جمع"))) {
                const row = cellAddr.replace(/[A-Z]/g, "");
                let number = null;

                for (let colCode = 65; colCode <= 90; colCode++) {
                    const addr = String.fromCharCode(colCode) + row;
                    if (sheet[addr] && sheet[addr].v) {
                        const num = normalizeNumber(sheet[addr].v);
                        if (num !== null) {
                            number = num;
                            break;
                        }
                    }
                }

                if (number !== null && number !== 0) {
                    results.push({ title, label: val, value: number });
                    console.log(`✅ Found in [${title}] [${sheetName}]: ${val} = ${number}`);
                }
            }
        }
    }
    return results;
}

// 🛠 پردازش HTML
function extractFromHtml(html, title) {
    const $ = cheerio.load(html);
    let results = [];

    $("table tr").each((i, tr) => {
        const tds = $(tr).find("td").map((j, td) => $(td).text().trim()).get();
        if (tds.length >= 2) {
            const label = tds[0];
            const number = normalizeNumber(tds[1]);
            if (
                number !== null &&
                number !== 0 &&
                (label.includes("سرمایه") || label.includes("جمع"))
            ) {
                results.push({ title, label, value: number });
                console.log(`✅ Found in [${title}] [HTML]: ${label} = ${number}`);
            }
        }
    });

    if (results.length === 0) {
        console.log(`⚠️ No سرمایه/جمع found in HTML [${title}]`);
    }
    return results;
}

// 🛠 تشخیص نوع و پردازش
async function processFile(url, title) {
    console.log(`📥 Downloading: ${url}`);
    const res = await fetch(url);
    if (!res.ok) {
        console.log(`❌ Failed to download ${url}`);
        return [];
    }

    const buffer = await res.arrayBuffer();
    const contentType = res.headers.get("content-type") || "";

    if (
        contentType.includes("spreadsheet") ||
        contentType.includes("excel") ||
        url.endsWith(".xls") ||
        url.endsWith(".xlsx")
    ) {
        try {
            const wb = XLSX.read(buffer, { type: "buffer" });
            return extractFromWorkbook(wb, title);
        } catch (e) {
            console.log(`⚠️ Excel parse failed, fallback to HTML`);
        }
    }

    const text = Buffer.from(buffer).toString("utf8");
    if (text.includes("<table")) {
        return extractFromHtml(text, title);
    }

    console.log(`⚠️ Unknown format for ${title}`);
    return [];
}

// 🛠 اجرای اصلی
async function main(symbol) {
    const reports = await getAllReports(symbol);

    const financials = reports.filter((r) => r.Title.includes("صورت‌های مالی"));
    console.log(`📌 گزارش‌های صورت‌های مالی: ${financials.length}`);

    let annualData = [];
    let interimData = [];

    for (const r of financials) {
        const fileUrl = r.ExcelUrl
            ? r.ExcelUrl
            : `https://excel.codal.ir/service/Excel/GetAll/${r.TracingNo}/0`;

        console.log(`\n📄 ${r.Title}`);
        console.log(`   📅 Date: ${r.PublishDateTime}`);

        const rows = await processFile(fileUrl, r.Title);

        if (r.Title.includes("سال مالی منتهی")) {
            annualData = annualData.concat(rows);
        } else if (r.Title.includes("میاندوره‌ای")) {
            interimData = interimData.concat(rows);
        }
    }

    const wb = XLSX.utils.book_new();

    if (annualData.length > 0) {
        const wsAnnual = XLSX.utils.json_to_sheet(annualData);
        XLSX.utils.book_append_sheet(wb, wsAnnual, "صورت‌های مالی سالانه");
    }

    if (interimData.length > 0) {
        const wsInterim = XLSX.utils.json_to_sheet(interimData);
        XLSX.utils.book_append_sheet(wb, wsInterim, "میاندوره‌ای");
    }

    const outFile = `${symbol}-12month.xlsx`;
    XLSX.writeFile(wb, outFile);
    console.log(`\n✅ فایل خروجی ساخته شد: ${outFile}`);
}

// 📌 گرفتن ورودی از کاربر
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
});

rl.question("🔎 lotfan namad borsi ra vared konid befarsi va format estefade shode dar codal: ", (symbol) => {
    if (!symbol || symbol.trim() === "") {
        console.log("❌ نماد وارد نشد!");
        rl.close();
        return;
    }
    rl.close();
    main(symbol.trim()).catch((err) => console.error("❌ Error in main:", err));
});