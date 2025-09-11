const XLSX = require("xlsx");
const fetch = (...args) =>
    import("node-fetch").then(({ default: fetch }) => fetch(...args));

const BASE_URL =
    "https://search.codal.ir/api/search/v2/q?Audited=true&AuditorRef=-1&Category=3&Childs=true&CompanyState=0&CompanyType=1&Consolidatable=true&IndustryGroup=70&IsNotAudited=false&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&PageNumber=1&Publisher=false&ReportingType=1000002&TracingNo=-1&search=true";

// 🛠 نرمال‌سازی عدد
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
async function getReports(symbol, name) {
    const url = `${BASE_URL}&Symbol=${encodeURIComponent(
        symbol
    )}&Name=${encodeURIComponent(name)}&name=${encodeURIComponent(name)}`;
    console.log(`📡 Fetching: ${url}`);

    const res = await fetch(url, {
        headers: {
            Accept: "application/json, text/plain, */*",
            "User-Agent":
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0.0.0 Safari/537.36",
            Referer: "https://www.codal.ir/",
            Origin: "https://www.codal.ir",
        },
    });

    if (!res.ok) throw new Error(`❌ HTTP ${res.status}`);
    return res.json();
}

// 🛠 دانلود و خواندن اکسل
async function parseExcel(url, title) {
    console.log(`📥 Download Excel: ${url}`);
    const res = await fetch(url, {
        headers: {
            Accept:
                "application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, */*",
            "User-Agent":
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0.0.0 Safari/537.36",
            Referer: "https://www.codal.ir/",
            Origin: "https://www.codal.ir",
        },
    });

    if (!res.ok) {
        console.log(`❌ Failed to fetch excel for ${title}`);
        return null;
    }

    const buffer = await res.arrayBuffer();
    const contentType = res.headers.get("content-type") || "";

    if (contentType.includes("excel") || contentType.includes("spreadsheet")) {
        try {
            const wb = XLSX.read(buffer, { type: "buffer" });
            let results = [];

            wb.SheetNames.forEach((sheetName) => {
                const sheet = wb.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                rows.forEach((row) => {
                    if (!row || row.length < 2) return;
                    const label = row[0] ? row[0].toString().trim() : "";
                    const value = normalizeNumber(row[1]);

                    // 📌 فقط ردیف‌های "جمع" و حذف "جمع سرمایه‌گذاری در املاک"
                    if (
                        label.startsWith("جمع") &&
                        !label.includes("جمع سرمایه‌گذاری در املاک") &&
                        value !== null &&
                        value !== 0
                    ) {
                        results.push({ title, sheet: sheetName, label, value });
                    }
                });
            });

            if (results.length > 0) {
                console.log(`✅ ${title} → ${results.length} ردیف جمع ذخیره شد`);
            } else {
                console.log(`⚠️ ${title} → جمع پیدا نشد`);
            }
            return results;
        } catch (e) {
            console.log(`⚠️ Parse error for ${title}: ${e.message}`);
            return null;
        }
    } else {
        console.log("⚠️ این لینک اکسل واقعی نیست (HTML برگشته) → رد شد");
        return null;
    }
}

// 🛠 اجرای اصلی
async function main() {
    const symbol = "وآذر";
    const name = "سرمایه گذاری توسعه آذربایجان";

    const data = await getReports(symbol, name);
    console.log(`📑 تعداد گزارش‌ها: ${data.Letters.length}`);

    let allResults = [];

    for (const r of data.Letters) {
        if (!r.Title.includes("گزارش فعالیت ماهانه")) continue;

        const excelUrl = r.ExcelUrl;
        if (!excelUrl) continue;

        const rows = await parseExcel(excelUrl, r.Title);
        if (rows) allResults = allResults.concat(rows);
    }

    if (allResults.length === 0) {
        console.log("⛔ هیچ دیتایی ذخیره نشد");
        return;
    }

    // خروجی به اکسل
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(allResults);
    XLSX.utils.book_append_sheet(wb, ws, "گزارش فعالیت ماهانه");
    const outFile = `${symbol}-monthly.xlsx`;
    XLSX.writeFile(wb, outFile);

    console.log(`📊 خروجی ذخیره شد: ${outFile}`);
}

main().catch((err) => console.error(err));
