/*
// index.js
const XLSX = require("xlsx");
const fetch = (...args) =>
    import("node-fetch").then(({ default: fetch }) => fetch(...args));

const BASE_Q =
    "https://search.codal.ir/api/search/v2/q?Audited=true&AuditorRef=-1&Category=3&Childs=true&CompanyState=0&CompanyType=1&Consolidatable=true&IndustryGroup=70&IsNotAudited=false&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&Publisher=false&ReportingType=1000002&TracingNo=-1&search=true";

// -------------------- helpers --------------------
const UA =
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36";

function faDigitsToEn(s) {
    if (!s) return "";
    return s.replace(/[۰-۹]/g, (d) => "0123456789"["۰۱۲۳۴۵۶۷۸۹".indexOf(d)]);
}

function normalizeNumber(str) {
    if (str === undefined || str === null) return null;
    let s = String(str);
    s = faDigitsToEn(s);
    // (123) -> -123
    s = s.replace(/\(/g, "-").replace(/\)/g, "");
    // remove thousand separators and non-numeric except . and -
    s = s.replace(/[,،\s]/g, "").replace(/[^\d.-]/g, "");
    if (s === "" || s === "-" || s === "." || isNaN(Number(s))) return null;
    return Number(s);
}

function normalizeFaText(s) {
    if (!s) return "";
    return s
        .replace(/\u200c|‌/g, " ") // ZWNJ to space
        .replace(/ي/g, "ی")
        .replace(/ك/g, "ک")
        .replace(/\s+/g, " ")
        .trim();
}

function extractJalaliDateFromTitle(title) {
    // منتهی به  ۱۴۰۴/۰۵/۳۱
    const m = title.match(/منتهی\s+به\s+([۰-۹0-9\/]+)/);
    if (!m) return null;
    let d = faDigitsToEn(m[1]).split(/\s+/)[0]; // "1404/05/31"
    return d;
}

function jalaliKey(d) {
    // "1404/05/31" -> "14040531" for sorting
    return d ? d.replace(/\//g, "") : "";
}

function parsePublishTs(s) {
    // "۱۴۰۴/۰۶/۰۵ ۱۲:۳۵:۰۹" -> number for comparison (no real calendar conversion needed for ordering)
    if (!s) return 0;
    s = faDigitsToEn(s);
    const m = s.match(
        /(\d{4})\/(\d{2})\/(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/
    );
    if (!m) return 0;
    const [, Y, M, D, h, m2, s2] = m.map((x) => Number(x));
    // Treat as a big integer-like key (Jalali but fine for ordering)
    return Y * 1e10 + M * 1e8 + D * 1e6 + h * 1e4 + m2 * 1e2 + s2;
}

function looksLikeHtml(buf) {
    try {
        const txt = Buffer.from(buf).toString("utf8");
        return /<html|<!doctype/i.test(txt);
    } catch {
        return false;
    }
}

function isExcludedRealEstateSum(label) {
    // حذف «جمع سرمایه‌گذاری در املاک» با انواع املاء
    const t = normalizeFaText(label);
    if (!t.startsWith("جمع")) return false;
    // بسیار tolerant: شامل "سرما" و "املاک"
    return t.includes("سرما") && t.includes("املاک");
}

// -------------------- fetch search with pagination --------------------
async function fetchLettersAllPages({ symbol, name }) {
    let page = 1;
    let totalPages = 1;
    const all = [];

    do {
        const url = `${BASE_Q}&Symbol=${encodeURIComponent(
            symbol
        )}&Name=${encodeURIComponent(name)}&name=${encodeURIComponent(
            name
        )}&PageNumber=${page}`;
        console.log(`📡 Fetching page ${page}: ${url}`);

        const res = await fetch(url, {
            headers: {
                Accept: "application/json, text/plain, *!/!*",
                "User-Agent": UA,
                Referer: "https://www.codal.ir/",
                Origin: "https://www.codal.ir",
                "Accept-Language":
                    "fa-IR,fa;q=0.9,en-GB;q=0.8,en;q=0.7,ar;q=0.5,en-US;q=0.4",
                "Cache-Control": "no-cache",
                Pragma: "no-cache",
            },
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();

        if (page === 1) {
            totalPages = data.Page || 1;
            console.log(
                `🧾 Total: ${data.Total} | Pages: ${totalPages} | This page: ${
                    data.Letters?.length || 0
                }`
            );
        }

        if (Array.isArray(data.Letters)) all.push(...data.Letters);
        page++;
    } while (page <= totalPages);

    return all;
}

// -------------------- excel parsing --------------------
async function parseExcelReport(excelUrl, title, dateStr) {
    console.log(`📥 Excel: ${excelUrl}`);
    const res = await fetch(excelUrl, {
        headers: {
            Accept:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel, *!/!*",
            "User-Agent": UA,
            Referer: "https://www.codal.ir/",
            Origin: "https://www.codal.ir",
        },
    });
    if (!res.ok) {
        console.log(`❌ Excel fetch failed: ${res.status} for "${title}"`);
        return [];
    }

    const buf = await res.arrayBuffer();
    if (looksLikeHtml(buf)) {
        console.log("⚠️ Excel URL returned HTML → skipped");
        return [];
    }

    let wb;
    try {
        wb = XLSX.read(buf, { type: "buffer" });
    } catch (e) {
        console.log(`⚠️ XLSX parse error for "${title}": ${e.message}`);
        return [];
    }

    const out = [];
    for (const sheetName of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
        for (const row of rows) {
            if (!Array.isArray(row) || row.length === 0) continue;
            const labelRaw = row[0] ? String(row[0]).trim() : "";
            if (!labelRaw) continue;

            const labelNorm = normalizeFaText(labelRaw);

            // فقط ردیف‌هایی که با "جمع" شروع میشن
            if (!labelNorm.startsWith("جمع")) continue;

            // ❌ حذف ردیف‌هایی که کلمه "املاک" دارن
            if (labelNorm.includes("املاک")) continue;

            // مقدار = آخرین عدد توی ردیف
            let value = null;
            for (let i = 1; i < row.length; i++) {
                const v = normalizeNumber(row[i]);
                if (v !== null) value = v;
            }
            if (value === null || value === 0) continue;

            out.push({
                date: dateStr,
                title,
                sheet: sheetName,
                label: labelRaw,
                value,
            });
        }
    }

    if (out.length) {
        console.log(`✅ ${title} → ${out.length} سطر "جمع" ذخیره شد`);
    } else {
        console.log(`⚠️ ${title} → ردیف "جمع" پیدا نشد`);
    }
    return out;
}
// -------------------- main --------------------
async function main() {
    // قابل تغییر
    const symbol = "وآذر";
    const name = "سرمایه گذاری توسعه آذربایجان";

    // 1) تمام صفحات
    const allLetters = await fetchLettersAllPages({ symbol, name });

    // 2) فقط گزارش‌های فعالیت ماهانه
    const monthly = allLetters.filter((l) =>
        String(l.Title || "").includes("گزارش فعالیت ماهانه")
    );

    console.log(`📑 Monthly letters found: ${monthly.length}`);

    // 3) گروه‌بندی بر اساس تاریخ داخل Title و انتخاب آخرین انتشار برای هر تاریخ
    const bestByDate = new Map(); // key: "YYYY/MM/DD" → letter
    for (const l of monthly) {
        const d = extractJalaliDateFromTitle(l.Title || "");
        if (!d) continue;
        const ts = parsePublishTs(l.PublishDateTime || l.SentDateTime || "");
        const prev = bestByDate.get(d);
        if (!prev || ts > prev._ts) {
            bestByDate.set(d, { ...l, _ts: ts });
        }
    }

    const chosen = [...bestByDate.entries()]
        .map(([d, l]) => ({ date: d, letter: l }))
        // اختیار: جدیدترین تاریخ‌ها آخر یا اول؟ (اینجا صعودی)
        .sort((a, b) => jalaliKey(a.date).localeCompare(jalaliKey(b.date)));

    console.log(
        `🗂 Unique periods (by title date): ${chosen.length} (deduped from ${monthly.length})`
    );

    // 4) خواندن اکسل‌های انتخابی و استخراج فقط سطرهای "جمع" (به‌جز املاک)
    let results = [];
    for (const item of chosen) {
        const l = item.letter;
        const dateStr = item.date;

        const excelUrl =
            l.ExcelUrl ||
            `https://excel.codal.ir/service/Excel/GetAll/${encodeURIComponent(
                l.TracingNo || ""
            )}/0`;

        const rows = await parseExcelReport(excelUrl, l.Title, dateStr);
        results.push(...rows);
    }

    if (!results.length) {
        console.log("⛔ هیچ داده‌ای برای خروجی پیدا نشد");
        return;
    }

    // 5) سورت نهایی بر اساس تاریخ
    results.sort((a, b) =>
        jalaliKey(a.date).localeCompare(jalaliKey(b.date))
    );

    // 6) خروجی اکسل
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(results);
    XLSX.utils.book_append_sheet(wb, ws, "گزارش فعالیت ماهانه - جمع");
    const outFile = `${symbol}-monthly.xlsx`;
    XLSX.writeFile(wb, outFile);

    console.log(`📊 خروجی ذخیره شد: ${outFile}`);
}

main().catch((err) => console.error("❌", err));
*/
const XLSX = require("xlsx");
const fetch = (...args) =>
    import("node-fetch").then(({ default: fetch }) => fetch(...args));

// 🔗 آدرس پایه برای سرچ
const BASE_URL =
    "https://search.codal.ir/api/search/v2/q?Audited=true&AuditorRef=-1&Category=3&Childs=true&CompanyState=0&CompanyType=1&Consolidatable=true&IndustryGroup=70&IsNotAudited=false&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&Publisher=false&ReportingType=1000002&TracingNo=-1&search=true";
function normalizeText(str) {
    if (!str) return "";
    return str
        .replace(/ي/g, "ی") // ی عربی → ی فارسی
        .replace(/ك/g, "ک") // ک عربی → ک فارسی
        .replace(/\s+/g, "") // حذف فاصله‌ها
        .replace(/\u200c/g, ""); // حذف نیم‌فاصله
}
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

// 🛠 گرفتن همه صفحات گزارش‌ها
async function getAllReports(symbol, name) {
    let page = 1;
    let allLetters = [];
    let total = 0;

    while (true) {
        const url = `${BASE_URL}&Symbol=${encodeURIComponent(
            symbol
        )}&Name=${encodeURIComponent(name)}&name=${encodeURIComponent(
            name
        )}&PageNumber=${page}`;
        console.log(`📡 Fetching: صفحه ${page} → ${url}`);

        const res = await fetch(url, {
            headers: {
                Accept: "application/json, text/plain, */*",
                "User-Agent":
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
                Referer: "https://www.codal.ir/",
                Origin: "https://www.codal.ir",
            },
        });

        if (!res.ok) throw new Error(`❌ HTTP ${res.status}`);
        const data = await res.json();

        if (page === 1) total = data.Total;
        allLetters = allLetters.concat(data.Letters);

        // اگر رسیدیم به آخر صفحه‌ها، متوقف شو
        if (allLetters.length >= total) break;

        page++;
    }

    console.log(`📑 کل گزارش‌ها: ${allLetters.length} از ${total}`);
    return allLetters;
}

// 🛠 دانلود و خواندن اکسل
async function parseExcel(url, title) {
    console.log(`📥 Download Excel: ${url}`);
    const res = await fetch(url, {
        headers: {
            Accept:
                "application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, */*",
            "User-Agent":
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
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
                    const cleanLabel = normalizeText(label);
                    const value = normalizeNumber(row[1]);

                    // فقط "جمع"هایی که "سرمایه گذاری در املاک" نیستند
                    if (
                        cleanLabel.includes("جمع") &&
                        !cleanLabel.includes("جمعسرمایهگذاریدراملاک") &&
                        value !== null &&
                        value !== 0
                    ) {
                        results.push({ title, sheet: sheetName, label, value });
                    }
                });
            });

            if (results.length > 0) {
                console.log(`✅ ${title} → جمع معتبر پیدا شد`);
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

    const letters = await getAllReports(symbol, name);

    let allResults = [];

    for (const r of letters) {
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

    // مرتب‌سازی بر اساس تاریخ داخل عنوان
    allResults.sort((a, b) => {
        const dateA = (a.title.match(/\d{4}\/\d{2}\/\d{2}/) || [])[0] || "";
        const dateB = (b.title.match(/\d{4}\/\d{2}\/\d{2}/) || [])[0] || "";
        return dateA.localeCompare(dateB, "fa");
    });

    // ذخیره در اکسل
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(allResults);
    XLSX.utils.book_append_sheet(wb, ws, "گزارش فعالیت ماهانه");
    const outFile = `${symbol}-monthly.xlsx`;
    XLSX.writeFile(wb, outFile);

    console.log(`📊 خروجی ذخیره شد: ${outFile}`);
}

main().catch((err) => console.error("❌", err));
