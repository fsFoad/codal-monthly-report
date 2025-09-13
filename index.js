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
/*

// const { fetchAllSymbols } = require("./symbols");
const XLSX = require("xlsx");
const fetch = (...args) =>
    import("node-fetch").then(({ default: fetch }) => fetch(...args));

// 🔗 آدرس پایه برای سرچ
const BASE_SYMBOLS =
    "https://search.codal.ir/api/search/v2/q?Audited=true&AuditorRef=-1&Category=-1&Childs=true&CompanyState=-1&CompanyType=-1&Consolidatable=true&IsNotAudited=false&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&Publisher=false&ReportingType=-1&TracingNo=-1&search=false";function normalizeText(str) {
    const BASE_URL =
        "https://search.codal.ir/api/search/v2/q?Audited=true&AuditorRef=-1&Category=3&Childs=true&CompanyState=0&CompanyType=1&Consolidatable=true&IndustryGroup=70&IsNotAudited=false&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&Publisher=false&ReportingType=1000002&TracingNo=-1&search=true";
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
                Accept: "application/json, text/plain, *!/!*",
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
                "application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, *!/!*",
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
    // اول همه نمادها رو بگیر
    const symbols = await fetchAllSymbols();

    console.log("📋 لیست چند نماد اول:");
    symbols.slice(0, 20).forEach((s, i) => {
        console.log(`${i + 1}. ${s.Symbol} - ${s.CompanyName}`);
    });

    // اینجا ادامه‌ی لاجیک قبلی تو برای گزارش‌ها

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
async function fetchAllSymbols(limitPages = 100) {
    let page = 1;
    let totalPages = limitPages; // پیش‌فرض تا ۱۰۰ صفحه
    const seen = new Set();
    const symbols = [];

    do {

        const url = `${BASE_SYMBOLS}&PageNumber=${page}`;
        console.log(`📡 Fetching symbols page ${page} ...`);
        const UA =
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36";
        const res = await fetch(url, {
            headers: {
                Accept: "application/json, text/plain, *!/!*",
                "User-Agent": UA,
                Referer: "https://www.codal.ir/",
                Origin: "https://www.codal.ir",
            },
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();

        if (page === 1) {
            // اگه کل صفحات کمتر از limitPages بود، به همون مقدار محدود بشه
            totalPages = Math.min(data.Page || 1, limitPages);
            console.log(`🧾 Total companies: ${data.Total} | Pages: ${data.Page}`);
            console.log(`⚡ فقط ${totalPages} صفحه اول برای تست خونده میشه`);
        }

        for (const l of data.Letters || []) {
            if (!seen.has(l.Symbol)) {
                seen.add(l.Symbol);
                symbols.push({ Symbol: l.Symbol, CompanyName: l.CompanyName });
            }
        }

        page++;
    } while (page <= totalPages);

    console.log(`✅ Fetched ${symbols.length} unique symbols (up to ${limitPages} pages)`);
    return symbols;
}

main().catch((err) => console.error("❌", err));
*/
const readline = require("readline");
const XLSX = require("xlsx");
const axios = require("axios");

/* ========== Utils ========== */
const UA =
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36";

function getNowFilename() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return `codal_sales_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(
        d.getDate()
    )}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}.xlsx`;
}

function fa2enDigits(s = "") {
    return s.replace(/[۰-۹]/g, (d) => "0123456789"["۰۱۲۳۴۵۶۷۸۹".indexOf(d)]);
}
function en2faDigits(s = "") {
    return s.replace(/[0-9]/g, (d) => "۰۱۲۳۴۵۶۷۸۹"[d]);
}
function normalizeNumber(x) {
    if (x === undefined || x === null) return null;
    let s = String(x);
    s = fa2enDigits(s).replace(/[,،\s]/g, "").replace(/[^\d.-]/g, "");
    if (!s || s === "-" || isNaN(Number(s))) return null;
    return Number(s);
}
function formatFaInt(n) {
    return en2faDigits(Math.trunc(n).toLocaleString("en-US"));
}
function cleanFa(s = "") {
    return String(s)
        .replace(/\u200c|‌/g, "")
        .replace(/ي/g, "ی")
        .replace(/ك/g, "ک")
        .trim();
}
function rowText(row) {
    return (row || []).map((c) => (c == null ? "" : String(c))).join(" ");
}

function extractDateFromTitle(title = "") {
    const m = String(title).match(/([۰-۹]{4}\/[۰-۹]{2}\/[۰-۹]{2}|\d{4}\/\d{2}\/\d{2})/);
    return m ? m[1] : "";
}
function periodLabelFromTitle(title = "") {
    const d = extractDateFromTitle(title);
    return d ? `دوره یک ماهه تا تاریخ ${d}` : "";
}

/* ========== Fetch reports ========== */
async function fetchCodalReports(symbol) {
    let page = 1,
        all = [],
        total = 0;
    while (true) {
        const url = `https://search.codal.ir/api/search/v2/q?Audited=true&AuditorRef=-1&Category=3&Childs=true&CompanyState=-1&CompanyType=-1&Consolidatable=true&IsNotAudited=false&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&PageNumber=${page}&Publisher=false&ReportingType=1000000&Symbol=${encodeURIComponent(
            symbol
        )}&TracingNo=-1&search=true`;
        const res = await axios.get(url, { headers: { Accept: "application/json, text/plain, */*", "User-Agent": UA } });
        const data = res.data || {};
        if (page === 1) total = data.Total || 0;
        if (Array.isArray(data.Letters)) all.push(...data.Letters);
        if (all.length >= total || (data.Page && page >= data.Page)) break;
        page++;
    }
    return all;
}

/* ========== Excel helpers ========== */
function isExactJamRow(row) {
    const first3 = [row[0], row[1], row[2]].map((x) => cleanFa(x || ""));
    return first3.some((v) => v === "جمع");
}
function looksLikeGroupRow(row) {
    const t = cleanFa(rowText(row));
    return (
        (t.includes("دوره") && t.includes("ماهه")) ||
        t.includes("ازابتدایسالمالی") ||
        t.includes("وضعیتمحصول-واحد")
    );
}
function findPeriodGroup(rows, targetDate) {
    const targetDateFa = cleanFa(targetDate);
    const targetDateEn = cleanFa(fa2enDigits(targetDate));
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const r = rows[i] || [];
        for (let c = 0; c < r.length; c++) {
            const cell = cleanFa(r[c] || "");
            if (cell.includes("دوره") && cell.includes("ماهه") &&
                (cell.includes(targetDateFa) || cell.includes(targetDateEn))) {
                return { groupRowIdx: i, groupColStart: c };
            }
        }
    }
    return { groupRowIdx: -1, groupColStart: -1 };
}
function findSalesIdxUnderGroup(rows, groupRowIdx, groupColStart) {
    for (let h = groupRowIdx + 1; h <= groupRowIdx + 6 && h < rows.length; h++) {
        const r = rows[h] || [];
        for (let c = groupColStart; c < groupColStart + 6 && c < r.length; c++) {
            const cell = cleanFa(r[c] || "");
            if (cell.includes("مبلغ") && cell.includes("فروش")) {
                return { headerIdx: h, salesIdx: c };
            }
        }
    }
    return { headerIdx: -1, salesIdx: -1 };
}

/* ========== Parse Excel ========== */
async function parseExcel(excelUrl, report, symbol) {
    const res = await axios.get(excelUrl, { responseType: "arraybuffer", headers: { "User-Agent": UA } });
    const wb = XLSX.read(res.data, { type: "buffer" });
    const targetDate = extractDateFromTitle(report.Title);

    for (const sh of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sh], { header: 1, raw: false });
        const { groupRowIdx, groupColStart } = findPeriodGroup(rows, targetDate);
        if (groupRowIdx === -1) continue;
        const { headerIdx, salesIdx } = findSalesIdxUnderGroup(rows, groupRowIdx, groupColStart);
        if (salesIdx === -1) continue;

        let lastSale = null;
        for (let j = headerIdx + 1; j < rows.length; j++) {
            const r = rows[j] || [];
            if (looksLikeGroupRow(r)) break;
            if (isExactJamRow(r)) {
                const v = normalizeNumber(r[salesIdx]);
                if (v !== null) lastSale = v;
            }
        }

        if (lastSale !== null) {
            return [
                {
                    Symbol: symbol,
                    Period: periodLabelFromTitle(report.Title),
                    SalesAmount: formatFaInt(lastSale),
                },
            ];
        }
    }
    return [];
}

/* ========== Process one symbol ========== */
async function processSymbol(symbol) {
    console.log(`\n--- بررسی نماد: ${symbol} ---`);
    const letters = await fetchCodalReports(symbol);
    const monthly = letters.filter((l) => String(l.Title || "").includes("گزارش فعالیت ماهانه"));
    console.log(`📑 ${symbol}: ${monthly.length} گزارش فعالیت ماهانه پیدا شد`);

    const bestByPeriod = new Map();
    for (const l of monthly) {
        const key = periodLabelFromTitle(l.Title);
        if (!key) continue;
        const ts = fa2enDigits(String(l.PublishDateTime || l.SentDateTime || "")).replace(/[^\d]/g, "");
        const prev = bestByPeriod.get(key);
        if (!prev || ts > prev._ts) bestByPeriod.set(key, { ...l, _ts: ts });
    }

    const results = [];
    for (const [, report] of bestByPeriod.entries()) {
        if (!report.ExcelUrl) continue;
        try {
            const rows = await parseExcel(report.ExcelUrl, report, symbol);
            results.push(...rows);
        } catch (e) {}
    }
    return results;
}

/* ========== Choose symbols ========== */
async function chooseSymbols() {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    const ask = (q) => new Promise((res) => rl.question(q, res));
    const opt = await ask("گزینه را انتخاب کنید (1: گرفتن از فایل symbols.xlsx / 2: نمونه تست): ");
    rl.close();

    if (opt === "1") {
        try {
            const wb = XLSX.readFile("symbols.xlsx");
            const sh = wb.Sheets[wb.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sh);
            const syms = data.map((r) => r.symbol || r.Symbol || r["نماد"]).filter(Boolean);
            console.log("📋 نمادهای خوانده‌شده:", syms);
            return syms;
        } catch (e) {
            console.error("❌ خطا در خواندن symbols.xlsx:", e.message);
            return [];
        }
    } else if (opt === "2") {
        return ["کورز"]; // تست
    } else {
        console.log("گزینه نامعتبر!");
        return [];
    }
}

/* ========== Main ========== */
(async function main() {
    const symbols = await chooseSymbols();
    if (!symbols.length) {
        console.log("⛔ نمادی انتخاب نشد");
        return;
    }

    let all = [];
    for (const s of symbols) {
        const rows = await processSymbol(s);
        all.push(...rows);
    }
    if (!all.length) {
        console.log("⛔ هیچ داده‌ای پیدا نشد");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(all);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reports");
    const out = getNowFilename();
    XLSX.writeFile(wb, out);
    console.log("✅ خروجی ذخیره شد:", out);
})();

