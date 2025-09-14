// fetchProductionSymbols.js
const axios = require("axios");
const XLSX = require("xlsx");

const UA =
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36";

const BASE_URL =
    "https://search.codal.ir/api/search/v2/q?PageSize=100&Childs=true&Mains=true&CompanyState=-1&Audited=true&NotAudited=true";

/**
 * گرفتن همه شرکت‌ها
 */
async function fetchAllCompanies(limitPages = 200) {
    let page = 1;
    let totalPages = limitPages;
    const companies = [];

    do {
        const url = `${BASE_URL}&PageNumber=${page}`;
        console.log(`📡 Fetching page ${page} ...`);

        const res = await axios.get(url, {
            headers: { "User-Agent": UA, Accept: "application/json" },
        });

        if (res.status !== 200) throw new Error(`HTTP ${res.status}`);
        const data = res.data;

        if (page === 1) {
            totalPages = Math.min(data.Page || 1, limitPages);
            console.log(
                `🧾 Total companies: ${data.Total} | Pages: ${totalPages}`
            );
            if (data.Letters?.length) {
                console.log("🔍 Sample company:", data.Letters[0]);
            }
        }

        if (Array.isArray(data.Letters)) {
            companies.push(
                ...data.Letters.map((l) => ({
                    Symbol: l.Symbol,
                    CompanyName: l.CompanyName,
                    IndustryGroup: l.IndustryGroup || "",
                }))
            );
        }

        page++;
    } while (page <= totalPages);

    return companies;
}

/**
 * فیلتر فقط تولیدی‌ها
 */
async function fetchProductionSymbols() {
    const all = await fetchAllCompanies(200);

    // 🔍 شرط: IndustryGroup شامل "تولید"
    const filtered = all.filter((c) =>
        (c.IndustryGroup || "").includes("تولید")
    );

    console.log(`✅ Found ${filtered.length} تولیدی companies`);

    const data = filtered.map((f) => ({
        symbol: f.Symbol,
        name: f.CompanyName,
        industry: f.IndustryGroup,
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Symbols");

    try {
        XLSX.writeFile(wb, "symbols.xlsx");
        console.log("📊 Saved to symbols.xlsx");
    } catch (err) {
        console.error("❌ Error writing Excel:", err.message);
    }

    return filtered;
}

module.exports = { fetchProductionSymbols };
