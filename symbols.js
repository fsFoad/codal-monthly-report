// symbols.js
const fetch = (...args) =>
    import("node-fetch").then(({ default: fetch }) => fetch(...args));

const UA =
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36";
const BASE_SYMBOLS = "https://codal.ir/CompanyList.aspx";

async function fetchAllSymbols() {
    let page = 1;
    let totalPages = 1;
    const seen = new Set();
    const symbols = [];

    do {
        const url = `${BASE_SYMBOLS}&PageNumber=${page}`;
        console.log(`📡 Fetching symbols page ${page} ...`);

        const res = await fetch(url, {
            headers: {
                Accept: "application/json, text/plain, */*",
                "User-Agent": UA,
                Referer: "https://www.codal.ir/",
                Origin: "https://www.codal.ir",
            },
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();

        if (page === 1) {
            totalPages = data.Page || 1;
            console.log(`🧾 Total companies: ${data.Total} | Pages: ${totalPages}`);
        }

        for (const l of data.Letters || []) {
            if (!seen.has(l.Symbol)) {
                seen.add(l.Symbol);
                symbols.push({ Symbol: l.Symbol, CompanyName: l.CompanyName });
            }
        }

        page++;
    } while (page <= totalPages);

    console.log(`✅ Fetched ${symbols.length} unique symbols`);
    return symbols;
}

module.exports = { fetchAllSymbols };
