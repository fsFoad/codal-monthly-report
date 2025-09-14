// getSymbolsHtml.js
const axios = require("axios");
const cheerio = require("cheerio");
const XLSX = require("xlsx");

const UA =
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36";

const BASE_URL = "https://www.codal.ir/CompanyList.aspx";

/**
 * یک صفحه از لیست شرکت‌ها را بگیر
 */
async function fetchPage(page, cookies, viewState, eventValidation) {
    const body = new URLSearchParams({
        "ctl00$ScriptManager1":
            "ctl00$ContentPlaceHolder1$UpdatePanel1|ctl00$ContentPlaceHolder1$ucPager1$btnPage2",
        "ctl00$ContentPlaceHolder1$txbValue": "",
        "ctl00$ContentPlaceHolder1$ucPager1$hdfFromRowIndex": "0",
        "ctl00$ContentPlaceHolder1$ucPager1$hdfCurrentGroup": "1",
        "ctl00$ContentPlaceHolder1$ucPager1$hdfNavigatorIndex": "1",
        "ctl00$ContentPlaceHolder1$ucPager1$hdfActivePage": page.toString(),
        "ctl00$ContentPlaceHolder1$ucPager1$hdfSerial": "-1",
        "ctl00$ContentPlaceHolder1$ucPager1$hdfThumbPrint": "",
        "__EVENTTARGET": "",
        "__EVENTARGUMENT": "",
        "__VIEWSTATE": viewState,
        "__VIEWSTATEGENERATOR": "B825C6E2",
        "__VIEWSTATEENCRYPTED": "",
        "__EVENTVALIDATION": eventValidation,
        "__ASYNCPOST": "true",
        "ctl00$ContentPlaceHolder1$ucPager1$btnPage2": page.toString(),
    });

    const res = await axios.post(BASE_URL, body, {
        headers: {
            "User-Agent": UA,
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "X-MicrosoftAjax": "Delta=true",
            Referer: BASE_URL,
            Origin: "https://www.codal.ir",
            Cookie: cookies,
        },
    });

    return res.data;
}

/**
 * استخراج شرکت‌ها از HTML
 */
function parseCompanies(html) {
    const $ = cheerio.load(html);
    const rows = $("#ctl00_ContentPlaceHolder1_gvList tr").slice(1);

    const companies = [];
    rows.each((_, row) => {
        const tds = $(row).find("td");
        if (tds.length >= 3) {
            const symbol = $(tds[0]).text().trim();
            const name = $(tds[1]).text().trim();
            const isic = $(tds[2]).text().trim();

            companies.push({ symbol, name, isic });
        }
    });

    return companies;
}

/**
 * گرفتن همه شرکت‌ها (Scraping)
 */
async function getSymbolsHtml(limitPages = 2) {
    const init = await axios.get(BASE_URL, { headers: { "User-Agent": UA } });
    const cookies = init.headers["set-cookie"]
        .map((c) => c.split(";")[0])
        .join("; ");

    const $ = cheerio.load(init.data);
    const viewState = $("#__VIEWSTATE").val();
    const eventValidation = $("#__EVENTVALIDATION").val();

    let all = [];
    for (let page = 1; page <= limitPages; page++) {
        console.log(`📡 Fetching HTML page ${page}`);
        const html = await fetchPage(page, cookies, viewState, eventValidation);
        const companies = parseCompanies(html);
        console.log(`➡️ Page ${page}: ${companies.length} companies`);
        all = all.concat(companies);
    }

    console.log(`✅ Total companies: ${all.length}`);

    if (all.length) {
        const ws = XLSX.utils.json_to_sheet(all);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Symbols");
        XLSX.writeFile(wb, "symbols.xlsx");
        console.log("📊 Saved to symbols.xlsx");
    }

    return all;
}

module.exports = { getSymbolsHtml };
