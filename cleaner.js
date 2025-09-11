const XLSX = require("xlsx");

// مپ ماه‌ها
const monthMap = {
    "01": "فروردین",
    "02": "اردیبهشت",
    "03": "خرداد",
    "04": "تیر",
    "05": "مرداد",
    "06": "شهریور",
    "07": "مهر",
    "08": "آبان",
    "09": "آذر",
    "10": "دی",
    "11": "بهمن",
    "12": "اسفند",
};

// تبدیل تاریخ به ماه فارسی
function formatDate(dateStr) {
    if (!dateStr || typeof dateStr !== "string") return dateStr;
    const parts = dateStr.split("/");
    if (parts.length < 2) return dateStr;
    const year = parts[0];
    const month = parts[1].padStart(2, "0");
    return `${monthMap[month] || month} ${year}`;
}

// حذف ردیف‌های اضافی
function cleanFinancialRows(rows) {
    return rows.filter((row) => {
        const firstCell = row[0] ? row[0].toString().trim() : "";

        // نگه داشتن سطرهای "سرمایه" یا "جمع"
        if (firstCell.includes("سرمایه") || firstCell.includes("جمع")) {
            return true;
        }

        // حذف ردیف‌هایی که فقط عدد تکی دارند
        if (row.length === 1 && /^\d/.test(firstCell)) {
            return false;
        }

        return true;
    });
}

// ذخیره اکسل
function saveToExcel(tables, filename = "output.xlsx") {
    const workbook = XLSX.utils.book_new();

    tables.forEach((table, idx) => {
        let cleaned = cleanFinancialRows(table);

        // تبدیل تاریخ‌ها
        cleaned = cleaned.map((row) =>
            row.map((cell) => {
                if (typeof cell === "string" && cell.match(/^\d{4}\/\d{2}\/\d{2}$/)) {
                    return formatDate(cell);
                }
                return cell;
            })
        );

        const ws = XLSX.utils.aoa_to_sheet(cleaned);
        XLSX.utils.book_append_sheet(workbook, ws, `جدول${idx + 1}`);
    });

    XLSX.writeFile(workbook, filename, { bookType: "xlsx" });
    console.log(`✅ فایل اکسل ذخیره شد: ${filename}`);
}

module.exports = {
    formatDate,
    cleanFinancialRows,
    saveToExcel,
};