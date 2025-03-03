import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
export async function POST(request) {
    try {
        const formData = await request.formData();
        const file = formData.get("file");
        if (!file) {
            return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
        }
        const buffer = await file.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet(1);
        const sheetData = {
            name: worksheet.name,
            color: "",
            config: {
                merge: {},
                rowlen: {},
                columnlen: {},
            },
            celldata: [],
        };
        if (worksheet._merges) {
            Object.keys(worksheet._merges).forEach((mergeKey) => {
                const mergeRange = worksheet._merges[mergeKey];
                const key = `${mergeRange.top}_${mergeRange.left}`;
                sheetData.config.merge[key] = {
                    r: mergeRange.top - 1,
                    c: mergeRange.left - 1,
                    rs: mergeRange.bottom - mergeRange.top + 1,
                    cs: mergeRange.right - mergeRange.left + 1,
                };
            });
        }
        worksheet.columns.forEach((column, index) => {
            if (column.width) {
                sheetData.config.columnlen[index] = Math.round(column.width * 7.5);
            }
        });
        worksheet.eachRow((row, rowNumber) => {
            if (row.height) {
                sheetData.config.rowlen[rowNumber - 1] = Math.round(row.height);
            }
        });
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                const cellData = {
                    r: rowNumber - 1,
                    c: colNumber - 1,
                    v: {},
                };
                cellData.v.v = cell.text || cell.value;
                if (
                    cell.fill &&
                    cell.fill.type === "pattern" &&
                    cell.fill.pattern === "solid"
                ) {
                    cellData.v.bg =
                        rgbToHex(cell.fill.fgColor) || rgbToHex(cell.fill.bgColor);
                }
                if (cell.font) {
                    cellData.v.bl = cell.font.bold ? 1 : 0;
                    cellData.v.it = cell.font.italic ? 1 : 0;
                    cellData.v.ff = cell.font.name || 0;
                    cellData.v.fs = cell.font.size || 10;
                    cellData.v.fc = rgbToHex(cell.font.color) || "#000000";
                    cellData.v.ul = cell.font.underline ? 1 : 0;
                }
                if (cell.alignment) {
                    cellData.v.vt = getVerticalAlignment(cell.alignment.vertical);
                    cellData.v.ht = getHorizontalAlignment(cell.alignment.horizontal);
                }
                if (cell.numFmt) {
                    cellData.v.fm = cell.numFmt;
                }
                sheetData.celldata.push(cellData);
            });
        });
        return NextResponse.json(sheetData);
    } catch (error) {
        console.error("Error reading Excel file:", error);
        return NextResponse.json(
            { error: "Failed to read Excel file", details: error.message },
            { status: 500 }
        );
    }
}
function rgbToHex(color) {
    if (!color || !color.rgb) return null;
    let hex = color.rgb.toString(16);
    return "#" + hex.padStart(6, "0");
}
function getVerticalAlignment(alignment) {
    switch (alignment) {
        case "top":
            return 1;
        case "middle":
            return 0;
        case "bottom":
            return 2;
        default:
            return 0;
    }
}
function getHorizontalAlignment(alignment) {
    switch (alignment) {
        case "left":
            return 1;
        case "center":
            return 0;
        case "right":
            return 2;
        default:
            return 1;
    }
}
export const config = {
    api: {
        bodyParser: false,
    },
};





