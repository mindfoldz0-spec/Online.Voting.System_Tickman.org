/**
 * Export all GovernmentRecord rows (the "government database" used for
 * voter registration) to an Excel file in the project root.
 *
 * Run with: npm run export:records
 * or:       npx tsx scripts/export-records-xlsx.ts
 */
import { PrismaClient } from '@prisma/client';
import ExcelJS from 'exceljs';
import path from 'path';
import * as dotenv from 'dotenv';

dotenv.config();

const prisma = new PrismaClient();

async function main() {
    console.log('🔄 Fetching government records from database...');

    const records = await prisma.governmentRecord.findMany({
        orderBy: { documentNumber: 'asc' },
    });

    console.log(`✅ Found ${records.length} records.`);

    if (records.length === 0) {
        console.log('⚠️ No records found to export. Run `npx prisma db seed` first.');
        return;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'VoteSecure Export';
    workbook.created = new Date();

    const sheet = workbook.addWorksheet('Government Records');
    sheet.columns = [
        { header: 'ID', key: 'id', width: 12 },
        { header: 'Full Name', key: 'fullName', width: 26 },
        { header: 'Date of Birth', key: 'dateOfBirth', width: 15 },
        { header: 'Document Number (Aadhaar)', key: 'documentNumber', width: 28 },
        { header: 'Address', key: 'address', width: 42 },
        { header: 'Photo URL', key: 'photoUrl', width: 28 },
        { header: 'Created At', key: 'createdAt', width: 26 },
    ];

    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9EAD3' },
    };

    for (const r of records) {
        sheet.addRow({
            id: r.id,
            fullName: r.fullName,
            dateOfBirth: r.dateOfBirth,
            documentNumber: r.documentNumber,
            address: r.address,
            photoUrl: r.photoUrl,
            createdAt: r.createdAt.toISOString(),
        });
    }

    // Freeze the header row for easier browsing
    sheet.views = [{ state: 'frozen', ySplit: 1 }];

    const outputPath = path.join(process.cwd(), 'Government_Records.xlsx');
    await workbook.xlsx.writeFile(outputPath);

    console.log(`📂 Exported ${records.length} records to: ${outputPath}`);
}

main()
    .catch((error) => {
        console.error('❌ Export failed:', error);
        process.exit(1);
    })
    .finally(async () => {
        await prisma.$disconnect();
    });
