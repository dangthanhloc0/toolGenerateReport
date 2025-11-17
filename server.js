const express = require("express");
const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const cors = require("cors");
const { title } = require("process");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const path = require("path");
const archiver = require("archiver");


const app = express();
app.use(cors({
  exposedHeaders: ['Content-Disposition']
}));
app.use(express.json({ limit: "5mb" }));
app.use(express.static(__dirname));

app.post("/crawl", async (req, res) => {
  try {
    const { url,urlChangeLog } = req.body;
    if (!url) return res.status(400).send("Missing url");
    
    // Sử dụng chromium thông thường cho production
    const browser = await chromium.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 30000 });

 
    const data = await page.evaluate(() => ({
      ticketNumber: document.querySelector(".issue-link")?.firstChild?.textContent || "",
      title : document.getElementById("summary-val")?.querySelector("h2")?.innerText || "",
      description : document.querySelector(".user-content-block")?.innerText || "",
      today : new Date().toLocaleDateString(),
      shortname : document.querySelector('meta[name="ajs-remote-user"]')?.getAttribute("content") || "",
      name : document.querySelector('meta[name="ajs-remote-user-fullname"]')?.getAttribute("content") || ""
    }));
    
    // Format today thành ddMMyyyy (17112025)
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const todayFormatted = `${day}${month}${year}`;
    
    // Kiểm tra nếu title có "PCUT xxxx" thì extract ra
    let pcutString = "";
    const pcutMatch = data.title.match(/PCUT[\s-]*\d+/i);
    if (pcutMatch) {
      pcutString = pcutMatch[0];
      console.log("Found PCUT:", pcutString);
    }
    
    console.log(data);

    console.log("Raw urlChangeLog:", urlChangeLog);
    var listUrlChangeLog = [];
    if (urlChangeLog && urlChangeLog.length > 0) {
        listUrlChangeLog = urlChangeLog.split(",").map(url => url.trim()).filter(url => url !== "");
        console.log(`Processing ${listUrlChangeLog.length} change log URLs:`, listUrlChangeLog);
    }
    const ELList = [];
    for (let itemUrl of listUrlChangeLog) {
        const page2 = await browser.newPage();
        console.log("Processing change log URL:", itemUrl);
        itemUrl+='/files';
        await page2.goto(itemUrl.trim(), { waitUntil: "domcontentloaded", timeout: 60000 });
 

        const rootFolder = await page2.evaluate(() => {
            const element = document.querySelector(".Truncate-text");
            return element?.textContent?.trim() || "";
        });
        
        const dataChangeLog = await page2.evaluate(({ folder, jiraData }) => {
            const elements = Array.from(document.querySelectorAll(".Truncate-text"));
            const fromFourth = elements.slice(3);
            
            console.log(`Total elements: ${elements.length}, From 4th: ${fromFourth.length}`);
            
            return fromFourth.map(el => {
                const title = el.getAttribute("title") || el.innerText || "";
                const fullPath = folder ? folder + '/' + title : title;
                const parts = fullPath.split('/');
                const filename = parts[parts.length - 1];
                const filenameParts = filename.split('.');
                const fileType = filenameParts.length > 1 ? filenameParts[filenameParts.length - 1] : 'unknown';
                
                return {
                    fullPath: fullPath,
                    filename: filename,
                    fileType: fileType,
                    JIRATicket: jiraData.ticketNumber,
                    Owner: jiraData.shortname,
                    StartDate: jiraData.today,
                    EndDate: jiraData.today,
                    UpdatedDate: jiraData.today
                };
            });
        }, { folder: rootFolder, jiraData: data });
            
        console.log(`Found ${dataChangeLog.length} files from ${itemUrl}`);
        console.log('dataChangeLog content:', JSON.stringify(dataChangeLog, null, 2));
        
        if (dataChangeLog.length > 0) {
            // Filter out empty items
            const validItems = dataChangeLog.filter(item => item.filename && item.filename.trim() !== '');
            console.log(`Valid items: ${validItems.length}`);
            ELList.push(...validItems);
        } else {
            console.log(`Warning: No files found for ${itemUrl}`);
        }
    }

    await browser.close();

    const log = path.join(__dirname, "template", "logSourceChangeTemplate.xlsx");

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(log);
    
    // Ghi vào sheet Cover
    const CoverSheet = workbook.getWorksheet("Cover");
    CoverSheet.getCell("E15").value = data.today;
    CoverSheet.getCell("F15").value = data.shortname;
    CoverSheet.getCell("H15").value = data.ticketNumber;
    CoverSheet.getCell("I15").value = data.title;
    
    // Ghi vào sheet EL từ dòng 3
    const ELSheet = workbook.getWorksheet("EL");
    let rowIndex = 3;
    let count  = 1;
    
    // Define border style
    const borderStyle = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    };
    
    ELList.forEach(item => {
      const columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M','N'];
      
      ELSheet.getCell(`A${rowIndex}`).value = count++;
      ELSheet.getCell(`B${rowIndex}`).value ='EL P&C';
      ELSheet.getCell(`E${rowIndex}`).value ='Modify';
      ELSheet.getCell(`F${rowIndex}`).value ='Modify';
      ELSheet.getCell(`G${rowIndex}`).value ='Modify';
      ELSheet.getCell(`M${rowIndex}`).value = item.fullPath;
      ELSheet.getCell(`C${rowIndex}`).value = item.filename;
      ELSheet.getCell(`D${rowIndex}`).value = item.fileType;
      ELSheet.getCell(`H${rowIndex}`).value = item.JIRATicket;
      ELSheet.getCell(`I${rowIndex}`).value = item.Owner;
      ELSheet.getCell(`J${rowIndex}`).value = item.StartDate;
      ELSheet.getCell(`K${rowIndex}`).value = item.EndDate;
      ELSheet.getCell(`L${rowIndex}`).value = item.UpdatedDate;
      ELSheet.getCell(`N${rowIndex}`).value = '';
      
      // Apply border to all cells in the row
      columns.forEach(col => {
        ELSheet.getCell(`${col}${rowIndex}`).border = borderStyle;
      });
      
      rowIndex++;
    });

    const GIT_DiffSheet = workbook.getWorksheet("GIT_Diff");
    rowIndex = 2;
    count  = 1;
    listUrlChangeLog.forEach(item => {
      GIT_DiffSheet.getCell(`A${rowIndex}`).value = count++;
      
      // Tạo hyperlink cho URL với màu xanh
      const cellB = GIT_DiffSheet.getCell(`B${rowIndex}`);
      cellB.value = {
        text: item,
        hyperlink: item,
        tooltip: 'Click to open'
      };
      cellB.font = {
        color: { argb: 'FF0000FF' }, // Màu xanh
        underline: true
      };
      
      // Apply border to GIT_Diff cells
      ['A', 'B'].forEach(col => {
        GIT_DiffSheet.getCell(`${col}${rowIndex}`).border = borderStyle;
      });
      
      rowIndex++;
    }   );

    // Tạo ZIP file thay vì trả về 1 file Excel
    const zipFilename = `${data.ticketNumber}.zip`;
    
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(zipFilename)}`);
    
    const archive = archiver('zip', {
      zlib: { level: 9 } // Compression level
    });
    
    // Pipe archive to response
    archive.pipe(res);
    
    // Thêm file Excel vào ZIP
    const filename = `${data.ticketNumber} - Source Change Log v1.0.xlsx`;
    const buffer = await workbook.xlsx.writeBuffer();
    archive.append(buffer, { name: filename });
    
    // Thêm file Excel thứ 2 (ví dụ: copy của template hoặc file khác)
    const Checklist  = path.join(__dirname, "template", "Checklist.xlsx");

    const workbook1 = new ExcelJS.Workbook();
    await workbook1.xlsx.readFile(Checklist );
    const buffer1 = await workbook1.xlsx.writeBuffer();
    archive.append(buffer1, { name: `${data.ticketNumber} - Integral Java Checklist v1.0 (DEV).xlsx` });

    // Thêm file Word SonarLint_Report_Find_Bugs
    const SonarLint_Report_Find_Bugs = path.join(__dirname, "template", "SonarLint_Report_Find_Bugs.docx");
    if (fs.existsSync(SonarLint_Report_Find_Bugs)) {
      const wordContent = fs.readFileSync(SonarLint_Report_Find_Bugs, "binary");
      const zip = new PizZip(wordContent);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
      doc.render();
      const wordBuffer = doc.getZip().generate({ type: "nodebuffer" });
      archive.append(wordBuffer, { name: `${data.ticketNumber} - SonarLint_Report_Find_Bugs v1.0.docx` });
    }


    // UT
    const UT  = path.join(__dirname, "template", "UT.xlsx");

    const workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.readFile(UT);

    const UTS = workbook2.getWorksheet("UTS");
    UTS.getCell("E3").value = data.ticketNumber + " " + data.title;
    UTS.getCell("E5").value = data.description;

    const buffer2 = await workbook2.xlsx.writeBuffer();
    archive.append(buffer2, { name: `${data.ticketNumber} - UT v1.0.xlsx` });


    // Thêm file Word MSIGID-RELEASE  
    const MSIGIDRELEASE = path.join(__dirname, "template", "MSIGID-RELEASE.docx");
    if (fs.existsSync(MSIGIDRELEASE)) {
      const wordContent = fs.readFileSync(MSIGIDRELEASE, "binary");
      const zip = new PizZip(wordContent);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
      
      // Fill data vào Word template
      doc.setData({
        day: day,
        month: month,
        name: data.name,
        today: data.today,
        ticket: data.ticketNumber,
        title: data.title,
        pcut: pcutString
      });
      
      doc.render();
      const wordBuffer = doc.getZip().generate({ type: "nodebuffer" });
      archive.append(wordBuffer, { name: `MSIGID-RELEASE REQUEST_${todayFormatted}_${data.ticketNumber}.docx` });
    }
    
    // Finalize ZIP
    await archive.finalize();
    
  } catch (err) {
    console.error("Crawl error:", err);
    res.status(500).send("Server error: " + (err.message || err));
  }
});


async function generateExcel(data) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("template.xlsx"); // đọc file mẫu
  const CoverSheet = workbook.getWorksheet("Cover");
  CoverSheet.getCell("E15").value = data.today;
  CoverSheet.getCell("F15").value = data.shortname;
  CoverSheet.getCell("H15").value = data.ticketNumber;
  CoverSheet.getCell("I15").value = data.title;
  

  await workbook.xlsx.writeFile("result.xlsx"); // xuất file mới
}

function generateWord(data) {
  const content = fs.readFileSync("template.docx", "binary");
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

  doc.setData({
    ticketNumber: data.ticketNumber,
    title: data.title,
    description: data.description,
    today: data.today,
  });

  try {
    doc.render();
  } catch (error) {
    console.error(error);
  }

  const buf = doc.getZip().generate({ type: "nodebuffer" });
  fs.writeFileSync("result.docx", buf);
}

const PORT = process.env.PORT || 3001;
app.listen(PORT, '0.0.0.0', () => console.log(`Backend chạy tại http://localhost:${PORT}`));
