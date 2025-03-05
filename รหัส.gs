var TEMPLATE_ID = "1FrMrPGfaerD-JopWtAc7k5cOudFQJLmkaTX39PenmQw";  // ใส่ Google Slides ID

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("S1");
  var data = sheet.getRange("F2:I400").getValues();
  
  return data.filter(row => row.some(cell => cell !== ""));
}

function generateCertificate(studentID, fullName, grade) {
  const templateId = "1FrMrPGfaerD-JopWtAc7k5cOudFQJLmkaTX39PenmQw";  // Google Slides Template ID
  const folderId = "1oDqFAh8tgmw4g-hLpmk2-ro_a1rt2oDX";  // Google Drive Folder ID

  // ดึงไฟล์ต้นฉบับ (Template)
  const templateFile = DriveApp.getFileById(templateId);
  const folder = DriveApp.getFolderById(folderId);

  // ทำสำเนาของเทมเพลตไปไว้ในโฟลเดอร์ที่กำหนด
  const slideCopy = templateFile.makeCopy(`${fullName}_เกียรติบัตร`, folder);
  const slide = SlidesApp.openById(slideCopy.getId());
  const slides = slide.getSlides()[0]; // เลือกหน้าแรกของสไลด์

  // แทนค่าข้อมูลในเทมเพลต
  slides.replaceAllText("{รหัสนักเรียน}", studentID);
  slides.replaceAllText("{ชื่อเต็ม}", fullName);
  slides.replaceAllText("{ชั้น}", grade);
  slide.saveAndClose();

  // แปลงเป็น PDF
  const pdfBlob = slideCopy.getAs(MimeType.PDF); // แปลง Google Slides เป็น PDF
  const pdfFile = folder.createFile(pdfBlob).setName(`เกียรติบัตร_${fullName}.pdf`);

  // ตั้งค่าสิทธิ์ให้สามารถเข้าถึงได้ทุกคนที่มีลิงก์
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // ลบไฟล์ Google Slides ที่สร้างขึ้น
  slideCopy.setTrashed(true);

  // ส่ง URL ของ PDF กลับไปที่ frontend
  return pdfFile.getUrl();
}


