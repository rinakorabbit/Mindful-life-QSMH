/**
 * โครงการกิจกรรมการพัฒนาจิตเพื่อคุณภาพชีวิต (Mindful Life QSMH)
 * งานกิจกรรมวิชาการและการจัดการความรู้ ฝ่ายวิชาการและวิจัย
 * โรงพยาบาลสมเด็จพระบรมราชเทวี ณ ศรีราชา สภากาชาดไทย
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Mindful Life QSMH - กิจกรรมการพัฒนาจิตเพื่อคุณภาพชีวิต')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// แผ่นงานบันทึกข้อมูล
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Registrations');
  if (!sheet) {
    sheet = ss.insertSheet('Registrations');
    sheet.appendRow([
      'รุ่น/รอบที่สมัคร',
      'วัน-เวลาที่ลงทะเบียน',
      'ชื่อ-นามสกุล',
      'อีเมล',
      'เบอร์โทรศัพท์',
      'ประเภทบุคลากร',
      'ฝ่าย/หน่วยงานต้นสังกัด',
      'ตำแหน่งงาน'
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#E2E8F0');
  }
  return sheet;
}

// แผ่นงานตั้งค่ารอบกิจกรรม (Config)
function getEventConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfgSheet = ss.getSheetByName('Config');
  
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet('Config');
    const defaultData = [
      ['Key', 'Value'],
      ['Active_Batch', 'รุ่นที่ 1'],
      ['Event_Date', 'วันศุกร์ที่ 27 มีนาคม 2570'],
      ['Event_Time', '08.30 - 16.00 น.'],
      ['Event_Place', 'ห้องประชุม ชั้น 4 อาคารอนุสรณ์ ๑๐๐ ปี'],
      ['Max_Quota', 100],
      ['Is_Open', 'YES']
    ];
    cfgSheet.getRange(1, 1, defaultData.length, 2).setValues(defaultData);
    cfgSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#E2E8F0');
  }

  const values = cfgSheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < values.length; i++) {
    config[values[i][0]] = values[i][1];
  }

  // นับจำนวนผู้สมัครปัจจุบันของรุ่นนี้
  const regSheet = getSheet();
  const lastRow = regSheet.getLastRow();
  let currentRegisteredCount = 0;
  
  if (lastRow > 1) {
    const batchValues = regSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    currentRegisteredCount = batchValues.filter(b => String(b).trim() === String(config['Active_Batch']).trim()).length;
  }
  
  config['Current_Count'] = currentRegisteredCount;
  return config;
}

// บันทึกข้อมูลการลงทะเบียน
function submitRegistration(data) {
  try {
    const config = getEventConfig();
    
    if (String(config['Is_Open']).toUpperCase() !== 'YES') {
      return { success: false, message: 'ขออภัย ระบบปิดรับสมัครสำหรับรอบนี้แล้ว' };
    }
    
    if (Number(config['Current_Count']) >= Number(config['Max_Quota'])) {
      return { success: false, message: 'ขออภัย จำนวนผู้สมัครรอบนี้เต็มครบตามโควตาแล้ว' };
    }

    const sheet = getSheet();
    sheet.appendRow([
      config['Active_Batch'],
      new Date(),
      data.fullname,
      data.email,
      "'" + data.phone,
      data.userType,
      data.department || '-',
      data.position || '-'
    ]);

    // ส่งอีเมลตอบกลับ
    sendConfirmationEmail(data, config);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ส่งอีเมลยืนยันผลการลงทะเบียน
function sendConfirmationEmail(data, config) {
  const subject = `ยืนยันการลงทะเบียน: กิจกรรมการพัฒนาจิตเพื่อคุณภาพชีวิต (${config['Active_Batch']})`;
  const htmlBody = `
    <div style="font-family: 'Sarabun', Arial, sans-serif; line-height: 1.8; color: #2D3748; max-width: 620px; margin: auto; border: 1px solid #E2E8F0; border-radius: 14px; padding: 26px; background-color: #FFFFFF;">
      <h2 style="color: #166534; text-align: center; margin-top: 0;">ยืนยันการลงทะเบียนสำเร็จ</h2>
      <p style="font-size: 16px;">เรียนคุณ <b>${data.fullname}</b>,</p>
      <p style="font-size: 16px;">ท่านได้ลงทะเบียนเข้าร่วม <b>กิจกรรมการพัฒนาจิตเพื่อคุณภาพชีวิต (${config['Active_Batch']})</b> เรียบร้อยแล้ว โดยมีกำหนดการดังนี้</p>
      
      <div style="background-color: #F0FDF4; border: 1px solid #BBF7D0; padding: 18px; border-radius: 10px; margin: 18px 0; font-size: 16px;">
        <p style="margin: 6px 0; color: #166534;"><b>📅 วันที่จัดกิจกรรม:</b> ${config['Event_Date']}</p>
        <p style="margin: 6px 0; color: #166534;"><b>⏰ เวลา:</b> ${config['Event_Time']}</p>
        <p style="margin: 6px 0; color: #166534;"><b>📍 สถานที่:</b> ${config['Event_Place']}</p>
      </div>

      <div style="background-color: #F8FAFC; padding: 16px; border-radius: 10px; margin: 18px 0; font-size: 15px; border-left: 4px solid #16a34a;">
        <p style="margin: 4px 0;"><b>ประเภท:</b> ${data.userType}</p>
        ${data.department ? `<p style="margin: 4px 0;"><b>ฝ่าย/หน่วยงาน:</b> ${data.department}</p>` : ''}
        ${data.position ? `<p style="margin: 4px 0;"><b>ตำแหน่ง:</b> ${data.position}</p>` : ''}
        <p style="margin: 4px 0;"><b>เบอร์โทรศัพท์:</b> ${data.phone}</p>
      </div>

      <hr style="border: none; border-top: 1px solid #E2E8F0; margin: 24px 0;">
      
      <p style="font-size: 15px; color: #4A5568; margin-bottom: 6px;"><b>ติดต่อสอบถามข้อมูลเพิ่มเติมได้ที่:</b></p>
      <p style="margin: 2px 0; color: #4A5568; font-size: 14px;">งานกิจกรรมวิชาการและการจัดการความรู้ ฝ่ายวิชาการและวิจัย</p>
      <p style="margin: 2px 0; color: #4A5568; font-size: 14px;">ชั้น 4 อาคารอนุสรณ์ ๑๐๐ ปี โรงพยาบาลสมเด็จพระบรมราชเทวี ณ ศรีราชา สภากาชาดไทย</p>
      <p style="margin: 4px 0; color: #166534; font-size: 14px;"><b>📞 เบอร์โทรศัพท์:</b> 038-320-200 ต่อ 23458</p>
    </div>
  `;

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    htmlBody: htmlBody
  });
}

// ดึงรายชื่อเฉพาะรอบปัจจุบัน เรียง ก-ฮ
function getRegisteredList() {
  const config = getEventConfig();
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return { batch: config['Active_Batch'], date: config['Event_Date'], list: [] };
  }

  // Col A = Batch, Col C = Fullname
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const currentBatchNames = data
    .filter(row => String(row[0]).trim() === String(config['Active_Batch']).trim() && String(row[2]).trim() !== '')
    .map(row => String(row[2]).trim());

  currentBatchNames.sort((a, b) => a.localeCompare(b, 'th'));
  return {
    batch: config['Active_Batch'],
    date: config['Event_Date'],
    list: currentBatchNames
  };
}
