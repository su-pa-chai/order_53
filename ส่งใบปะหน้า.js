function my32() {
  var docFile = DriveApp.getFileById('13gSJKCLckWVXAnK387AkS1Q4x8qUrPyJ-FxhcGVQOVU')
  var TempFolder = DriveApp.getFolderById('1pFCPen0kwB7KV7NIlGyrNzKH6l6mwkPo')
  var PDFFolder = DriveApp.getFolderById('1KNl0-pXbWSWU6uBbxu1d8k-eYcnRnQlw') 
  var ss = SpreadsheetApp.openById('1fmpqAZlRAPCJFYeTxG1Igu8-6AU0osJcnL5xcgwXmzI')
  var sh = ss.getSheetByName('ส่ง')
  var data = sh.getRange(3,1,1,20).getValues()


/// ###,###,##0 ใส่ลูกน้ำให้ตัวเลข

      function numberWithCommas(number) {
      var parts = number.split(".");
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      return parts.join(".");

 /// ดึงข้อมูล     

  }
      data.forEach(r=>{
        var d_email = r[1]//อีเมล์
        var d_name = r[2]//ชื่อ-สกุล
        //var item3 = r[3]//เบี้ย
        var d_senddate = r[5]//วันที่ส่งงาน
        var d_lastday = r[6]//วันสุดท้าย
        var d_date = r[7]//วันที่  
        var d_admin = r[8]//จัดทำ
        var d_no = r[0]//ลำดับ
        var d_branch = r[9]//เรียน สาขา
        var d_pdf = r[10]//Pdf name
        var d_note = r[16]//หมายเหตุ
        var d_product = sh.getRange("J2").getValue()
        var d_temp = sh.getRange("I2").getValue()
        var data1 = numberWithCommas(sh.getRange("D3").getValue().toFixed(2));
        
        // Logger.log("Data1 = " + data1 + " Type : " + typeof(data1));
        var d_premium = data1;

// สร้าง PDF

CreatePDF(docFile,TempFolder,PDFFolder,d_email,d_name,d_senddate,d_lastday,d_date,d_admin,d_no,d_branch,d_pdf,d_note,d_premium,d_product,d_temp)

      })
}

function CreatePDF(docFile,TempFolder,PDFFolder,d_email,d_name,d_senddate,d_lastday,d_date,d_admin,d_no,d_branch,d_pdf,d_note,d_premium,d_product,d_temp){
var tempFile = docFile.makeCopy(TempFolder)
var tempDoc = DocumentApp.openById(tempFile.getId())
    tempDoc.getBody().replaceText("{ชื่อ สกุล}",d_name)
    tempDoc.getBody().replaceText("{เบี้ย}",d_premium)
    tempDoc.getBody().replaceText("{วันที่ส่งงาน}",d_senddate)
    tempDoc.getBody().replaceText("{วันสุดท้าย}",d_lastday)
    tempDoc.getBody().replaceText("{วันที่}",d_date)
    tempDoc.getBody().replaceText("{จัดทำ}",d_admin)
    tempDoc.getBody().replaceText("{ลำดับ}",d_no)
    tempDoc.getBody().replaceText("{หมายเหตุ}",d_note)
    tempDoc.getBody().replaceText("{แบบประกัน}",d_product)
    tempDoc.saveAndClose()


//  var namepdf = item8 + "." + item2;
//  Logger.log(namepdf + typeof(namepdf));
 var PdfContent = tempFile.getAs(MimeType.PDF)
 var PdfFile = PDFFolder.createFile(PdfContent).setName(d_pdf)
 var attachments = PdfFile.getAs(MimeType.PDF)
 var subject = 'เอกสารสำหรับแจ้งการส่งงาน'+d_product
// MailApp.sendEmail(d_email,'เอกสารสำหรับแจ้งการส่งงาน'+d_product,d_branch,{
// attachments: [PdfFile.getAs(MimeType.PDF)]
// })
// ลบไฟล์ทั้งหมดใน TempFolder
 var filesInTempFolder = TempFolder.getFiles();
 while (filesInTempFolder.hasNext()) {
  var file = filesInTempFolder.next();
  file.setTrashed(true);  // ให้ไฟล์ถูกลบไปที่ถังขยะ
}


sentEmails(d_email,d_branch,subject,attachments,d_admin,d_temp)
sendA()
}

 //add a menu when the spreadsheet is opened
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];  
  menuEntries.push({name: "ใบปะหน้า", functionName: "my32"}); 
  sheet.addMenu("👻ส่งเมล์", menuEntries);  
}

  function sendA() {
      SpreadsheetApp.getActive().toast("ส่งอีเมล์สำเร็จแล้ว '_' ");
  }


 function sentEmails(emailAddress,text,subject,attachments,d_admin,d_temp){
     var dear = "<font style='font-size: 14px; color: black;'>เรียน ท่านผู้จองสิทธิ์ " +" "+"</font><br>";
     message = "<p style='text-indent: 20px;'><font style='font-size: 14px; color: black;'>"+"   "+text +"<br>"
     message = message + "<font style='font-size: 14px; color: black;'>ขอแสดงความนับถือ"  +"</font><br>";
     message = message + "<font style='font-size: 14px; color: black;'>"+d_admin+"<br>"+d_temp+"</font><br>";//ชื่อ
    //  message = message + "<font style='font-size: 14px; color: black;'>ทีม"+row[num_name+1]+"</font>"; //ทีม
    //  message = message + "<font style='font-size: 14px; color: black;'>ฝ่าย"+row[num_name+2]+"</font>"; //ฝ่าย
    //  message = message + "<font style='font-size: 14px; color: black;'>โทร."+row[num_name+3]+"</font><br>"; //โทร

Logger.log(message);
    // var emailSent = emailAddress
    // Logger.log(emailSent);
    // if (emailSent != EMAIL_SENT) {
    //   var subject = row[4] 
  
var htmlBody = dear+ message ;
      MailApp.sendEmail({
        to: emailAddress,
        // cc: ccemail,
        subject: subject,
        body: "",
        htmlBody: htmlBody,
        attachments: [attachments]
      });
    }
  //  }