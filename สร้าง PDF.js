/*
function myFunction() {
 var docFile =  DriveApp.getFileById('')
 var TempFolder = DriveApp.getFileById('')
 var PDFFoder = DriveApp.getFileById('1rJzzF13yvG-bTp4uzHmNnwb5PfkAPY3C')
 var ss = SpreadsheetApp.openById('1KXAEajI1Nrw-v-lDAGUD1KY2RnC6fUxNbm4SdM8tU10')
 var sh = ss.getSheetByName('Booking')
 var data = sh.getRange(ss.getLastColumn(),1,1,30).getValues()
     data.forEach(r=>{
       var item1 = r[1] // email
       var item2 = r[2] // คำนำหน้า
       var item3 = r[3] // ชื่อ
       var item4 = r[4] // สกุล
       var item5 = r[5] // วันที่จอง
       var item6 = r[6] // จำทำ
       var item6 = r[7] // ลำดับ
     })
}
function CreateDDF (docFile,TempFoder,PDFFolder,item1,item2,item3,item4,item5,item6){
  var tempFile = docFile.makeCopy(TempFoder)
  var tempDoc = SlidesApp.openById(tempFile.getId)
      tempDoc.getBody().replaceText("{คำนำหน้า}",item2)
      tempDoc.getBody().replaceText("{ชื่อ}",item3)
      tempDoc.getBody().replaceText("{สกุล}",item4)
      tempDoc.getBody().replaceText("{วันที่จอง}",item5)
      tempDoc.getBody().replaceText("{จัดทำ}",item6)
      tempDoc.getBody().replaceText("{ลำดับ}",item7)

  var PdfContent = tempFile.getAs(MimeType.PDF)
  var PdfFile = PDFFoder.CreateDDF(PdfContent).setName(item2)

  MailApp.sendEmail(item1,'ส่ง PDF','teat mail',{
    attachments: [PdfFile.getAs(MimeType.PDF)]

  })

}

*/