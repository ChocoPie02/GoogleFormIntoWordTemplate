# GoogleFormIntoWordTemplate
Auto Fill Google Form Response into Word doc Template

i know this script still have a lot room of improvement, then i really happy to hear about your advice

## Script :

~~~javascript
function autoFillGoogleDocFromForm(e) {

  //Function to insert image into replaceable text
   var replaceTextToImage = function(body, searchText, image, width) {
    var next = body.findText(searchText);
    if (!next) return;
    var r = next.getElement();
    r.asText().setText("");
    var img = r.getParent().asParagraph().insertInlineImage(0, image);
    if (width && typeof width == "number") {
      var w = img.getWidth();
      var h = img.getHeight();
      img.setWidth(width);
      img.setHeight(width * h / w);
    }
    return next;
  };
  //e.values is an array of form values
  //informative variable
  var timestamp = e.values[0];
  var departemen = e.values[1];
  var divisi_GA = e.values[2];
  var divisi_finance = e.values[3];
  var divisi_operasional = e.values[4];
  var divisi_warehouse = e.values[5];

  //variable of main data
  var kegiatan_1 = e.values[6];
  var foto_kegiatan_1 = e.values[7];
  var kegiatan_2 = e.values[9];
  var foto_kegiatan_2 = e.values[10];
  var kegiatan_3 = e.values[12];
  var foto_kegiatan_3 = e.values[13];
  var kegiatan_4 = e.values[15];
  var foto_kegiatan_4 = e.values[16];
  var kegiatan_5 = e.values[18];
  var foto_kegiatan_5 = e.values[19];
  var kegiatan_6 = e.values[21];
  var foto_kegiatan_6 = e.values[22];
  
  //getting ID of image uploaded and prepare as blob (image)
  //use try catch to prevent error when user did not upload image 
  //at this time, only 1 image can uploaded, still try make a way to multi image at once
  try{
    var link1 = DriveApp.getFileById(foto_kegiatan_1.split("https://drive.google.com/open?id=")[1]).getBlob();
  }catch{}
    try{
    var link2 = DriveApp.getFileById(foto_kegiatan_2.split("https://drive.google.com/open?id=")[1]).getBlob();
  }catch{}
  try{
    var link3 = DriveApp.getFileById(foto_kegiatan_3.split("https://drive.google.com/open?id=")[1]).getBlob();
  }catch{}
  try{
    var link4 = DriveApp.getFileById(foto_kegiatan_4.split("https://drive.google.com/open?id=")[1]).getBlob();
  }catch{}
  try{
    var link5 = DriveApp.getFileById(foto_kegiatan_5.split("https://drive.google.com/open?id=")[1]).getBlob();
  }catch{}
  try{
    var link6 = DriveApp.getFileById(foto_kegiatan_6.split("https://drive.google.com/open?id=")[1]).getBlob();
  }catch{}
  
  //normal var date on specific timezone
  var date = Utilities.formatDate(new Date(), "GMT+7", "dd/MMM/yyyy")
  //just add all divisi into one place
  var divisi = divisi_GA+divisi_finance+divisi_operasional+divisi_warehouse
  
  //file is the template file, and you get it by ID
  var file = DriveApp.getFileById('1D9IqR8qu0oWV-O0nfltP63m9H3EMoI5170SI8RcOCb46'); 
  
  //We can make a copy of the template, name it, and optionally tell it what folder to live in
  //file.makeCopy will return a Google Drive file object
  var folder = DriveApp.getFolderById('1c8CXfufIU7mUBRlSxhNZA8XtM6ic0FA0q')
  var copy = file.makeCopy('Daily Report '+departemen+ ' / ' +divisi+' ~ Tgl '+ date, folder); 
  
  //Once we've got the new file created, we need to open it as a document by using its ID
  var doc = DocumentApp.openById(copy.getId()); 
  
  //Since everything we need to change is in the body, we need to get that
  var body = doc.getBody(); 
  
  //Then we call all of our replaceText methods
  body.replaceText('<<Departemen>>', departemen); 
  body.replaceText('<<Divisi>>', divisi);  
  body.replaceText('<<Today>>', date); 
  body.replaceText('<<Deskripsi Kegiatan Pertama>>', kegiatan_1);  
  body.replaceText('<<Deskripsi Kegiatan Kedua>>', kegiatan_2); 
  body.replaceText('<<Deskripsi Kegiatan Ketiga>>', kegiatan_3);
  body.replaceText('<<Deskripsi Kegiatan Keempat>>', kegiatan_4); 
  body.replaceText('<<Deskripsi Kegiatan Kelima>>', kegiatan_5);
  body.replaceText('<<Deskripsi Kegiatan Keenam>>', kegiatan_6); 

  //try catch to prevent error when there is no image, and replace the Image Tag <<image>> into blank space
  try{
    do {
    var next = replaceTextToImage(body,'<<Foto kegiatan Pertama>>', link1, 180);
    } while (next);
  }catch{
    body.replaceText('<<Foto kegiatan Pertama>>', " "); 
  }
  try{
    do {
      var next = replaceTextToImage(body,'<<Foto kegiatan Kedua>>', link2, 180);
    } while (next);
  }catch{
    body.replaceText('<<Foto kegiatan Kedua>>', " "); 
  }
  try{
    do {
      var next = replaceTextToImage(body,'<<Foto kegiatan Ketiga>>', link3, 180);
    } while (next);
  }catch{
    body.replaceText('<<Foto kegiatan Ketiga>>', " "); 
  }
  try{
    do {
      var next = replaceTextToImage(body,'<<Foto kegiatan Keempat>>', link4, 180);
    } while (next);
  }catch{
    body.replaceText('<<Foto kegiatan Keempat>>', " "); 
  }
  try{
    do {
      var next = replaceTextToImage(body,'<<Foto kegiatan Kelima>>', link5, 180);
    } while (next);
  }catch{
    body.replaceText('<<Foto kegiatan Kelima>>', " "); 
  }
  try{
    do {
      var next = replaceTextToImage(body,'<<Foto kegiatan Keenam>>', link6, 180);
    } while (next);
  }catch{
    body.replaceText('<<Foto kegiatan Keenam>>', " "); 
  }

  //Lastly we save and close the document to persist our changes
  doc.saveAndClose(); 

  //Get ready send to email
  //next update
}
~~~


## References :
- [AutoFillText to Word](https://jeffreyeverhart.com/2018/09/17/auto-fill-google-doc-from-google-form-submission/)
- [Insert Inline Image()](https://www.labnol.org/code/20078-insert-image-in-google-document)
- [Replace text to image()](https://gist.github.com/tanaikech/f84831455dea5c394e48caaee0058b26)

