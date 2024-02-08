function doGet(e) {
  var output = HtmlService.createHtmlOutputFromFile('Index').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return output;
}


function autoFillGoogleSlideFromForm(response) {

  //Id of the template-google-slide file
  var srcPresentationId = "13EcBln-UhIY3AcD2QcwMiiKikMiIxgsNTN8zaOpKQis";

  //file is the template file, and you get it by ID
  var file = DriveApp.getFileById(srcPresentationId); 
  
  //We can make a copy of the template, name it, and optionally tell it what folder to live in
  //file.makeCopy will return a Google Drive file object
  var folder = DriveApp.getFolderById('1nVBpoLP-6iH8TeSMlPsgliNGoySMlMRh')
  var copy = file.makeCopy(response[1], folder); 

 
  
  var copysrcSlideIndex = 0; // 0 means page 1.

  var copydstSlideIndex = 0; // 0 means page 1.

  var src = SlidesApp.openById(srcPresentationId).getSlides()[copysrcSlideIndex];

  //Once we've got the new file created, we need to open it as a document by using its ID
  var slide = SlidesApp.openById(copy.getId());
  //slide.getSlides()[0].insertShape(src.getShapes());
  
  //Since everything we need to change is in the body, we need to get that
  var slides = slide.getSlides(); 
  var range=[];
  var day = [];
  var daystring;
  var month = {
    1:'มกราคม',
    2:'กุมภาพันธ์',
    3:'มีนาคม',
    4:'เมษายน',
    5:'พฤษภาคม',
    6:'มิถุนายน',
    7:'กรกฎาคม',
    8:'สิงหาคม',
    9:'กันยายน',
    10:'ตุลาคม',
    11:'พฤษจิกายน',
    12:'ธันวาคม'
  }
  slides.forEach(function(page){
     var shapes = (page.getShapes());
     shapes.forEach(function(shape){
      day = response[0].split('-');
      day[0]=(parseInt(day[0])+543).toString();
      daystring = parseInt(day[2]).toString()+' '+month[parseInt(day[1])]+' '+day[0];
       shape.getText().replaceAllText('{{date}}',daystring);
       shape.getText().replaceAllText('{{name}}',response[1]);
       shape.getText().replaceAllText('{{amount}}',response[2]);
       for(var i=3;i<14;i++){
        if(response[i]==0){shape.getText().replaceAllText('{{mark'+(i-2)+'}}',"");}
        else{shape.getText().replaceAllText('{{mark'+(i-2)+'}}',"✔");}
       }
       /*
       shape.getText().replaceAllText('{{mark2}}',firstName);
       shape.getText().replaceAllText('{{mark3}}',firstName);
       shape.getText().replaceAllText('{{mark4}}',firstName);
       shape.getText().replaceAllText('{{mark5}}',firstName);
       shape.getText().replaceAllText('{{mark6}}',firstName);
       shape.getText().replaceAllText('{{mark7}}',firstName);
       shape.getText().replaceAllText('{{mark8}}',firstName);
       shape.getText().replaceAllText('{{mark9}}',firstName);
       shape.getText().replaceAllText('{{mark10}}',firstName);
       shape.getText().replaceAllText('{{mark11}}',firstName);*/

       shape.getText().replaceAllText('{{citizenid}}',response[14]);
       shape.getText().replaceAllText('{{address}}',response[15]);
       shape.getText().replaceAllText('{{telnumber}}',response[16]);
       shape.getText().replaceAllText('{{mail}}',response[17]);
       day=response[18].split('-');
       day[0]=(parseInt(day[0])+543).toString();
       daystring=parseInt(day[2]).toString()+' '+month[parseInt(day[1])]+' '+day[0];
       shape.getText().replaceAllText('{{dateofbirth}}',daystring);
       if(response[19]==0){shape.getText().replaceAllText('{{mark12}}',"");}
       else{shape.getText().replaceAllText('{{mark12}}',"✔");}
       shape.getText().replaceAllText('{{amountthai}}',response[20]);
       for(var j=21; j<25; j++){
        if(response[j]==0||response[j]=='0'){
          shape.getText().replaceAllText('{{amountmark'+(j-20).toString()+'}}',"");
        }
        else{shape.getText().replaceAllText('{{amountmark'+(j-20).toString()+'}}',response[j]+' บาท');}
       }


     }); 
   })
     slide.saveAndClose(); 
     return "yay";

}


