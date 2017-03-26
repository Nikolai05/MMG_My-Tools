var myExcel = new ExcelPlus();
myExcel.createFile([ "Christian", "Wireless Bulb", "LED Strip","USB Spy cam","Shipstation","Others" ])
myExcel.selectSheet("Christian");
myExcel.write({ "content":[ ["ID","QTY","SKU","Name","Country","Address_1","Address_2","City","State","Zip","Item","Tel_num"] ] })


/* processing array buffers, only required for readAsArrayBuffer */
function fixdata(data) {
  var o = "", l = 0, w = 10240;
  for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
  o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
  return o;
}

var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
/* set up drag-and-drop event */
var drop_dom_element = document.getElementById("dropHere");
function handleDrop(e){
  e.stopPropagation();
  e.preventDefault();
  var files = e.dataTransfer.files;
  var i,f;
  for (i = 0; i != files.length; ++i) {
    f = files[i];
    var reader = new FileReader();
    var name = f.name;
    reader.onload = function(e) {
      var data = e.target.result;

      var workbook;
      if(rABS) {
        /* if binary string, read with type 'binary' */
        workbook = XLSX.read(data, {type: 'binary'});
      } else {
        /* if array buffer, convert to base64 */
        var arr = fixdata(data);
        workbook = XLSX.read(btoa(arr), {type: 'base64'});
      }

      /* DO SOMETHING WITH workbook HERE */
      var first_sheet_name = workbook.SheetNames[0];
      /* Get worksheet */
      var worksheet = workbook.Sheets[first_sheet_name];
      generateExcel(XLSX.utils.sheet_to_json(worksheet));
    };
    if(rABS) reader.readAsBinaryString(f);
    else reader.readAsArrayBuffer(f);
    }
}

function handleDrag(e){
  e.preventDefault();
}

drop_dom_element.addEventListener('drop', handleDrop, false);
drop_dom_element.addEventListener("dragover", handleDrag, false);


function generateExcel(sheet_obj){
  var newRow_Christian=0,
      newRow_WirelessBulb=0,
      newRow_LEDStrip=0;
      newRow_USBSpy=0;
      newRow_Shipstation=0,
      newRow_Others=0;

  for(x=0;x<(Object.keys(sheet_obj).length);x++){
      var cust_id=sheet_obj[x].ID,
          cust_qty=sheet_obj[x].QTY,
          cust_sku=sheet_obj[x].SKU.toString(),
          cust_name=sheet_obj[x].Name,
          cust_country=sheet_obj[x].Country,
          cust_add1=sheet_obj[x].Address_1,
          cust_add2=sheet_obj[x].Address_2,
          cust_city=sheet_obj[x].City,
          cust_state=sheet_obj[x].State,
          cust_zip=sheet_obj[x].Zip,
          cust_item=sheet_obj[x].Item,
          cust_tel=sheet_obj[x].Tel_num;

          switch(cust_sku){
            //Police box case
            case "930238-for-iphone-5-5s":
            case "930238-for-iphone-6":
            case "930238-for-iphone-6s":
            case "930238-for-iphone-6-plus":
            case "930238-for-iphone-6s-plus":
            case "2174978-for-iphone-7":
            case "2174978-for-iphone-7-plus":
            case "930238-for-samsung-s6":
            case "930238-for-samsung-s6-edge":
            case "930238-for-s6-edge-plus":
            case "930238-for-samsung-s7":
            case "930238-for-samsung-s7-edge":
            case "930238-for-samsung-s7-plus":
            case "930238-for-samsung-note-5":
            //smart charging cable
            case "878527-gold-for-android":
            case "878527-gold-for-iphone":
            case "878527-sliver-for-android":
            case "878527-sliver-for-iphone":
            case "878527-white-for-android":
            case "878527-white-for-iphone":
            //Smart key organizer
            case "1916202-green":
            case "1916202-yellow":
            case "1916202-red":
            case "1916202-black":
            case "1916202-orange":
            //Antigravity iphone case
            case "2147803-white-for-iphone-7-plus":
            case "2147803-white-for-iphone-7":
            case "2147803-white-for-ip-6plus-6s-plus":
            case "2147803-white-for-iphone-6-6s":
            case "2147803-white-for-iphone-5-5s-se":
            case "2147803-black-for-iphone-7-plus":
            case "2147803-black-for-iphone-7":
            case "2147803-black-for-ip-6plus-6s-plus":
            case "2147803-black-for-iphone-6-6s":
            case "2147803-black-for-iphone-5-5s-se":
            //Aux bluetooth
            case "1739766":
            //Flexible phone dock tripod
            case "1835190-black-for-iphone5-6":
            case "1835190-black-for-andriod-phone":
            case "1835190-silver-for-iphone5-6":
            case "1835190-silver-for-andriod-phone":
            //Wireless charging station for iphone
            case "2617342-wirelesscharging-black":
            case "2617342-wirelesscharging-white":
            //Smart fitness tracker for ios and android
            case "1094276-black":
            case "1094276-blue":
            case "1094276-orange":
            case "1094276-green":
            case "1094276-yellow":
            case "1094276-with-green-strap":
            case "1094276-with-blue-strap":
            case "1094276-with-orange-strap":
            case "1094276-with-yellow-strap":
            case "1094276-with-4-colors-strap":
            //Glow in the dark earphone
            case "999941-green":
            case "999941-gold":
            case "999941-blue":
            case "999941-pink":
            case "999941-white":
            case "999941-purple":
            //mini style wireless bluetooth earphone
            case "2338825-black":
            case "2338825-blue":
            case "2338825-gold":
            case "2338825-pink":
            case "2338825-white":
            //mini in-ear wireless bluetooth earbud
            case "1732142-black":
            case "1732142-blue":
            case "1732142-green":
            case "1732142-white":
            //black car dashboard sticky pad mat
            case "3727390":
            //premium tempered glass screen protector for iphone
            case "536895-for-iphone-4-4s":
            case "536895-for-iphone-5-5s-se":
            case "536895-for-iphone-6plus":
            case "536895-for-iphone-6":
            case "536895-for-iphone-7-plus":
            case "536895-for-iphone-7":
            //patented slide card case
            case "1746176-rose-gold-for-iphone-7":
            case "1746176-mint-for-iphone-7":
            case "1746176-green-for-iphone-7":
            case "1746176-hot-pink-for-iphone-7":
            case "1746176-gold-for-iphone-7":
            case "1746176-black-for-iphone-7":
            case "1746176-navy-blue-for-iphone-7":
            case "1746176-silver-for-iphone-7":
            case "1746176-red-for-iphone-7":
            case "1746176-pink-for-iphone-7":
            case "1746176-white-for-iphone-7":
            case "1746176-rose-gold-for-iphone-7-plus":
            case "1746176-pink-for-iphone-7-plus":
            case "1746176-white-for-iphone-7-plus":
            case "1746176-mint-for-iphone-7-plus":
            case "1746176-green-for-iphone-7-plus":
            case "1746176-hot-pink-for-iphone-7-plus":
            case "1746176-gold-for-iphone-7-plus":
            case "1746176-black-for-iphone-7-plus":
            case "1746176-navy-blue-for-iphone-7-plus":
            case "1746176-silver-for-iphone-7-plus":
            case "1746176-red-for-iphone-7-plus":
            case "1746176-rose-gold-for-iphone-6-6s":
            case "1746176-rose-gold-for-iphone-6-6s-plus":
            case "1746176-rose-gold-for-iphone-5-5s-se":
            case "1746176-navy-blue-for-iphone-6-6s":
            case "1746176-navy-blue-for-iphone-6-6s-plus":
            case "1746176-navy-blue-for-iphone-5-5s-se":
            case "1746176-pink-for-iphone-6-6s":
            case "1746176-pink-for-iphone-6-6s-plus":
            case "1746176-pink-for-iphone-5-5s-se":
            case "1746176-black-for-iphone-6-6s":
            case "1746176-white-for-iphone-6-6s":
            case "1746176-red-for-iphone-6-6s":
            case "1746176-hot-pink-for-iphone-6-6s":
            case "1746176-gold-for-iphone-6-6s":
            case "1746176-silver-for-iphone-6-6s":
            case "1746176-mint-for-iphone-6-6s":
            case "1746176-green-for-iphone-6-6s":
            case "1746176-black-for-iphone-6-6s-plus":
            case "1746176-white-for-iphone-6-6s-plus":
            case "1746176-red-for-iphone-6-6s-plus":
            case "1746176-hot-pink-for-iphone-6-6s-plus":
            case "1746176-gold-for-iphone-6-6s-plus":
            case "1746176-silver-for-iphone-6-6s-plus":
            case "1746176-mint-for-iphone-6-6s-plus":
            case "1746176-green-for-iphone-6-6s-plus":
            case "1746176-black-for-iphone-5-5s-se":
            case "1746176-red-for-iphone-5-5s-se":
            case "1746176-hot-pink-for-iphone-5-5s-se":
            case "1746176-gold-for-iphone-5-5s-se":
            case "1746176-mint-for-iphone-5-5s-se":
            case "1746176-green-for-iphone-5-5s-se":
            case "1746176-silver-for-iphone-5-5s-se":
            case "1746176-white-for-iphone-5-5s-se":
            //patented slide card case 2
            case "2312-iphone-6-black":
            case "2312-iphone-6-plus-black":
            case "2312-samsung-s6-black":
            case "2312-samsung-s6-edge-black":
            case "2312-iphone-5-5s-black":
            case "2312-iphone-5c-black":
            case "2312-samsung-s6-edge-plus-black":
            case "2312-samsung-s7-edge-black":
            case "2312-samsung-s7-black":
            case "2312-iphone-7-black":
            case "2312-iphone-7-plus-black":
            case "2312-samsung-Note-7-black":
            case "2312-iphone-6-gold":
            case "2312-iphone-6-plus-gold":
            case "2312-samsung-s6-gold":
            case "2312-samsung-s6-edge-gold":
            case "2312-iphone-5-5s-gold":
            case "2312-iphone-5c-gold":
            case "2312-samsung-s6-edge-plus-gold":
            case "2312-samsung-s7-edge-gold":
            case "2312-samsung-s7-gold":
            case "2312-iphone-7-gold":
            case "2312-iphone-7-plus-gold":
            case "2312-samsung-Note-7-gold":
            case "2312-iphone-6-silver":
            case "2312-iphone-6-plus-silver":
            case "2312-samsung-s6-silver":
            case "2312-samsung-s6-edge-silver":
            case "2312-iphone-5-5s-silver":
            case "2312-iphone-5c-silver":
            case "2312-samsung-s6-edge-plus-silver":
            case "2312-samsung-s7-edge-silver":
            case "2312-samsung-s7-silver":
            case "2312-iphone-7-silver":
            case "2312-iphone-7-plus-silver":
            case "2312-samsung-Note-7-silver":
            //waterproof wireless bluetooth speaker
            case "3587876-blue":
            case "3587876-red":
            case "3587876-green":
            case "3587876-yellow":
            case "3587876-charge-cable":
            selectChristian(newRow_Christian);
            newRow_Christian++;
            break;
            case "1762267-EU-Plug":
            case "1762267-US-Plug":
            selectSpy(newRow_USBSpy);
            newRow_USBSpy++;
            break;
            case "3038350-warm-white-eu-plug":
            case "3038350-warm-white-us-plug":
            selectLed(newRow_LEDStrip);
            newRow_LEDStrip++;
            break;
            case "2925539-changeable-12w":
            case "4680659-changeable-12w":
            selectBulb(newRow_WirelessBulb);
            newRow_WirelessBulb++;
            break;
            //Earphone Charging Cable
            case "2165676":
            //360 phone holder
            case "579378-white":
            case "579378-black":
            case "579378-red":
            case "579378-blue":
            //Antigravity magical samsung galaxy case
            case "2147803-black-for-samsung-s6":
            case "2147803-black-for-samsung-s6-edge":
            case "2147803-black-for-sm-s6-edge-plus":
            case "2147803-black-for-samsung-s7":
            case "2147803-black-for-samsung-s7-edge":
            case "2147803-black-for-samsung-note-5":
            case "2147803-white-for-samsung-s6":
            case "2147803-white-for-samsung-s6-edge":
            case "2147803-white-for-sm-s6-edge-plus":
            case "2147803-white-for-samsung-s7":
            case "2147803-white-for-samsung-s7-edge":
            case "2147803-white-for-samsung-note-5":
            //Molle waist pouch
            case "836151-black":
            case "836151-armygreen":
            case "836151-acu":
            case "836151-sansha":
            case "836151-cpcolor":
            case "836151-jungle":
            case "836151-khaki":
            case "836151-digital-desert":
            selectShipstation(newRow_Shipstation);
            newRow_Shipstation++;
            break;
            default:
            selectOthers(newRow_Others);
            newRow_Others++;
            break;
          }
        }

        myExcel.saveAs("output.xlsx");


    function selectChristian(nextRow){
    myExcel.selectSheet("Christian");
    writeData(nextRow);
    }

    function selectBulb(nextRow){
    myExcel.selectSheet("Wireless Bulb");
    writeData(nextRow);
    }

    function selectLed(nextRow){
    myExcel.selectSheet("LED Strip");
    writeData(nextRow);
    }

    function selectSpy(nextRow){
    myExcel.selectSheet("USB Spy cam");
    writeData(nextRow);
    }

    function selectOthers(nextRow){
    myExcel.selectSheet("Others");
    writeData(nextRow);
    }

    function selectShipstation(nextRow){
    myExcel.selectSheet("Shipstation");
    writeData(nextRow);
    }

    function writeData(thisCell){
    myExcel.write({ "cell":"A"+(thisCell+2).toString(), "content":cust_id })
    myExcel.write({ "cell":"B"+(thisCell+2).toString(), "content":cust_qty })
    myExcel.write({ "cell":"C"+(thisCell+2).toString(), "content":cust_sku })
    myExcel.write({ "cell":"D"+(thisCell+2).toString(), "content":cust_name })
    myExcel.write({ "cell":"E"+(thisCell+2).toString(), "content":cust_country })
    myExcel.write({ "cell":"F"+(thisCell+2).toString(), "content":cust_add1 })
    myExcel.write({ "cell":"G"+(thisCell+2).toString(), "content":cust_add2 })
    myExcel.write({ "cell":"H"+(thisCell+2).toString(), "content":cust_city })
    myExcel.write({ "cell":"I"+(thisCell+2).toString(), "content":cust_state })
    myExcel.write({ "cell":"J"+(thisCell+2).toString(), "content":cust_zip })
    myExcel.write({ "cell":"K"+(thisCell+2).toString(), "content":cust_item })
    myExcel.write({ "cell":"L"+(thisCell+2).toString(), "content":cust_tel })
    }
}
