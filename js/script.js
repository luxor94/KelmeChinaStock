let kelmePrice
//импорт файла из xlsx в JSON
var ExcelToJSON = function() {

  this.parseExcel = function(file) {
    var reader = new FileReader();

    reader.onload = function(e) {
      var data = e.target.result;
      var workbook = XLSX.read(data, {
        type: 'binary'
      });
      workbook.SheetNames.forEach(function(sheetName) {
        // Here is your object
        var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        var json_object = JSON.stringify(XL_row_object);
        kelmePrice = (JSON.parse(json_object));
        jQuery( '#xlx_json' ).val( json_object );
      })
    };

    reader.onerror = function(ex) {
      console.log(ex);
    };

    reader.readAsBinaryString(file);
  };
};

function replacement() {

  //замена артикула
  let listArt = [
    "871001-809",
    "871303-600",
    "871306-488",
    "871307-251",
    "871307-924",
    "871308-251",
    "871308-924",
    "871309-251",
    "871309-924",
    "871310-400",
    "871310-600",
    "871311-400",
    "871313-251",
    "871320-600",
    "871321-107",
    "871951-904",
    "872306-488",
    "872307-251",
    "872307-924",
    "872308-251",
    "872308-924",
    "872309-251",
    "872309-924",
    "872310-400",
    "872310-600",
    "872311-400",
    "872313-251",
    "872320-600",
    "872321-107",
    "873701-905",
    "873701-907",

  ]; //массив(список) где надо записать артикул без цвета


  for (let i = 0; i < kelmePrice.length; i++) {
    if (listArt.includes(kelmePrice[i].art)) kelmePrice[i].art = kelmePrice[i].art.slice(0, -4);
      
  }



  //замена размеров
  let listArtSize = [
    "K15Z908",
    "9876312",
    "9876311",
    "9782706",
    "9496837",
    "9096946",
    "K15Z976",
    "K15Z958",
    "K15Z934",
    "K15Z909",
    "K15Z907",
    "9896318",
    "9876309",
    "9876308",
    "9876307",
    "9876306",
    "9876305",
    "9876304",
    "9876303",
    "9876302",
    "9876301",
    "9596317",
    "9086832",
    "ZPDW001",
    "9996578",
    "9996577",
    "96914003",
    "96914002",
    "96910004",
    "96840009",
    "96840008",
    "96830004",
    "9881406",
    "KMA16003",
    "K15Z9110",
    "9996571",
    "9886405",
    "9886404",
    "9886711",
    "9886015",
    "9886016",
    "9886215",
    "9886214",
    "9886211",
    "9886210",
    "9886209",
    "9876203",
    "9876202",
    "9876200",
    "K15S948",
    "9886207",
    "9492217",
    "9491215",
    "7100WZ5001",
    "8101HJ5001",
    
  ]; //массив(список) где не надо менять размеры

  for (let i = 0; i < kelmePrice.length; i++) {

    if (kelmePrice[i].size === '2XS' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "3XS";
    }
    if (kelmePrice[i].size === 'XS' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "2XS";
    }

    if (kelmePrice[i].size === "S" && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "XS";
    }
    if (kelmePrice[i].size === 'M' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "S";
    }
    if (kelmePrice[i].size === 'L' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "M";
    }

    if (kelmePrice[i].size === 'XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "L";
    }
    if (kelmePrice[i].size === '2XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "XL";
    }
    if (kelmePrice[i].size == '3XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "2XL";
    }
    if (kelmePrice[i].size == '4XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "3XL";
    }
    if (kelmePrice[i].size == '5XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "4XL";
    }
    if (kelmePrice[i].size == '6XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "5XL";
    }
    if (kelmePrice[i].size == '7XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "6XL";
    }
    if (kelmePrice[i].size == '8XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "7XL";
      
          if (kelmePrice[i].size == '9XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "8XL";
    }
          if (kelmePrice[i].size == '10XL' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "9XL";
    }
    }
    if (kelmePrice[i].size == 'Free Size' && !listArtSize.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "UNI";
    } else {
      continue
    }

  }



  //список артикулов с размеров гетр L
  let listSocksL = [
    "K15Z901",
    "9876313",
    "K16A810",
    "K16A812",
    "K16A806",
    "K16A811",
    "WZ60201010",
    "WZ60301006",
    "WZ60201004",
    "WZ60201006",
    "WZ60201009",
    "WZ60201011",
    "WZ60301007",
    "WZ60201003",
    "WZ60201005",
    "WZ60201007",
    "WZ60302003",
    "8101WZ5001",
    "8101WZ5003",
    "9996568",
    "8101WZ5001",
    "8101WZ5002",
    "8101WZ5003",
    "0101WZ5002",
    "8101WZ5005",
    "8102WZ1001",
    "8102WZ1002",
    "8102WZ1003",
    "8102WZ1004",
    "8102WZ5002",
    "WZ60201008",
    
  ]

  //список артикулов с размеров гетр M
  let listSocksM = [
    "K16A809",
    "K16A808",
    "9991548",
    "9492217",
    "9491215",
    "WZ60202010",
    "WZ60202011",
    "WZ60202015",
    "WZ60302004",
    "WZ60302005",



  ]


  //список артикулов с размеров гетр 8
  let listSocks8 = ["K15Z931",
    "9993574",
    "8101WZ3001",
    "0101WZ3001",  
    "8102WZ5001",
    "8102WZ5003",
                    
   ]



  //список артикулов с размеров гетр 6
  let listSocks6 = [
    "9893319",


  ]


  for (let i = 0; i < kelmePrice.length; i++) {
    if (kelmePrice[i].size == 'UNI' && listSocksL.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "L";
    }
    if (kelmePrice[i].size == 'UNI' && listSocksM.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "M";
    }
    if (kelmePrice[i].size == 'UNI' && listSocks8.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "8";
    }
    if (kelmePrice[i].size == 'UNI' && listSocks6.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "6";
    }
  }

  for (let i = 0; i < kelmePrice.length; i++) {
    kelmePrice[i].art = `${kelmePrice[i].art}-${kelmePrice[i].color}`;
    delete kelmePrice[i].color;
    if (kelmePrice[i].size === undefined) delete  kelmePrice[i];
  }

  





  //экспорт в json
  function downloadObjectAsJson(kelmePrice, exportName) {
    var dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(kelmePrice));
    var downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", exportName + ".json");
    document.body.appendChild(downloadAnchorNode); // required for firefox
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  }
  //downloadObjectAsJson(kelmePrice, 'kelme');  //экспорт json файла

  //импорт в XLSX

  // We will make a Workbook contains 2 Worksheets

  // export json to Worksheet of Excel
  // only array possible
  var stockList = XLSX.utils.json_to_sheet(kelmePrice)


  // A workbook is the name given to an Excel file
  var wb = XLSX.utils.book_new() // make Workbook of Excel

  // add Worksheet to Workbook
  // Workbook contains one or more worksheets
  XLSX.utils.book_append_sheet(wb, stockList, 'KelmePriceChina') // sheetAName is name of Worksheet


  // export Excel file
  XLSX.writeFile(wb, 'book.xlsx') // name of the file is 'book.xlsx'
}
