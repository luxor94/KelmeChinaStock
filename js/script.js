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
    "7100WZ5001",
    "8101HJ5001",
    "9001HJ5001",
    "9496817",
    "8201WZ5018",
    "8201HJ5003",
    "8301HJ5005",
    "96830005",
    "96840006",
    "96840007",
    "9086836",
    "0101WZ5003",
    "K15Z933",
    "8201WZ5016",
    "8201WZ5017",
    "9302HJ5008",
    "9302HJ5009",
    "9491172",
    "9996549",
    "9876201",
    "9302HJ5006",
    "9302HJ5007",
    "7301WZ5104",
    "9402WZ5156",
    "9402WZ5162",
    "9402WZ5163",
    "9402WZ5169",
    "9401WZ5164",
    "9402WZ5157",
    "9402WZ5167",
    "7451MJ1018",
    "7451PL1112",
    "7451PL1113",
    "9402WZ5160",
    "9402WZ5178",
    "9402WZ5179",
    "9401WZ5165",
    "9402WZ5161",
    "7401WZ5173",
    "7401WZ5180",

    
    
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
    "8201WZ5001",
    "9492217",
    "8201WZ5002",
    "9491213",
    "8301WZ5056",
    "8301WZ5071",
    "8301WZ5076",
    "8301WZ5087",
    "9993573",
    "9491214",
    "9491215",
    "8301WZ5025",
    "6307WZ5004",
    "6307WZ5006",
    "6307WZ5010",
    "6307WZ5007",
    "6307WZ5008",
    "6307WZ5005",
    "6307WZ5012",
    "6307WZ5009",
    "8302WZ5030",
    "8302WZ5032",
    "8302WZ5027",
    "8302WZ5029",
    "8302WZ5028",
    "8302WZ5031",
    "9302WZ5043",
    "9302WZ5044",
    "9302WZ5045",
    "9302WZ5049",
    "9302WZ5052",
    "9302WZ5053",
    "9302WZ5062",
    "9302WZ5063",
    "9302WZ5064",
    "9302WZ5065",
    "9302WZ5070",
    "9302WZ5058",
    "9302WZ5059",
    "9302WZ5060",
    "9302WZ5061",
    "8301WZ5057",
    "9303WZ5054",
    "6367WZ5007",
    "6367WZ5002",
    "6367WZ5004",
    "6367WZ5006",
    "6367WZ5003",
    "9886216",
    "9886213",
    "96831000",
    "9986530",
    "9986531",
    "8201HJ5004",
    "9302HJ5010",
    "8301HJ5012",
    "9302WZ5093",
    "9302WZ1082",
    "9302WZ1086",
    "9302WZ1084",
    "9302WZ5092",
    "9302WZ5091",
    "8301HJ5011",
    "9301WZ5047",
    "9301WZ5066",
    "9301HJ5014",
    "9302WZ5098",
    "9302WZ5102",
    "9302WZ5103",
    "8401WZ5095",
    "8401WZ5096",
    "9301WZ5101",
    "9402WZ5109",
    "9402WZ5105",
    "9402WZ5106",
    "9402WZ5107",
    "9402WZ5108",
    "8401WZ5097",
    "6407WZ5019",
    "6403WZ1008",
    "6407WZ5018",
    "6403WZ5013",
    "6403WZ5016",
    "6403WZ5017",
    "6403WZ5009",
    "6407WZ5020",
    "6403WZ5015",
    "6403WZ5012",
    "6407WZ5021",
    "6403WZ5011",
    "6403WZ5010",
    "6403WZ5014",
    "9402WZ5110",
    "7401WZ5121",
    "9402WZ5120",
    "9402WZ5149",
    "9402WZ5152",
    "9402WZ5116",
    "9402WZ5119",
    "9402WZ5115",
    "9402WZ5118",
    "9402WZ5114",
    "9402WZ5117",
    "7401WZ5153",
    "8301WZ5099",
    "9412ZX1176",
    "9401WZ5155",
    "9402WZ5128",
    "9402WZ5129",
    "9402WZ5146",
    "9402WZ5127",
    "9402WZ5139",
    "9402WZ5142",
    "9402WZ5137",
    "9402WZ5143",
    "9402WZ5150",
    "9402WZ5151",
    "9402WZ5131",
    "9402WZ5135",
    "9402WZ5136",
    "9402WZ5130",
    "9402WZ5138",
    "9402WZ5148",
    "9402WZ5134",
    "9402WZ5147",
    "9402WZ5113",
    "9402WZ5132",
    "9402WZ5133",
    "9402WZ5140",
    "9402WZ5144",
    "9402WZ5170",
    "9402WZ5145",
    "9402WZ5141",
    "9402WZ5166",
    "7401WZ5176",  
    "7401WZ5183",
    "8431ZX1233",
    "7401WZ5184",
    "6467WZ5012",
    "6467WZ5019",
    "6467WZ5017",
    "6467WZ5018",
    "6467WZ5016"
    
  ]

  //список артикулов с размеров гетр M
  let listSocksM = [
    "K16A809",
    "K16A808",
    "9991548",
    "WZ60202010",
    "WZ60202011",
    "WZ60202015",
    "WZ60302004",
    "WZ60302005",
    "8301WZ5055",
    "K16A807",
    "K16A808",
    "K16A809",
    "6307WZ5002",
    "6307WZ5001",
    "6307WZ5011",
    "9492216",
    "8161ST5002",
    "8161ST5001",
    "9491171",
    "8103ST5001",
    "8161ST5003",
    "K15Z911",
    "8261ST5006",
    "9302WZ2081",
    "9302WZ2085",
    "9302WZ2079",
    "9302WZ2083",

  ]


  //список артикулов с размеров гетр 8
  let listSocks8 = ["K15Z931",
    "9993574",
    "8101WZ3001",
    "0101WZ3001",  
    "8102WZ5001",
    "8102WZ5003",
    "9301WZ3048",
    "8301WZ3056",
    "8301WZ3087",
    "9302WZ3078",
    "9302WZ3077",
    "9302WZ3046",
    "8401WZ3097",
    "9402WZ3111",   
    "8401WZ5112",
    "7401WZ3176",
                    
                    
   ]



  //список артикулов с размеров гетр 6
  let listSocks6 = [
    "9893319",
    "8201WZ5003",


  ]
  
    //список артикулов с размером мячей 5
  let listSocks5 = [
    "9096148",


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
    if (kelmePrice[i].size == 'UNI' && listSocks5.includes(kelmePrice[i].art)) {
      kelmePrice[i].size = "5";
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
