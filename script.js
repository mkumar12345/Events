var url = "Join Events.xlsx";

/* set up async GET request */
var req = new XMLHttpRequest();
req.open("GET", url, true);
req.responseType = "arraybuffer";

req.onload = function (e) {
  var data = new Uint8Array(req.response);
  var workbook = XLSX.read(data, { type: "array" });
  /* DO SOMETHING WITH workbook HERE */
  var sheet_name_list = workbook.SheetNames;
  var json_sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  for (let i = 0; i < json_sheet.length; i++) {
    var card = document.createElement("a");
    card.classList.add("card");

    // top div
    var top = document.createElement("div");
    top.classList.add("top")
    var company = document.createElement("img");
    company.classList.add("company");
    var price = document.createElement("img");
    price.classList.add("price");
    top.appendChild(company);
    top.appendChild(price);
    card.appendChild(top)


    // image div
    var img = document.createElement("div");
    img.classList.add("img");
    var illustration = document.createElement("img");
    illustration.classList.add = "illustration";
    img.appendChild(illustration);
    card.appendChild(img)

    //time div
    var time = document.createElement("div");
    time.classList.add("time");
    card.appendChild(time)

    //footer div
    var card_footer = document.createElement("div");
    card_footer.classList.add("card_footer");

    var title = document.createElement("div");
    title.classList.add("title");
    card_footer.appendChild(title)
    card.appendChild(card_footer)


s = json_sheet[i];

// getting date in correct format
var date = new Date(s.Date);
var js_date = ExcelDateToJSDate(date);
var final_date = formatDate(js_date); 
// assigning values to dom
console.log(s)
card.style.backgroundColor=s.Color
company.src = s["Company Image"]
price.src=s['Price Image'];
illustration.src = s.Illustration;
time.innerHTML = final_date + s.Time;
title.innerHTML = s.Title;

card.href=s.Link;
card.target="_blank"
document.getElementById("cards").appendChild(card); //apend to cards flexbox

  }
}

function ExcelDateToJSDate(serial) {
  var utc_days  = Math.floor(serial - 25569);
  var utc_value = utc_days * 86400;                                        
  var date_info = new Date(utc_value * 1000);

  var fractional_day = serial - Math.floor(serial) + 0.0000001;

  var total_seconds = Math.floor(86400 * fractional_day);

  var seconds = total_seconds % 60;

  total_seconds -= seconds;

  var hours = Math.floor(total_seconds / (60 * 60));
  var minutes = Math.floor(total_seconds / 60) % 60;

  return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}
function formatDate(date) {
  var d = new Date(date),
      month = '' + (d.getMonth() + 1),
      day = '' + d.getDate(),
      year = d.getFullYear();

  if (month.length < 2) 
      month = '0' + month;
  if (day.length < 2) 
      day = '0' + day;

  return [ month,day,year ].join('/');
}
req.send();

