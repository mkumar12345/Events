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
  create_filters(json_sheet)
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
    card.style.backgroundColor = s.Color
    company.src = s["Company Image"]
    price.src = s['Price Image'];
    illustration.src = s.Illustration;
    time.innerHTML = final_date + s.Time;
    title.innerHTML = s.Title;
    card.href = s.Link;
    card.target = "_blank"
    card.dataset.producer = s['Producer']
    card.dataset.place = s['Place']
    card.dataset.productType = s['Product Type']
    card.dataset.type = s['Type']
    card.dataset.whoCanAttend = s['Who Can Attend']
    document.getElementById("cards").appendChild(card); //apend to cards flexbox

  }
}

function ExcelDateToJSDate(serial) {
  var utc_days = Math.floor(serial - 25569);
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

  return [month, day, year].join('/');
}
req.send();

function create_filters(sheet) {
  var options = Object.keys(sheet[0]);
  options.forEach(function (fetched_option) { //option = filter label
    var drop_down = document.createElement("div");  //create drop down filter
    drop_down.classList.add("drop_down");
    var option = document.createElement("div"); //create name of filter
    option.classList.add("option");
    option.innerHTML = `+ ` + fetched_option;

    var values = document.createElement('div');    //create values list div
    values.classList.add("values");
    var fetched_values = []
    for (let i = 0; i < sheet.length; i++) {         //fetch all values
      //save all values in array
      if (!fetched_values.includes(sheet[i][fetched_option])) {
        fetched_values.push(sheet[i][fetched_option])

      }
    }

    var selected = [];
    var selected_holder = document.createElement("div");
    selected_holder.classList.add("selected_holder");

    fetched_values.forEach(function (value) {
      var enabled = document.createElement("div");
      enabled.innerHTML="Enabled"
      var disabled = document.createElement("div");
      disabled.innerHTML="Disabled"
      var new_value = document.createElement('div');     //create list of values from array for each option
      new_value.dataset.selected = false;
      new_value.id = value;
      new_value.dataset.name = value;
      new_value.dataset.type = fetched_option;
      new_value.classList.add("value");
      new_value.innerHTML = value;
      new_value.appendChild(disabled)
      new_value.onclick = function () {
        new_value.dataset.selected = true;
        new_value.style.display = "none"
        var clicked = document.createElement("div");
        clicked.classList.add("selected");
        clicked.dataset.name = new_value.dataset.name;
        clicked.innerHTML =value;
        clicked.appendChild(enabled)
        clicked.dataset.type = fetched_option;
        selected.push(clicked);
        selected.forEach(function (val) {
          selected_holder.appendChild(val);
          val.onclick = function (e) {
            var values_content = values.getElementsByClassName('value');
            for (let i = 0; i < values_content.length; i++) {
              if (values_content[i].dataset.name === e.target.dataset.name) {
                values_content[i].dataset.selected = "false"
                values_content[i].style.display = "flex"
              }
            }
            // new_value.dataset.selected = false;
            // new_value.style.display = "block"
            val.remove()
            const exists = selected.indexOf(val);
            if (exists > -1) {
              selected.splice(exists, 1)
            }
            filter()
          }
        })


        filter();
      }
      values.appendChild(new_value)
    })


    if (fetched_option === "Producer" || fetched_option === "Place" || fetched_option === "Type" || fetched_option === "Product Type" || fetched_option === "Who Can Attend") {
      drop_down.appendChild(option);
      drop_down.appendChild(values);
      drop_down.appendChild(selected_holder)
      document.getElementById("drop_downs").appendChild(drop_down);

    }

  })

}

// filtering function
function filter() {
  var all_options = Array.from(document.getElementsByClassName("value"))
  var active_options = [];
  for (let i = 0; i < all_options.length; i++) {
    if (all_options[i].dataset.selected === "true") {
      active_options.push(all_options[i])
    }
  }

  var selected_types = [];
  active_options.forEach(function(opt){           //count selected number of types
    if(!selected_types.includes(opt.dataset.type)){
      selected_types.push(opt.dataset.type)
    }
  })

  var all_cards = document.getElementsByClassName('card');
  var matches = []
  for (let i = 0; i < all_cards.length; i++) {
    all_cards[i].style.display = "inline-block"
  }
  if (active_options.length > 0) {
    for (let i = 0; i < all_cards.length; i++) {
      all_cards[i].style.display = "none"
      matches.push(0)
    }
    for (let i = 0; i < active_options.length; i++) {
      var current_type = _.camelCase(active_options[i].dataset.type);
      var current_name = active_options[i].dataset.name.toLowerCase();
      if (current_type !== undefined && current_name !== undefined) {
        for (let j = 0; j < all_cards.length; j++) {
          if (all_cards[j].dataset[current_type].toLowerCase() === current_name) {
            matches[j]++;
           if(matches[j]===selected_types.length){
            all_cards[j].style.display = "inline-block"
            console.log(all_cards[j])
           }
           else{
            all_cards[j].style.display="none"
           
          }
          }
         
        }
      }
    }
  }
 
}



  