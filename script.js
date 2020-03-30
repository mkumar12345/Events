var url = "Join Events.xlsx";
var json_sheet;
/* set up async GET request */
var req = new XMLHttpRequest();
req.open("GET", url, true);
req.responseType = "arraybuffer";

req.onload = function (e) {
  var data = new Uint8Array(req.response);
  var workbook = XLSX.read(data, { type: "array" });
  /* DO SOMETHING WITH workbook HERE */
  var sheet_name_list = workbook.SheetNames;
  json_sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
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
    time.innerHTML = final_date + " " + s.Time;
    title.innerHTML = s.Title;
    // card.href = s.Link;
    // card.target = "_blank"
    card.dataset.thumbnail = s.Illustration;
    card.dataset.link = s.Link;
    card.dataset.time = time.innerHTML;
    card.dataset.title = s['Title']
    card.dataset.producer = s['Producer']
    card.dataset.description = s["Description"]
    card.dataset.time = time.innerHTML;
    card.dataset.place = s['Place']
    card.dataset.productType = s['Product Type']
    card.dataset.type = s['Type']
    card.dataset.whoCanAttend = s['Who Can Attend']
    card.dataset.serialDate = s['Date']
    document.getElementById("cards").appendChild(card); //apend to cards flexbox
    card.addEventListener("click", function () {
      pop(this)
    })
  }
  paginate(document.getElementsByClassName("card"))

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


var global_keys = [];                //MAKING ALL OPTIONS BE AVAILABL GLOBALLY


function create_filters(sheet) {
  var options = Object.keys(sheet[0]);
  global_keys = options;
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
      enabled.innerHTML = "Enabled"
      var disabled = document.createElement("div");
      disabled.innerHTML = "Disabled"
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
        clicked.innerHTML = value;
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
  var active_cards = []
  var all_options = Array.from(document.getElementsByClassName("value"))
  var active_options = [];
  for (let i = 0; i < all_options.length; i++) {
    if (all_options[i].dataset.selected === "true") {
      active_options.push(all_options[i])
    }
  }

  var selected_types = [];
  active_options.forEach(function (opt) {           //count selected number of types
    if (!selected_types.includes(opt.dataset.type)) {
      selected_types.push(opt.dataset.type)
    }
  })

  var all_cards = document.getElementsByClassName('card');
  var matches = []
  for (let i = 0; i < all_cards.length; i++) {
    all_cards[i].style.display = "inline-block"
    all_cards[i].classList.remove("hidden-by-filter")
  }
  if (active_options.length > 0) {
    for (let i = 0; i < all_cards.length; i++) {
      all_cards[i].style.display = "none"
      matches.push(0);
      all_cards[i].classList.add("hidden-by-filter")
    }
    for (let i = 0; i < active_options.length; i++) {
      var current_type = _.camelCase(active_options[i].dataset.type);
      var current_name = active_options[i].dataset.name.toLowerCase();                //look man you will just need to dry run this and and try to understand what's going on, I can't explain. But the job of this part is to do the main filtering.
      if (current_type !== undefined && current_name !== undefined) {
        for (let j = 0; j < all_cards.length; j++) {
          if (all_cards[j].dataset[current_type].toLowerCase() === current_name) {
            matches[j]++;
            if (matches[j] === selected_types.length) {
              all_cards[j].style.display = "inline-block"
              all_cards[j].classList.remove("hidden-by-filter");
              active_cards.push(all_cards[j])
            }
            else {
              all_cards[j].style.display = "none"
              all_cards[j].classList.add("hidden-by-filter")
            }
          }

        }
      }
    }
  }
  if (active_cards.length > 0) {
    paginate(active_cards)
  }
  else {
    paginate(document.getElementsByClassName("card"))
  }
}



var max_date = 43725;
var min_date = 43550;
var current_max = 175;
var current_min = -21;


dragElement(document.getElementById("max"));
dragElement(document.getElementById("min"));

function dragElement(elmnt) {
  var pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
  // otherwise, move the DIV from anywhere inside the DIV:
  elmnt.onmousedown = dragMouseDown;
  elmnt.ontouchstart = dragMouseDown;
  function dragMouseDown(e) {
    e = e || window.event;
    e.preventDefault();
    // get the mouse cursor position at startup:
    pos3 = e.clientX || e.touches[0].clientX
    pos4 = e.clientY || e.touches[0].clientY
    document.onmouseup = closeDragElement;
    document.ontouchend = closeDragElement;
    // call a function whenever the cursor moves:
    document.onmousemove = elementDrag
    document.ontouchmove = elementDrag
  }

  function elementDrag(e) {
    e = e || window.event;
    e.preventDefault();
    // calculate the new cursor position:
    pos1 = pos3 - e.clientX ||  pos3 - e.touches[0].clientX
    pos2 = pos3 - e.clientY ||  pos4 - e.touches[0].clientY
    pos3 = e.clientX || e.touches[0].clientX
    pos4 = e.clientY || e.touches[0].clientY
    if (elmnt.id === "max") {
      current_max = elmnt.offsetLeft - pos1;
      if (current_max <= 175 && current_max > current_min + 20) {
        elmnt.style.left = (elmnt.offsetLeft - pos1) + "px";
        max_date = 43550 + current_max;
        var get_date = ExcelDateToJSDate(max_date);
        var final_date = formatDate(get_date);
        document.getElementById("current_max_date").innerHTML = final_date
      }
    }
    if (elmnt.id === "min") {
      current_min = elmnt.offsetLeft - pos1;
      if (current_min < current_max - 20 && current_min > -21) {
        elmnt.style.left = (elmnt.offsetLeft - pos1) + "px";
        min_date = 43550 + current_min;
        var get_date = ExcelDateToJSDate(min_date);
        var final_date = formatDate(get_date);
        document.getElementById("current_min_date").innerHTML = final_date
      }
    }
    // set the element's new position:

    filter_date();
  }
  function closeDragElement() {
    // stop moving when mouse button is released:
    document.onmouseup = null;
    document.onmousemove = null;
    document.ontouchend = null;
    document.ontouchmove = null;
  }
}
function filter_date() {
  var active_cards = [];
  var cards = document.getElementsByClassName("card");
  for (let i = 0; i < cards.length; i++) {
    var date = parseInt(cards[i].dataset.serialDate);
    if (date < min_date || date > max_date) {
      cards[i].style.display = "none"
    }

    else if (!cards[i].classList.contains("hidden-by-filter")) {
      cards[i].style.display = "block";
      active_cards.push(cards[i])
    }
  }
  if (active_cards.length > 0) {
    paginate(active_cards)
  }

}

function search(text) {
  var active_cards = [];
  var cards = document.getElementsByClassName("card");
  var filter = text.toUpperCase();
  for (let i = 0; i < json_sheet.length; i++) {
    s = json_sheet[i];
    var tags = s['Producer'] + s['Place'] + s['Region'] + s['Type'] + s['Product Type'] + s['Title'] + s['Who Can Attend'];
    if (tags.toUpperCase().indexOf(filter) < 0) {
      cards[i].style.display = "none"
    }
    else if (!cards[i].classList.contains("hidden-by-filter")) {
      cards[i].style.display = "block"
      active_cards.push(cards[i]);
      console.log(active_cards)
    }
  }
  if (active_cards.length > 0) {
    paginate(active_cards)
  }
}

document.getElementById("search").onfocus = function () {
  document.getElementById("search_holder").style.maxWidth = "100%"
}
document.getElementById("search").onblur = function () {
  document.getElementById("search_holder").style.maxWidth = "500px"
}

var popup = document.getElementById("popup_holder");
var popup_inner = document.getElementsByClassName("popup")[[0]]
document.getElementById("close").onclick = function () {
  popup.style.visibility = "hidden";
  popup_inner.style.transform = "scale(0)"
  popup_inner.style.opacity = "0"
}
function pop(element) {
  popup.style.visibility = "visible"
  popup_inner.style.transform = "scale(1)"
  popup_inner.style.opacity = "1"
  document.getElementById("popup_thumbnail").src = element.dataset['thumbnail']
  document.getElementById("popup_time").innerHTML = element.dataset['time']
  document.getElementById("popup_title").innerHTML = element.dataset['title']
  document.getElementById("popup_description").innerHTML = element.dataset['description']
  document.getElementById("property_producer").innerHTML = element.dataset['producer']
  document.getElementById("property_place").innerHTML = element.dataset['place']
  document.getElementById("property_type").innerHTML = element.dataset['type']
  document.getElementById("property_product_type").innerHTML = element.dataset.productType
  document.getElementById("property_who_can_attend").innerHTML = element.dataset.whoCanAttend
  document.getElementById("register_link").href = element.dataset.link;
  document.getElementById("register_link").target = "_blank"
}

////////////////PAGINATION///////////////////////



//CONSTANTS

const MAX_PER_PAGE = 8;
// Variables

var current_page = 1;
// pagination function

function paginate(cards) {
  document.getElementsByClassName("pagination_buttons")[0].innerHTML = ""
  current_page = 1;
  var page_assigner = 1;
  var page1 = document.createElement("div");
  page1.classList.add("page_button");
  page1.classList.add("active_page");
  page1.innerHTML = page_assigner
  document.getElementsByClassName("pagination_buttons")[0].appendChild(page1);
  page1.onclick = function () {
    current_page = parseInt(this.innerHTML)
    updatePage(cards, current_page);
  }
  for (let i = 0; i < cards.length; i++) {
    if (i > 1 && i % MAX_PER_PAGE === 0) {
      page_assigner++;
      var page = document.createElement("div");
      page.classList.add("page_button");
      page.innerHTML = page_assigner
      document.getElementsByClassName("pagination_buttons")[0].appendChild(page)
      page.onclick = function () {
        current_page = parseInt(this.innerHTML);
        updatePage(cards, current_page);
      }
    }
    cards[i].dataset.page = page_assigner
  }
  var buttons = document.getElementsByClassName("page_button");
  updatePage(cards, current_page);
  for (let i = 0; i < buttons.length; i++) {
    buttons[i].addEventListener("click", function () {
      for (let j = 0; j < buttons.length; j++) {
        buttons[j].classList.remove("active_page");
      }
      this.classList.add("active_page")
      console.log('ok')
    })
  }
}



function updatePage(cards, p) {
  for (let i = 0; i < cards.length; i++) {
    if (parseInt(cards[i].dataset.page) === p && !cards[i].classList.contains("hidden-by-filter")) {
      cards[i].style.display = "inline-block"
    }
    else {
      cards[i].style.display = "none"
    }
  }
}


