function $(id) {
  return document.getElementById(id);
}

const report = {
  log: (data) => {
    if (!$("log")) {
      document.body.appendChild(createElement("div", { id: "log", style: "background: red" }));
    }
    $("log").innerHTML += "<br/>" + new Date().toLocaleTimeString() + ": " + JSON.stringify(data);
  },
  error: (data) => {
    if (!$("log")) {
      document.body.appendChild(createElement("div", { id: "log", style: "background: red" }));
    }
    $("log").innerHTML +=
      "<br/>" + new Date().toLocaleTimeString() + ": " + "<p style='color: red'>" + JSON.stringify(data) + "</p>";
  },
};
