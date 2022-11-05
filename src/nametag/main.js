Office.onReady(() => {
  // If needed, Office.js is ready to be called
  officeReady = true;
});

let activeModel = null;
let officeReady = false;
let inputType = "selection";

window.onload = () => {
  loadModels();
  document.getElementById("modelSelector").addEventListener("change", (e) => {
    activeModel = e.target.value;
  });
  document.getElementById("recognizeButton").addEventListener("click", recognize);
  document.getElementById("userSelectionInputType").addEventListener("click", () => {
    inputType = "selection";
    document.getElementById("inputFieldContainer").style.display = "none";
    document.getElementById("selectionContainer").style.display = "block";
    document.getElementById("errorMessage").innerText = "Please select some text first.";
  });
  document.getElementById("inputFieldInputType").addEventListener("click", () => {
    inputType = "field";
    document.getElementById("inputFieldContainer").style.display = "block";
    document.getElementById("selectionContainer").style.display = "none";
    document.getElementById("errorMessage").innerText = "Please insert some text first.";
  });
};

function loadModels() {
  const modelsUrl = "https://lindat.mff.cuni.cz/services/nametag/api/models";
  fetch(modelsUrl, {
    method: "GET",
  })
    .then((response) => response.json())
    .then((data) => {
      if (data?.length === 0) return;
      const activeModels = Object.keys(data.models);
      activeModel = activeModels
        .filter((item) => item.includes("czech"))
        .filter((item) => !item.includes("no_numbers"))
        .map((item) => {
          const a = item.split("-");
          return [item, a[a.length - 1]];
        })
        .sort((a, b) => b[1] - a[1])?.[0]?.[0];

      if (!activeModel) {
        activeModel = Object.keys(data.models)[0];
      }
      document.getElementById("modelSelector").innerText = "";
      activeModels.forEach((item, index) => {
        const newOption = document
          .getElementById("modelSelector")
          .appendChild(document.createElement("option", { key: index, value: item, selected: activeModel === item }));
        newOption.appendChild(document.createTextNode(item));
      });
    })
    .catch((error) => {
      // todo handle no internet
      console.error(error);
    });
}

function recognize() {
  Word.run((context) => {
    const doc = context.document;
    let originalRange = doc.getSelection();
    originalRange.load("text");

    return context.sync().then(() => {
      const inputData =
        inputType === "selection" ? originalRange.text : document.getElementById("inputDataTextarea").value;

      document.getElementById("errorMessage").style.display = inputData.length === 0 ? "block" : "none";

      if (inputData.length === 0) {
        return;
      }

      sendToLindatAPI(inputData)
        .then((response) => {
          console.log("response", response);
          return response.json();
        })
        .then((data) => {
          data = data.result.split("\n").map((item) => item.split("\t"));
          data.pop();

          if (inputType === "selection") {
            insertResultToTable(data);
          } else {
            insertResultToTable(data);
          }

          //data = data.replace(/\s+$/g, ""); // replace newline on end
          //for (let row of data.split("\r")) doc.body.insertParagraph(row, "End");
        })
        .then(context.sync)
        .catch((error) => {
          console.error("Error:", error);
        });
    });
  }).catch((error) => {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertResultToTable(data) {
  document.getElementById("outputTable").innerHTML = [["Range", "Entity Type", "Entity Text"], ...data]
    .map(
      (item, index) => `
  <tr>
    <${index == 0 ? "th" : "td"}>${item[0]}</${index == 0 ? "th" : "td"}>
    <${index == 0 ? "th" : "td"}>${item[1]}</${index == 0 ? "th" : "td"}>
    <${index == 0 ? "th" : "td"}>${item[2]}</${index == 0 ? "th" : "td"}>
  </tr>
`
    )
    .join("");
}

function selectionLoad() {}

function sendToLindatAPI(data) {
  let url = "https://lindat.mff.cuni.cz/services/nametag/api/recognize";

  if (activeModel == null) {
    return;
  }

  const formDataValues = {
    model: activeModel,
    data,
    input: "untokenized",
    output: "vertical",
  };

  const formData = new FormData();
  for (let property in formDataValues) {
    formData.append(property, formDataValues[property]);
  }

  const response = fetch(url, {
    method: "POST",
    body: formData,
  });

  return response;
}
