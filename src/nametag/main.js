Office.onReady(() => {
	// If needed, Office.js is ready to be called
	document.getElementById("translateCsEn").addEventListener("click", translateSelection)

});

const console = {
	log: (data) => {
		document.getElementById("log").innerHTML += "<br/>" + (new Date()).toLocaleTimeString() + ": " + JSON.stringify(data)
	},
	error: (data) => {
		document.getElementById("log").innerHTML += "<br/>" + (new Date()).toLocaleTimeString() + ": " + "<p style='color: red'>" + JSON.stringify(data) + "</p>"
	},
	
}

function translateSelection() {
	console.log("here")
	Word.run((context) => {
		const doc = context.document
		let originalRange = doc.getSelection()
		originalRange.load("text")

		return context.sync()
			.then(() => {
				sendToLindatAPI("cs", "en", originalRange.text)
					.then(response => { console.log(response); return response.text()})
					.then(data => {
						console.log(data);
						data = data.replace(/\s+$/g, '') // replace newline on end
						for(let row of data.split('\r'))
							doc.body.insertParagraph(row, "End")
					})
					.then(context.sync)
					.catch((error) => {
						console.error('Error:', error)
					})
			})
	})
	.catch((error) => {
		console.log("Error: " + error)
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo))
		}
	})
}

function sendToLindatAPI(src, tgt, input_text) {
	let url = "https://lindat.mff.cuni.cz/services/translation/api/v2/languages/"

	// hofix becaouse LINDAT doesn have 'Access-Control-Allow-Origin' header
	//url = `https://quest.ms.mff.cuni.cz/prak/cors/${url}`

	const formData = {
		src,
		tgt,
		input_text,
	}

	let formBody = [];
	for (let property in formData) {
		const encodedKey = encodeURIComponent(property)
		const encodedValue = encodeURIComponent(formData[property])
		formBody.push(encodedKey + "=" + encodedValue)
	}
	formBody = formBody.join("&");
	
	const headers = {
		'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
	}

	const response = fetch(url, {
		method: 'POST',
		body: formBody,
		headers,
	})

	return response
}