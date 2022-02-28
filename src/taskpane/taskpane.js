document.getElementById("translateCsEn").addEventListener("click", translateSelection)

function translateSelection() {
	Word.run((context) => {
		const doc = context.document
		let originalRange = doc.getSelection()
		originalRange.load("text")

		return context.sync()
			.then(() => {
				sendToLindatAPI("cs", "en", originalRange.text)
					.then(response => response.text())
					.then(data => {
						data = data.replace(/\s+$/g, '') // replace newline on end
						doc.body.insertParagraph(data, "End")
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
	const url = "https://lindat.mff.cuni.cz/services/translation/api/v2/languages/"

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
		headers,
		body: formBody
	})

	return response
}