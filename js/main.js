document.addEventListener('DOMContentLoaded', (e) => {
	var colorPicker = document.getElementById('color-picker');
	var picker = new CP(colorPicker);
	picker.on("change", function(color) {
        this.target.value = '#' + color;
        var event = new Event('change');
        colorPicker.dispatchEvent(event);
	});
	document.getElementById('fontSize-picker').addEventListener('change', (e) => {
		var fontInput = e.currentTarget;
		var fontSize = fontInput.value;
		document.getElementById('quiz-table').style.fontSize = `${fontSize}px`;
	});
	class DropZone {
		constructor(zoneId) {
			this.Id = zoneId;
			this.Control = document.getElementById(this.Id);
			this.bindEvents();
			this.changeTableColor();
			document.getElementById('with-table').style.display = 'none';
			document.getElementById('settings').style.display = 'none';
			document.getElementById('select-file').style.display = '';
		}

		bindEvents() {
			this.Control.addEventListener('dragenter', (e) => this.handleDragover(e), false);
			this.Control.addEventListener('dragover', (e) => this.handleDragover(e), false);
			this.Control.addEventListener('drop', (e) => this.handleDrop(e), false);
			document.getElementById('file').addEventListener('change', (e) => {
				var input = e.currentTarget;
				var files = input.files;
				this.processFiles(files);
			});
			document.getElementById('downloadImg').addEventListener('click', (e) => {
				//this.exportTableAsImg();
				this.exportTableAsJpg();
			});
		}

		handleDragover(e) {
			e.stopPropagation();
			e.preventDefault();
			e.dataTransfer.dropEffect = 'copy';
		}

		to_json(workbook) {
			var result = {};
			workbook.SheetNames.forEach(function(sheetName) {
				var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {header:1});
				if(roa.length > 0) result[sheetName] = roa;
			});
			return result;
		}

		process_wb(wb, sheetidx) {
			var sheet = wb.SheetNames[sheetidx||0];
			var json = this.to_json(wb)[sheet];
			this.generateTableFromJson(json);
		}

		handleDrop(e) {
			e.stopPropagation();
			e.preventDefault();
			var files = e.dataTransfer.files;
			this.processFiles(files);
		}

		processFiles(files) {
			document.getElementById('with-table').style.display = '';
			document.getElementById('settings').style.display = '';
			document.getElementById('select-file').style.display = 'none';
			Array.from(files).forEach((f) => {
				var reader = new FileReader();
				var name = f.name;
				reader.onload = (e) => {
					var data = e.target.result;
					var wb;
					var readtype = {type: 'binary'};
					try {
						wb = XLSX.read(data, readtype);
						this.process_wb(wb);
					} catch(e) { console.log(e); }
				};
				reader.readAsBinaryString(f);
			});
		}

		generateTableFromJson(json) {
			var thead = document.querySelector('#quiz-table thead tr');
			var tbody = document.querySelector('#quiz-table tbody');
			// generate header
			var headerRow = json.shift();
			headerRow.forEach((row) => {
				if (row) {
					let th = `<th>${row}</th>`;
					thead.insertAdjacentHTML('beforeend', th);
				}
			});

			// generate another rows
			json.forEach((row) => {
				let trTemplate = '<tr>';
				row.forEach((cell, index) => {
					let tdClass = "";
					switch(index) {
						case 0:
							tdClass = "class='number'";
							break;
						case 1:
							tdClass = "class='taleft'";
							break;
						default:
							tdClass = "";
							break;
					}
					if (cell && cell.trim()) {
						trTemplate += `<td ${tdClass}>${cell}</td>`;
					}
				});
				trTemplate += '</tr>';
				tbody.insertAdjacentHTML('beforeend', trTemplate);
			});

			if (json.length > 12) {
				this.setAdditionalStyle();
			}
		}

		exportTableAsImg() {
			var node = document.getElementById('with-table');
		    domtoimage.toBlob(document.getElementById('with-table'), {
		    	width: 1000,
		    	height: 706
		    })
			.then(function (blob) {
				window.saveAs(blob, 'quiz.png');
			});
		}

		exportTableAsJpg() {
			var node = document.getElementById('with-table');
			domtoimage.toJpeg(node, { quality: 1 })
			.then(function (dataUrl) {
				var link = document.createElement('a');
				link.download = 'quiz.jpeg';
				link.href = dataUrl;
				link.click();
			});
		}

		changeTableColor() {
			document.getElementById('color-picker').addEventListener('change', (e) => {
				var element = e.currentTarget;
				document.documentElement.style.setProperty('--tdColor', element.value);
			});
		}

		setAdditionalStyle() {
			document.getElementById('fontSize-picker').value = 16;
			document.getElementById('quiz-table').style.fontSize = '16px';
			document.querySelectorAll('#quiz-table td').forEach(x => x.style.padding = "6px");
		}
	}

	var dropZone = new DropZone('drop-zone');
});