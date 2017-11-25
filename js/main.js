document.addEventListener('DOMContentLoaded', (e) => {
	class DropZone {
		constructor(zoneId) {
			this.Id = zoneId;
			this.Control = document.getElementById(this.Id);
			this.bindEvents();
			this.changeTableColor();
			document.getElementById('with-table').style.display = 'none';
			document.getElementById('drop-zone').style.display = '';
		}

		bindEvents() {
			this.Control.addEventListener('dragenter', (e) => this.handleDragover(e), false);
			this.Control.addEventListener('dragover', (e) => this.handleDragover(e), false);
			this.Control.addEventListener('drop', (e) => this.handleDrop(e), false);
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
			document.getElementById('with-table').style.display = '';
			document.getElementById('drop-zone').style.display = 'none';
			e.stopPropagation();
			e.preventDefault();
			var files = e.dataTransfer.files;
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
					trTemplate += `<td ${index == 0 ? "class='number'" : ""}>${cell}</td>`;
				});
				trTemplate += '</tr>';
				tbody.insertAdjacentHTML('beforeend', trTemplate);
			});

			this.exportTableAsImg();
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

		changeTableColor() {
			document.getElementById('color-picker').addEventListener('change', (e) => {
				var element = e.currentTarget;
				document.documentElement.style.setProperty('--tdColor', element.value);
			});
		}
	}

	var dropZone = new DropZone('drop-zone');
});