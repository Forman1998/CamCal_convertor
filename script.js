const dropArea = document.querySelector('.drop-section')
const listSection = document.querySelector('.list-section')
const listContainer = document.querySelector('.list')
const fileSelector = document.querySelector('.file-selector')
const fileSelectorInput = document.querySelector('.file-selector-input')
var file_number = 0;
// upload files with browse button
fileSelector.onclick = () => fileSelectorInput.click()
fileSelectorInput.onchange = () => {
    [...fileSelectorInput.files].forEach((file) => {
        if(typeValidation(file.type)){
            uploadFile(file)
        }
    })
}

// check the file type
function typeValidation(type){
    if(type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'){
        return true
    }
}

// when file is over the drag area
dropArea.ondragover = (e) => {
    e.preventDefault();
    [...e.dataTransfer.items].forEach((item) => {
        if(typeValidation(item.type)){
            dropArea.classList.add('drag-over-effect')
        }
    })
}
// when file leave the drag area
dropArea.ondragleave = () => {
    dropArea.classList.remove('drag-over-effect')
}
// when file drop on the drag area
dropArea.ondrop = (e) => {
    e.preventDefault();
    dropArea.classList.remove('drag-over-effect')
    if(e.dataTransfer.items){
        [...e.dataTransfer.items].forEach((item) => {
            if(item.kind === 'file'){
                const file = item.getAsFile();
                if(typeValidation(file.type)){
                    uploadFile(file)
                }
            }
        })
    }else{
        [...e.dataTransfer.files].forEach((file) => {
            if(typeValidation(file.type)){
                uploadFile(file)
            }
        })
    }
}
function processExcel(data) {
    var workbook = XLSX.read(data, {
      type: 'binary'
    });
  
    var firstSheet = workbook.SheetNames[0];
    var data = to_json(workbook);
    return data
};
function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function(sheetName) {
      var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1
      });
      if (roa.length) result[sheetName] = roa;
    });
    return JSON.stringify(result, 2, 2);
  };
// upload file function
function uploadFile(file){
    file_number++;
    listSection.style.display = 'block'
    var li = document.createElement('li')
    li.classList.add('in-prog')
    li.innerHTML = `
        <div class="col">
            <img src="icon/excel.png" alt="">
            <a id="csv_download${file_number}">Download CSV</a>
            <a id="error_text_download${file_number}">Error Text File</a>
        </div>
        <div class="col">
            <div class="file-name">
                <div class="name">${file.name}</div>
                <span>0%</span>
            </div>
            <div class="file-progress">
                <span></span>
            </div>
            <div class="file-size">${(file.size/(1024*1024)).toFixed(2)} MB</div>
        </div>
        <div class="col">
            <svg xmlns="http://www.w3.org/2000/svg" class="cross" height="20" width="20"><path d="m5.979 14.917-.854-.896 4-4.021-4-4.062.854-.896 4.042 4.062 4-4.062.854.896-4 4.062 4 4.021-.854.896-4-4.063Z"/></svg>
            <svg xmlns="http://www.w3.org/2000/svg" class="tick" height="20" width="20"><path d="m8.229 14.438-3.896-3.917 1.438-1.438 2.458 2.459 6-6L15.667 7Z"/></svg>
        </div>
    `
    listContainer.prepend(li);
    change_percentage(10,li);
    let csvData = 'Input, Output';
    csvData += "\r\n";
    let txtData = 'Row, Column';
    txtData += "\r\n";
    if (file) {
        change_percentage(20,li);
        var r = new FileReader();
        r.onload = e => {
          //var contents = processExcel(e.target.result);
            var workbook = XLSX.read(e.target.result, {
                type: 'binary'
            });
            var result = {};
            var i = 1;
            workbook.SheetNames.forEach(function(sheetName) {
                var roa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
                header: 1
                });
                if (roa.length) result[i] = roa;
                i++;
            });
            //var contents = JSON.stringify(result, 2, 2);
            change_percentage(30,li);
            var contents = result[1];
            var devices = contents[i].length - 3;
            console.log(devices);
            for(var j = 0; j<devices; j++)
            {
                for(var i = 0; i<contents.length; i++)
                {
                    var entry = contents[i];
                
                    if(!isNaN(entry[0]))
                    {
                        var output = entry[1];
                        output = output.replace(/A$/, '');
                        output = parseFloat(output);
                        if(!isNaN(entry[j+3]))
                        {
                            csvData += (output + output*parseFloat(entry[j+3])*100).toString() + ',' + entry[1].toString() + '';
                            csvData += "\r\n"; 
                        }   
                        else
                        {
                            if(typeof entry[j+3] == "undefined")
                            {
                                console.log(0);
                                csvData += (output + output*parseFloat(0)*100).toString() + ',' + entry[1].toString() + '';
                                csvData += "\r\n"; 
                            }
                            else{
                                entry[j+3] = entry[j+3].replace(/[^0-9$.-]/g, '');
                                var val = parseFloat(entry[j+3]);
                                if(val<=50 && val>=-50)
                                {
                                    csvData += (output + output*val).toString() + ',' + entry[1].toString() + '';
                                    csvData += "\r\n"; 
                                }
                                else
                                {
                                    txtData += (i+1).toString() +','+ (j+4).toString();
                                    txtData += "\r\n";
                                }
                            } 

                        }               
                    } 
                }
            }
            change_percentage(70,li);
            let anchor = document.getElementById('csv_download'+file_number);
            change_percentage(80,li);
            anchor.href = 'data:text/csv;charset=utf-8,' + encodeURI(csvData);
            anchor.target = '_blank';
            anchor.download = 'data.csv';
            let txtanchor = document.getElementById('error_text_download'+file_number);
            change_percentage(80,li);
            txtanchor.href = 'data:text/csv;charset=utf-8,' + encodeURI(txtData);
            txtanchor.target = '_blank';
            txtanchor.download = 'errordata.txt';
        }
        r.readAsBinaryString(file);
      } else {
        console.log("Failed to load file");
    }
    change_percentage(100, li);
    li.classList.add('complete')
    li.classList.remove('in-prog')
}
function change_percentage(percent, li){
    //var percent_complete = (e.loaded / e.total)*100
    li.querySelectorAll('span')[0].innerHTML = Math.round(percent) + '%';
    li.querySelectorAll('span')[1].style.width = percent + '%';
}