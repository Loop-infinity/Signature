var cssStyleSheet = "";

$(document).ready(function () {
    const myForm = document.getElementById("myForm");
    const csvFile = document.getElementById("csvFile");

    myForm.addEventListener("submit", function (e) {
      e.preventDefault();
      const input = csvFile.files[0];
      parseExcel(input);
    });

    $.when($.get("/css/signature.css"))
            .done(function(response) {
                cssStyleSheet = response;
            });
});
    

function parseExcel(file){
    var reader = new FileReader();

    reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
            type: 'binary'
        });

        workbook.SheetNames.forEach(function(sheetName) {
            var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            var json_object = JSON.stringify(XL_row_object);
            //console.log(json_object);
            generateSignatures(XL_row_object);

        });

    }

    reader.onerror = function(ex) {
        console.log(ex);
    };

    reader.readAsBinaryString(file);
};

function generateSignatures(data){
    var template = $('.container').clone();
    console.log(data);
    //$('body').append(`<div class="flex-container"></div>`);
    data.forEach(employee => {
        var clone = template.clone();
        console.log(employee)
        clone.find('.name').html(employee.Name);
        clone.find('.designation').html(employee.Designation);

        clone.find('.address').html(employee.Address);

        clone.find('.tel').html(employee.MobileNumber);
        clone.find('.mob').html(employee.TelephoneNumber);
        clone.find('.email').html(employee.Email);

        clone.appendTo("#output");
        generateHtmlCode(clone);
    });

    //$(".container").after("<br/> <br/> <br/> <br>");
}

function generateHtmlCode(element){
    var signatureHtml = element.wrap('<p/>').parent().html();

    var htmlCode = `
    <html>
    <head>
        <style>
            ${cssStyleSheet}
        </style>
    </head>
    <body>
        ${signatureHtml}
    </body>
    </html>
    `;
    var htmlCodeEditor = `
    <div class="">
        <textarea class="code-preview" name="" id="" cols="50" rows="10">${htmlCode}</textarea>
    </div>
    `;
    $('#output').append(htmlCodeEditor);

}