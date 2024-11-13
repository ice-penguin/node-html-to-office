var HtmlDocx = require('html-docx-js');
var fs = require('fs');
const generateMarkup = (markupText, { headerTitle }) => {
    const htmlStr = 
    `<html>
    <head>
    <style>
    table {
        border-collapse: collapse;
    }  
    table, td, th {
        border: 1px solid black;
    }
    </style>
        <style type="text/css">
        @page Section1 {
            margin:0.75in 0.75in 0.75in 0.75in;
            size:841.7pt 595.45pt;
            mso-page-orientation:landscape;
            mso-header-margin:0.5in;
            mso-header: h1;
            mso-footer-margin:0.5in;
            mso-footer: f1;
        }
    
        div.Section1 {page:Section1;}
    
        p.headerFooter { margin:0in; text-align: center; }
        </style>
    </head>
    <body><div class=Section1>
    
    
    <!-- header/footer:
        This element will appears in your main document (unless you save in a separate HTML),
        therefore, we move it off the page (left 50 inches) and relegate its height
        to 1pt by using a table with 1 exact-height row
    -->
    
    <p align=center style=\"font-weight: bold;font-size: xx-large;\">
        ${headerTitle}
    </p>
        ${markupText}
    <!-- Here's a page break:
    <br clear=all style='mso-special-character:line-break; page-break-before:always'>
    This is page 2 -->
    
    </div></body>
    </html>`

    return htmlStr;
}

const htmlStrToDocx = async (markupText, { headerTitle, outputFile }) => {
    return new Promise((resolve, reject) => {
        const inputMarkup = generateMarkup(markupText, { headerTitle })
        var outputFileName = `${outputFile || 'outFile'}.docx`;
        var docx = HtmlDocx.asBlob(inputMarkup);
        fs.writeFile(outputFileName, docx, function (err) {
            if (err) {
                reject(err)
                return 0
            };
            resolve(`${outputFileName}`)
        });
    })

}

module.exports = {
    htmlStrToDocx
}