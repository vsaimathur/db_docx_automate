var fs = require("fs");
var zip = require('adm-zip');

//paths to input and output files. (if not specified here, default -> input.docx & output->docx)
const pathToInputFile = undefined;
const pathToOutputFile = undefined;


let dbObject = {
    name: "Mathur",
    age: 20,
    covidData: [
        [
            "Country", "Active Cases", "Recovered Cases"
        ],
        [
            "India", "123", "456"
        ],
        [
            "America", "789", "159"
        ],
        [
            "Japan", "1012", "5478"
        ]
    ]
}

var zipRead = new zip(pathToInputFile || "./input.docx");
var zipCreate = new zip();

var zipEntries = zipRead.getEntries(); // an array of ZipEntry records

//global variable which stores the xml for word document. (excluding charts)
var xmlData = "";

//Storing the current xml data (document.xml file) from zip file (docx->zip rename) to xmlData var. 
//Also preserving the root directory structure. (we would write all those along with docuement.xml which we modify later at last) 
zipEntries.forEach(function (zipEntry) {
    if (zipEntry.entryName == "word/document.xml") {
        xmlData = zipEntry.getData().toString('utf8');
    } else if (zipEntry.entryName == "word/charts/chart1.xml") {
        xmlPieData = zipEntry.getData().toString('utf-8');
    } else {
        zipCreate.addFile(zipEntry.entryName, zipEntry.getData(), "adding remaining directory structure");
    }
});

//gets the xml for header in table.
const getTableHeadXmlData = (rowData) => {
    return `<w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="9121" w:type="dxa"/>
    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
    </w:tblPr>
<w:tblGrid>
  <w:gridCol w:w="3040"/>
  <w:gridCol w:w="3040"/>
  <w:gridCol w:w="3041"/>
</w:tblGrid> ${getTableRowXmlData(rowData)}`;
}

//get the xml for column in table.
const getTableColXmlData = (colDataVal) => {
    return ` <w:tc>
    <w:tcPr>
      <w:tcW w:w="3040" w:type="dxa"/>
    </w:tcPr>
    <w:p w14:paraId="788BB71B" w14:textId="6F68C6C7" w:rsidR="003B3DAE" w:rsidRDefault="003B3DAE">
      <w:pPr>
        <w:rPr>
          <w:lang w:val="en-US"/>
        </w:rPr>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:lang w:val="en-US"/>
        </w:rPr>
        <w:t>${colDataVal}</w:t>
      </w:r>
    </w:p>
  </w:tc>`;
}

//gets xml for row in table.
const getTableRowXmlData = (rowData) => {
    const rowDataWrapper = `<w:tr w:rsidR="003B3DAE" w14:paraId="4F6A71F7" w14:textId="77777777" w:rsidTr="00772F49">
    <w:trPr>
      <w:trHeight w:val="511"/>
    </w:trPr>`;
    let rowXmlData = rowDataWrapper;
    for (let i = 0; i < rowData.length; i++) {
        rowXmlData += getTableColXmlData(rowData[i]);
    }
    rowXmlData += "</w:tr>";
    return rowXmlData;
}

//gets the entire xml for table.
const getTableXmlData = (tableData) => {
    if (tableData == undefined) return "";
    let tableXmlData = "<w:tbl>";
    tableXmlData += getTableHeadXmlData(tableData[0]);
    for (let i = 1; i < tableData.length; i++) {
        tableXmlData += getTableRowXmlData(tableData[i]);
    }
    tableXmlData += "</w:tbl>";
    return tableXmlData;
}

// function for replacing the template literals in xml({{example}})
const replaceTemplateLiterals = () => {

    for (let i = xmlData.indexOf("{{"); i >= 0; i = xmlData.indexOf("{{", i + 1)) {
        let templateStart = i;
        //finding template end position.
        let templateEnd = xmlData.indexOf("}}", i + 1);
        let fetchedkeyName = xmlData.substr(templateStart + 2, templateEnd - templateStart - 2);
        console.log(fetchedkeyName,i,templateEnd);
        xmlData = xmlData.replace(`{{${fetchedkeyName}}}`, dbObject[fetchedkeyName].toString());
    }
    // console.log(xmlData);
}

//function which inserts table after a some word or paragraph in docx file.
const insertTable = (tableData) => {

    //assuming that last would be paragraph.
    let paragraphEndLastIndex = xmlData.lastIndexOf("</w:p>");
    // some basic insertion of new data at start for table
    // tableData = tableData.splice(1,0,["Europe","4561","854"]);

    //generating entire table dynamically here.
    xmlData = xmlData.slice(0, paragraphEndLastIndex + 6) + getTableXmlData(tableData) + xmlData.slice(paragraphEndLastIndex + 6);
}

//Also wrote logic for fetching the data from Column Chart in word. (didn't use it for now, but can be used)
const getPieData = () => {
    for (let i = xmlPieData.indexOf("<c:ptCount"); i >= 0; i = xmlPieData.indexOf("<c:ptCount", i + 1)) {
        let pntCountStart = i;
        //finding template end position.
        let pntCountEnd = xmlPieData.indexOf("<c:pt ", i + 1) - 1;
        let pntCount = parseInt(xmlPieData.substr(pntCountStart + 16, pntCountEnd - pntCountStart - 18));
        console.log(pntCount);
        // console.log(fetchedkeyName,i,pntCountEnd);
        for (let j = xmlPieData.indexOf("<c:v>", pntCountEnd + 1); pntCount && j >= 0; j = xmlPieData.indexOf("<c:v", j + 1), pntCount--) {
            let cvCountStart = j;
            let cvCountEnd = xmlPieData.indexOf("</c:v>", j + 1);
            let pntVal = xmlPieData.substr(cvCountStart + 5, cvCountEnd - cvCountStart - 5);
            console.log(pntVal);
        }
    }
}

//insertions in document.xml file in zip extracted from word file.
replaceTemplateLiterals();
insertTable(dbObject.covidData);

//writing all remaining files which are not included in the root structure while reading Data from zip. 

//adding main document.xml file to root structure
zipCreate.addFile("word/document.xml", Buffer.from(xmlData), "adding main word/document.xml file");

//outputting all the files to zip and finally to output.docx
zipCreate.writeZip(pathToOutputFile || "./output.docx");
console.log("Outputted the Data ✔");