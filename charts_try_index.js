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
var xmlChartData = "";

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
const getChartData = () => {
    for (let i = xmlChartData.indexOf("<c:ptCount"); i >= 0; i = xmlChartData.indexOf("<c:ptCount", i + 1)) {
        let pntCountStart = i;
        //finding template end position.
        let pntCountEnd = xmlChartData.indexOf("<c:pt ", i + 1) - 1;
        let pntCount = parseInt(xmlChartData.substr(pntCountStart + 16, pntCountEnd - pntCountStart - 18));
        console.log(pntCount);
        // console.log(fetchedkeyName,i,pntCountEnd);
        for (let j = xmlChartData.indexOf("<c:v>", pntCountEnd + 1); pntCount && j >= 0; j = xmlChartData.indexOf("<c:v", j + 1), pntCount--) {
            let cvCountStart = j;
            let cvCountEnd = xmlChartData.indexOf("</c:v>", j + 1);
            let pntVal = xmlChartData.substr(cvCountStart + 5, cvCountEnd - cvCountStart - 5);
            console.log(pntVal);
        }
    }
}


//getting the parameters like (coutries like india,america,japan in our db example).
const getChartParameterXmlData = (chartParameterData) => {
    chartParameterXmlData = `<c:cat>
    <c:strRef>
    <c:strCache>`;
    chartParameterXmlData += `<c:ptCount val="${chartParameterData.length}"/>`;
    for (let i = 0; i < chartParameterData.length; i++) {
        chartParameterXmlData += `<c:pt idx="${i}">
          <c:v>${chartParameterData[i]}</c:v>
        </c:pt>`;
    }
    chartParameterXmlData += `</c:strCache></c:strRef>
    </c:cat>`;
    return chartParameterXmlData;
}

//getting the parameter values (like numbers in charts) for charts.
const getChartParameterValueXmlData = (parameterValueData) => {
    let chartParameterValueXmlData = `<c:val>
  <c:numRef>
    <c:numCache>
      <c:formatCode>General</c:formatCode>
      `;
    chartParameterValueXmlData += `<c:ptCount val="${parameterValueData.length}"/>`;
    for (let i = 0; i < parameterValueData.length; i++) {
        chartParameterValueXmlData += `<c:pt idx="${i}">
        <c:v>${parameterValueData[i]}</c:v>
      </c:pt>`
    }
    chartParameterValueXmlData += `</c:numCache>
    </c:numRef>
  </c:val>`;
    return chartParameterValueXmlData;
}

//generating dynamic xml data for charts.
const getChartMainXmlData = (chartData) => {
    let chartParameterData = chartData.map((rowData) => {
        return rowData[0];
    }).slice(1);

    let chartMainXmlData = "";
    for (let i = 0; i < chartData[0].length - 1; i++) {
        chartMainXmlData += `<c:ser>
       <c:idx val="${i}"/>
       <c:order val="${i}"/>
       <c:tx>
         <c:strRef>
           <c:strCache>
             <c:ptCount val="1"/>
             <c:pt idx="0">
               <c:v>${chartData[0][i + 1]}</c:v>
             </c:pt>
           </c:strCache>
         </c:strRef>
       </c:tx>
       <c:spPr>
         <a:solidFill>
           <a:schemeClr val="${`accent${i + 1}`}"/>
         </a:solidFill>
         <a:ln>
           <a:noFill/>
         </a:ln>
         <a:effectLst/>
       </c:spPr>
       <c:invertIfNegative val="0"/>` +
            getChartParameterXmlData(chartParameterData) +
            getChartParameterValueXmlData(chartData.map((chartDataRowVal) => chartDataRowVal[i + 1]).slice(1)) +
            `<c:extLst>
         <c:ext xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart" uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}">
           <c16:uniqueId val="{00000000-772C-4B68-80C0-D7AFABE73807}"/>
         </c:ext>
       </c:extLst>
       </c:ser>`
    }
    // console.log(chartMainXmlData);
    return chartMainXmlData;
}

//appending the static xml header and footer data to our dynamically generated xml data for charts.
const getChartXmlData = (chartData) => {
    let chartMainXmlData = getChartMainXmlData(chartData);
    let staticHeaderChartXmlData = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">
      <c:date1904 val="0"/>
      <c:lang val="en-US"/>
      <c:roundedCorners val="0"/>
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
          <c14:style val="102"/>
        </mc:Choice>
        <mc:Fallback>
          <c:style val="2"/>
        </mc:Fallback>
      </mc:AlternateContent>
      <c:chart>
        <c:title>
          <c:overlay val="0"/>
          <c:spPr>
            <a:noFill/>
            <a:ln>
              <a:noFill/>
            </a:ln>
            <a:effectLst/>
          </c:spPr>
          <c:txPr>
            <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
                  <a:solidFill>
                    <a:schemeClr val="tx1">
                      <a:lumMod val="65000"/>
                      <a:lumOff val="35000"/>
                    </a:schemeClr>
                  </a:solidFill>
                  <a:latin typeface="+mn-lt"/>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </a:pPr>
              <a:endParaRPr lang="en-US"/>
            </a:p>
          </c:txPr>
        </c:title>
        <c:autoTitleDeleted val="0"/>
        <c:plotArea>
          <c:layout/>
          <c:chartChart>
            <c:chartDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>`;

    let staticFooterChartXmlData = `<c:dLbls>
    <c:showLegendKey val="0"/>
    <c:showVal val="0"/>
    <c:showCatName val="0"/>
    <c:showSerName val="0"/>
    <c:showPercent val="0"/>
    <c:showBubbleSize val="0"/>
  </c:dLbls>
  <c:gapWidth val="219"/>
  <c:overlap val="-27"/>
  <c:axId val="649153104"/>
  <c:axId val="649153520"/>
</c:chartChart>
<c:catAx>
  <c:axId val="649153104"/>
  <c:scaling>
    <c:orientation val="minMax"/>
  </c:scaling>
  <c:delete val="0"/>
  <c:axPos val="b"/>
  <c:numFmt formatCode="General" sourceLinked="1"/>
  <c:majorTickMark val="none"/>
  <c:minorTickMark val="none"/>
  <c:tickLblPos val="nextTo"/>
  <c:spPr>
    <a:noFill/>
    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
      <a:solidFill>
        <a:schemeClr val="tx1">
          <a:lumMod val="15000"/>
          <a:lumOff val="85000"/>
        </a:schemeClr>
      </a:solidFill>
      <a:round/>
    </a:ln>
    <a:effectLst/>
  </c:spPr>
  <c:txPr>
    <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="65000"/>
              <a:lumOff val="35000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:latin typeface="+mn-lt"/>
          <a:ea typeface="+mn-ea"/>
          <a:cs typeface="+mn-cs"/>
        </a:defRPr>
      </a:pPr>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </c:txPr>
  <c:crossAx val="649153520"/>
  <c:crosses val="autoZero"/>
  <c:auto val="1"/>
  <c:lblAlgn val="ctr"/>
  <c:lblOffset val="100"/>
  <c:noMultiLvlLbl val="0"/>
</c:catAx>
<c:valAx>
  <c:axId val="649153520"/>
  <c:scaling>
    <c:orientation val="minMax"/>
  </c:scaling>
  <c:delete val="0"/>
  <c:axPos val="l"/>
  <c:majorGridlines>
    <c:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="15000"/>
            <a:lumOff val="85000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
      <a:effectLst/>
    </c:spPr>
  </c:majorGridlines>
  <c:numFmt formatCode="General" sourceLinked="1"/>
  <c:majorTickMark val="none"/>
  <c:minorTickMark val="none"/>
  <c:tickLblPos val="nextTo"/>
  <c:spPr>
    <a:noFill/>
    <a:ln>
      <a:noFill/>
    </a:ln>
    <a:effectLst/>
  </c:spPr>
  <c:txPr>
    <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr>
        <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="65000"/>
              <a:lumOff val="35000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:latin typeface="+mn-lt"/>
          <a:ea typeface="+mn-ea"/>
          <a:cs typeface="+mn-cs"/>
        </a:defRPr>
      </a:pPr>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </c:txPr>
  <c:crossAx val="649153104"/>
  <c:crosses val="autoZero"/>
  <c:crossBetween val="between"/>
</c:valAx>
<c:spPr>
  <a:noFill/>
  <a:ln>
    <a:noFill/>
  </a:ln>
  <a:effectLst/>
</c:spPr>
</c:plotArea>
<c:legend>
<c:legendPos val="b"/>
<c:overlay val="0"/>
<c:spPr>
  <a:noFill/>
  <a:ln>
    <a:noFill/>
  </a:ln>
  <a:effectLst/>
</c:spPr>
<c:txPr>
  <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr>
      <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="65000"/>
            <a:lumOff val="35000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:latin typeface="+mn-lt"/>
        <a:ea typeface="+mn-ea"/>
        <a:cs typeface="+mn-cs"/>
      </a:defRPr>
    </a:pPr>
    <a:endParaRPr lang="en-US"/>
  </a:p>
</c:txPr>
</c:legend>
<c:plotVisOnly val="1"/>
<c:dispBlanksAs val="gap"/>
<c:extLst>
<c:ext xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart" uri="{56B9EC1D-385E-4148-901F-78D8002777C0}">
  <c16r3:dataDisplayOptions16>
    <c16r3:dispNaAsBlank val="1"/>
  </c16r3:dataDisplayOptions16>
</c:ext>
</c:extLst>
<c:showDLblsOverMax val="0"/>
</c:chart>
<c:spPr>
<a:solidFill>
<a:schemeClr val="bg1"/>
</a:solidFill>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
  <a:schemeClr val="tx1">
    <a:lumMod val="15000"/>
    <a:lumOff val="85000"/>
  </a:schemeClr>
</a:solidFill>
<a:round/>
</a:ln>
<a:effectLst/>
</c:spPr>
<c:txPr>
<a:bodyPr/>
<a:lstStyle/>
<a:p>
<a:pPr>
  <a:defRPr/>
</a:pPr>
<a:endParaRPr lang="en-US"/>
</a:p>
</c:txPr>
<c:externalData r:id="rId3">
<c:autoUpdate val="0"/>
</c:externalData>
</c:chartSpace>
`;

    return staticHeaderChartXmlData + chartMainXmlData + staticFooterChartXmlData;
}

//xml data for word/charts/colors1.xml file.
const getChartColorsXmlData = () => {
    let chartChartColorXmlData = `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/><a:schemeClr val="accent4"/><a:schemeClr val="accent5"/><a:schemeClr val="accent6"/><cs:variation/><cs:variation><a:lumMod val="60000"/></cs:variation><cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation><cs:variation><a:lumMod val="80000"/></cs:variation><cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation><cs:variation><a:lumMod val="50000"/></cs:variation><cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation><cs:variation><a:lumMod val="70000"/></cs:variation><cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation></cs:colorStyle>`;
    return chartChartColorXmlData;
}

//xml data for word/charts/style1.xml file
const getChartStyleXmlData = () => {
    let chartChartStyleXmlData = `<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201"><cs:axisTitle><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="1000" kern="1200"/></cs:axisTitle><cs:categoryAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="900" kern="1200"/></cs:categoryAxis><cs:chartArea mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="1000" kern="1200"/></cs:chartArea><cs:dataLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900" kern="1200"/></cs:dataLabel><cs:dataLabelCallout><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz="900" kern="1200"/><cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1"><a:spAutoFit/></cs:bodyPr></cs:dataLabelCallout><cs:dataPoint><cs:lnRef idx="0"/><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:dataPoint><cs:dataPoint3D><cs:lnRef idx="0"/><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:dataPoint3D><cs:dataPointLine><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="1"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointLine><cs:dataPointMarker><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></cs:spPr></cs:dataPointMarker><cs:dataPointMarkerLayout symbol="circle" size="5"/><cs:dataPointWireframe><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="1"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointWireframe><cs:dataTable><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="900" kern="1200"/></cs:dataTable><cs:downChart><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:downChart><cs:dropLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:dropLine><cs:errorChart><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:errorChart><cs:floor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:noFill/><a:ln><a:noFill/></a:ln></cs:spPr></cs:floor><cs:gridlineMajor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMajor><cs:gridlineMinor><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="5000"/><a:lumOff val="95000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMinor><cs:hiLoLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:hiLoLine><cs:leaderLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:leaderLine><cs:legend><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900" kern="1200"/></cs:legend><cs:plotArea mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea><cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride"><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:plotArea3D><cs:seriesAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900" kern="1200"/></cs:seriesAxis><cs:seriesLine><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:seriesLine><cs:title><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="1400" b="0" kern="1200" spc="0" baseline="0"/></cs:title><cs:trendline><cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="sysDot"/></a:ln></cs:spPr></cs:trendline><cs:trendlineLabel><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900" kern="1200"/></cs:trendlineLabel><cs:upChart><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="dk1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:upChart><cs:valueAxis><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:defRPr sz="900" kern="1200"/></cs:valueAxis><cs:wall><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:noFill/><a:ln><a:noFill/></a:ln></cs:spPr></cs:wall></cs:chartStyle>`;
    return chartChartStyleXmlData;
}

//xml data for word/charts/_rels/chart1.xml.rels file.
const getChartEmbedRelXmlData = () => {
    let chartChartEmbedRelXmlData = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/></Relationships>`;
    return chartChartEmbedRelXmlData;
}

//appending chart xml data in document.xml file using xmlData global variable.
const insertChart = () => {
    //including charts info in main document.xml file.
    let columnChartXmlData = `<w:p w14:paraId="09322F03" w14:textId="39FD51F5" w:rsidR="00B50E16" w:rsidRPr="000A6FB2" w:rsidRDefault="00B50E16" w:rsidP="000A6FB2">
<w:pPr>
  <w:rPr>
    <w:lang w:val="en-US"/>
  </w:rPr>
</w:pPr>
<w:r>
  <w:rPr>
    <w:noProof/>
    <w:lang w:val="en-US"/>
  </w:rPr>
  <w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="463E05B6" wp14:editId="546DEB8F">
      <wp:extent cx="5486400" cy="3200400"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="1" name="Chart 1"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId4"/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>
</w:p>`;
    //assuming insertion of chart after table.
    let lastTableIndex = xmlData.lastIndexOf("</w:tbl>");
    xmlData = xmlData.slice(0, lastTableIndex + 8) + columnChartXmlData + xmlData.slice(lastTableIndex + 8);
}

//insertions in document.xml file in zip extracted from word file.
replaceTemplateLiterals();
insertTable(dbObject.covidData);
insertChart(dbObject.covidData);

console.log(xmlData);

//writing all remaining files which are not included in the root structure while reading Data from zip. 

//adding main document.xml file to root structure
zipCreate.addFile("word/document.xml", Buffer.from(xmlData), "adding main word/document.xml file");

//adding chart1.xml for chart xml data to its location.
zipCreate.addFile("word/charts/chart1.xml", Buffer.from(getChartXmlData(dbObject.covidData)), "adding chart data xml file");
//adding style1.xml for style xml data used in charts to its location.
zipCreate.addFile("word/charts/style1.xml", Buffer.from(getChartStyleXmlData()), "adding chart styles xml file");
//adding colors1.xml for colors used by charts to its location in root structure.
zipCreate.addFile("word/charts/colors1.xml", Buffer.from(getChartColorsXmlData()), "adding chart colors xml file");

//adding chart1.xml.refs file which creats a relationship b/w the chart xml data and excel in embeddings folder.
zipCreate.addFile("word/charts/_rels/chart1.xml.rels", Buffer.from(getChartEmbedRelXmlData()), "adding chart embed xml file");

//outputting all the files to zip and finally to output.docx
zipCreate.writeZip(pathToOutputFile || "./output.docx");