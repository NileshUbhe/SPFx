import * as React from 'react';
import IOrgchartToolBarProps from './IOrgchartToolBarProps';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export default class OrgchartToolBar extends React.Component<IOrgchartToolBarProps, {}> {
  // constructor(props: any)
  // {
  //   super(props);
  //   this.state={
  //     zoomIn: () => this.props.zoomIn,
  //     zoomOut: () => this.props.zoomOut,
  //     resetTransform: () => this.props.resetTransform
  //   };
  // }
  exportToPDF = () => {
    const orgchart = document.getElementById('orgchart');

    // Use html2canvas to convert the org chart div to a canvas
    if (orgchart) {
      html2canvas(orgchart, { scale: 1 }).then(canvas => {
        const imgData = canvas.toDataURL('image/png'); // Convert canvas to image data
        const pdf = new jsPDF('landscape', 'mm', 'a4'); // Create a jsPDF document
        const imgProps = pdf.getImageProperties(imgData);
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
        pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight); // Add the image to the PDF
        pdf.save('org-chart.pdf'); // Save the PDF
        window.alert("Pdf Downloaded sucessfully!!!");
      }).catch(() => {
        window.alert("Could not download pdf!!!");
      })
    }
  };

  updateOrgChartLevel = (event: any) => {

    //find child count of top node 
    let orgcharttop = document.querySelector("#orgchart > ul > li");
    let directchildcount = orgcharttop?.childNodes.length && orgcharttop?.childNodes.length > 0 ? orgcharttop?.childNodes.length : 0

    //reset level to defualt.
    for (let i = 1; i <= directchildcount; i++) {
      let parentSelector = "#orgchart > ul > li > ul > li:nth-child(" + i + ") > ul";
      let orgchartlevel = document.querySelector(parentSelector);
      orgchartlevel?.setAttribute("style", "display:flex;");
      let secondlevelchildcount = orgchartlevel?.childNodes.length && orgchartlevel?.childNodes.length > 0 ? orgchartlevel?.childNodes.length : 0
      for (let j = 1; j <= secondlevelchildcount; j++) {
        let secondlevelqueryselector = parentSelector.concat(" > li:nth-child(" + j + ") > ul");
        orgchartlevel = document.querySelector(secondlevelqueryselector);
        orgchartlevel?.setAttribute("style", "display:flex;");
        let thirdlevelchildcount = orgchartlevel?.childNodes.length && orgchartlevel?.childNodes.length > 0 ? orgchartlevel?.childNodes.length : 0
        for (let p = 1; p <= thirdlevelchildcount; p++) {
          let thirdlevelqueryselector = secondlevelqueryselector.concat(" > li:nth-child(" + p + ") > ul");
          orgchartlevel = document.querySelector(thirdlevelqueryselector);
          orgchartlevel?.setAttribute("style", "display:flex;");
        }
      }
    }

    switch (event.target.value) {
      case "1":
        //Level 1
        for (let i = 1; i <= directchildcount; i++) {
          let querySelector = "#orgchart > ul > li > ul > li:nth-child(" + i + ") > ul"
          let orgchartlevel = document.querySelector(querySelector);
          orgchartlevel?.setAttribute("style", "display:none;")
        }
        break;
      case "2":
        //Level 2
        for (let i = 1; i <= directchildcount; i++) {
          let parentSelector = "#orgchart > ul > li > ul > li:nth-child(" + i + ") > ul";
          let orgchartlevel = document.querySelector(parentSelector);
          let secondlevelchildcount = orgchartlevel?.childNodes.length && orgchartlevel?.childNodes.length > 0 ? orgchartlevel?.childNodes.length : 0
          for (let j = 1; j <= secondlevelchildcount; j++) {
            let secondlevelqueryselector = parentSelector.concat(" > li:nth-child(" + j + ") > ul");
            orgchartlevel = document.querySelector(secondlevelqueryselector);
            orgchartlevel?.setAttribute("style", "display:none;");
          }
        }
        break;
      case "3":
        //Level 3
        for (let i = 1; i <= directchildcount; i++) {
          let parentSelector = "#orgchart > ul > li > ul > li:nth-child(" + i + ") > ul";
          let orgchartlevel = document.querySelector(parentSelector);
          let secondlevelchildcount = orgchartlevel?.childNodes.length && orgchartlevel?.childNodes.length > 0 ? orgchartlevel?.childNodes.length : 0
          for (let j = 1; j <= secondlevelchildcount; j++) {
            let secondlevelqueryselector = parentSelector.concat(" > li:nth-child(" + j + ") > ul");
            orgchartlevel = document.querySelector(secondlevelqueryselector);
            let thirdlevelchildcount = orgchartlevel?.childNodes.length && orgchartlevel?.childNodes.length > 0 ? orgchartlevel?.childNodes.length : 0
            for (let p = 1; p <= thirdlevelchildcount; p++) {
              let thirdlevelqueryselector = secondlevelqueryselector.concat(" > li:nth-child(" + p + ") > ul");
              orgchartlevel = document.querySelector(thirdlevelqueryselector);
              orgchartlevel?.setAttribute("style", "display:none;");
            }
          }
        }
        break;
      default:
        for (let i = 1; i <= directchildcount; i++) {
          let parentSelector = "#orgchart > ul > li > ul > li:nth-child(" + i + ") > ul";
          let orgchartlevel = document.querySelector(parentSelector);
          orgchartlevel?.setAttribute("style", "display:flex;");
          let secondlevelchildcount = orgchartlevel?.childNodes.length && orgchartlevel?.childNodes.length > 0 ? orgchartlevel?.childNodes.length : 0
          for (let j = 1; j <= secondlevelchildcount; j++) {
            let secondlevelqueryselector = parentSelector.concat(" > li:nth-child(" + j + ") > ul");
            orgchartlevel = document.querySelector(secondlevelqueryselector);
            orgchartlevel?.setAttribute("style", "display:flex;");
            let thirdlevelchildcount = orgchartlevel?.childNodes.length && orgchartlevel?.childNodes.length > 0 ? orgchartlevel?.childNodes.length : 0
            for (let p = 1; p <= thirdlevelchildcount; p++) {
              let thirdlevelqueryselector = secondlevelqueryselector.concat(" > li:nth-child(" + p + ") > ul");
              orgchartlevel = document.querySelector(thirdlevelqueryselector);
              orgchartlevel?.setAttribute("style", "display:flex;");
            }
          }
        }
        break;
    }
  }

  public render(): React.ReactElement<IOrgchartToolBarProps> {



    return (
      <div className="controls">
        <button onClick={() => this.props.zoomIn()}> <Icon iconName='plus' />Zoom In</button>
        <button onClick={() => this.props.zoomOut()}> <Icon iconName='minus' />Zoom Out</button>
        <button onClick={() => this.props.resetTransform()}><Icon iconName='refresh' />Reset</button>
        <button onClick={this.exportToPDF}><Icon iconName='filePDF' />Export to PDF</button>
        <select id="drplevel" onChange={this.updateOrgChartLevel}>
          <option value="Default">Default</option>
          <option value="1">Level 1</option>
          <option value="2">Level 2</option>
          <option value="3">Level 3</option>
        </select>
      </div>
    )
  }
}