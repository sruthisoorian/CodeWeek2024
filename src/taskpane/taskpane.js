/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */


// GLOBAL VARIABLES
var selectedOption = ""; // Variable to select radio button
var slideno = "n/a";  // Variable that hold current slide number -> for single slide text extraction
var currSlideText = [];  //Array that holds strings of current slide --> for single slide text extraction

var allSlideText = []; //2D Array that holds strings of all slides -> for all slides text extraction


//MNPI String Banks
const accountNumbers = ["6724301068", "8374882736", "2749930274"];
const SSN = ["738-26-3677", "145-44-7809", "288-49-1174"];
const OtherBankProducts = ["DreaMaker", "Eagle Community Home Loan"];
const MNPITriggerWords = ["Legal Disputes", "M&A", "Hiring Plans"];

//Output String Array
var displayOutput = [];


Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {

    document.getElementById('view1').style.display = "block";
    // document.getElementById('current-slide').checked =  true;

    document.getElementById('output-paragraph').innerText = '';

    document.getElementById('region-search-label').innerText = '';
    const infoIcons = document.getElementsByClassName("info-icon");

    // Show tooltip for the corresponding checklist item
    for (const icon of infoIcons) {
      icon.addEventListener("mouseover", function () {
        const lines = JSON.parse(this.getAttribute("data-info"));
        const text = lines.join('<br>');
        showTooltip(text, icon);
      });

      icon.addEventListener("mouseout", function () {
        hideTooltip();
      });
    }

    // Assign event handlers to the tabs
    document.getElementById("tabs").addEventListener("click", function (event) {
      // TODO
      // Reset - hide the labels when user switches between tabs
      labels.forEach(function (label) {
        label.style.display = "none";
      });
      if (event.target.classList.contains("tablinks")) {
        var tabName = event.target.getAttribute("onclick").split("'")[1];
        openTab(tabName);

        // Remove the 'active' class from all the tabs
        var tabs = document.getElementsByClassName("tablinks");
        for (var i = 0; i < tabs.length; i++) {
          tabs[i].classList.remove("active");
        }

        // Apply the 'active' class to the clicked tab
        event.target.classList.add("active");
      }
    });

    // Get references to all buttons and labels
    var buttons = document.querySelectorAll(".submit-button");
    var labels = document.querySelectorAll(".label");


    // Add click event listeners to all buttons
    buttons.forEach(function (button, index) {
      button.addEventListener("click", function () {
        // Hide all the labels
        labels.forEach(function (label) {
          label.style.display = "none";
        });

        // Show the message on selected button click
        labels[index].style.display = "block";
      });
    });

    //event bind the buttons
    document.getElementById("curr-slides-button").onclick = () => tryCatch(extractCurrentSlideText);
    document.getElementById("all-slides-button").onclick = () => tryCatch(extractAllSlideText);
    document.getElementById("check-bb-button").onclick = () => checkBBDisclaimer();
    document.getElementById("check-mnpi-button").onclick = () => checkMNPI();
    document.getElementById("check-source-button").onclick = () => checkSource();
    document.getElementById("check-all-button").onclick = () => checkAll();
    document.getElementById("check-bb-button").onclick = () => checkBBDisclaimer();

    document.getElementById("clear-output-button").onclick = () => clearOutputButtonPressed();
  }
});


function openTab(tabName) {
  // Hide all views
  var views = document.getElementsByClassName("tabcontent");
  for (var i = 0; i < views.length; i++) {
    views[i].style.display = "none";
  }

  // Show the selected view
  document.getElementById(tabName).style.display = "block";
}

// Tooltip functionality
function showTooltip(text, element) {
  const tooltip = document.getElementById("tooltip");
  tooltip.innerHTML = text;

  const rect = element.getBoundingClientRect();
  const top = rect.top + window.scrollY - tooltip.offsetHeight - 10;
  const left = rect.left + window.scrollX + (rect.width - tooltip.offsetWidth);
  tooltip.style.top = top + "px";
  tooltip.style.left = left + "px";
  tooltip.style.display = "block";
}

// Hide tooltip
function hideTooltip() {
  const tooltip = document.getElementById("tooltip");
  tooltip.style.display = "none";
}


// function selectRadioButton() {
//   // Check which radio button is selected and set the selectedOption variable to the selected radio button
//   if (document.getElementById('current-slide').checked) {
//     setSelection("current");
//   } else if (document.getElementById('all-slides').checked) {
//     setSelection("all");
//   }
//   console.log(selectedOption + " was selected");
// }

// function setSelection(sel){
//   selectedOption = sel;

// }


function checkBBDisclaimer() {
  resetDisplayOutput();
  if (selectedOption == "current") {
    checkBBSingle();
  } else if (selectedOption == "all") {
    checkBBAll();
  }
  setOutputDisplayText();
}

function checkMNPI() {
  resetDisplayOutput();
  console.log("Running " + selectedOption + " MNPI check");
  if (selectedOption == "current") {
    checkMNPISingle();
  } else if (selectedOption == "all") {
    checkMNPIAll();
  }
  setOutputDisplayText();
}

function checkSource() {
  resetDisplayOutput();
  console.log("Running " + selectedOption + " Sources check");
  if (selectedOption == "current") {
    checkSoucesSingle();
  } else if (selectedOption == "all") {
    checkSourcesAll();
  }
  setOutputDisplayText();
}

function checkAll() {
  resetDisplayOutput();
  if (selectedOption == "current") {
    console.log("Running all QA Presentation Assesments on Slide " + slideno + " ==> ");
    displayOutput.push("Running all QA Presentation Assesments on Slide " + slideno + " ==> ")
    checkBBSingle();
    checkMNPISingle();
    checkSoucesSingle();
  } else if (selectedOption == "all") {
    console.log("Running all QA Presentation Assesments on all Slides ==>");
    displayOutput.push("Running all QA Presentation Assesments on all Slides ==>");
    checkBBAll();
    checkMNPIAll();
    checkSourcesAll();
  }
  setOutputDisplayText();
}


//BUTTON ACTION FUNCTIONS HERE

/*
//Check if there is a BB Disclaimer
async function hasDisclaimer(text) {
  // Define the disclaimer text to search for
  const disclaimerText = "These materials have been prepared by one or more affiliates of Bank of America Corporation";
  // Convert both texts to lowercase for case-insensitive comparison
  const lowerCaseText = text.toLowerCase();
  const lowerCaseDisclaimer = disclaimerText.toLowerCase();
  // Check if the disclaimer text is present in the provided text
  return lowerCaseText.includes(lowerCaseDisclaimer);
}

async function checkBBAll() {
  let disclaimerFound = false;

  await PowerPoint.run(async (context) => {
    const sls = context.presentation.slides;
    sls.load("items");
    await context.sync();

    for (let j = 0; j < sls.items.length; j++) {
      const sheet = context.presentation.slides.getItemAt(j);
      const shapes = sheet.shapes;
      shapes.load("items");
      await context.sync();

      for (let i = 0; i < shapes.items.length; i++) {
        const s = shapes.getItemAt(i);
        const t = s.textFrame.textRange;
        t.load();
        try {
          await context.sync();

          // Check if the disclaimer is present in the current text
          if (await hasDisclaimer(t.text)) {
            console.log("Disclaimer found on slide " + (j + 1));
            disclaimerFound = true;
            return;
          }
        } catch (err) {
          console.log("Non-text shape skipped");
        }
      }
    }

    // If no disclaimer is found in any slide, log the result
    console.log("Disclaimer not found in any slide");
  });

  // Display result in the result box after PowerPoint.run completes
  const resultBox = document.getElementById("resultBox");
  resultBox.innerHTML = disclaimerFound
    ? "<div class='output-line'>Disclaimer found in at least one slide</div>"
    : "<div class='output-line'>Disclaimer not found in any slide</div>";
}
*/

// Function to check if the disclaimer is present on the current slide
// async function checkBBSingle() {
//   await PowerPoint.run(async (context) => {
//     const currentSlide = context.presentation.slides.getActiveSlide();
//     const shapes = currentSlide.shapes;
//     shapes.load("items");
//     await context.sync();

//     for (let i = 0; i < shapes.items.length; i++) {
//       const shape = shapes.items[i];
//       const textRange = shape.textFrame.textRange;
//       textRange.load();

//       try {
//         await context.sync();

//         // Check if the disclaimer is present in the current text
//         if (await hasDisclaimer(textRange.text)) {
//           console.log("Disclaimer found on the current slide");
//           return true;
//         }
//       } catch (err) {
//         console.log("Non-text shape skipped");
//       }
//     }

//     // If no disclaimer is found on the current slide, log the result
//     console.log("Disclaimer not found on the current slide");
//     return false;
//   });
// }

//function for check BB Disclaimer Single
function checkBBSingle() {
  const disclaimerText = "These materials have been prepared by one or more affiliates of Bank of America Corporation";
  var found = false;

  currSlideText.forEach(function (x) {
    if (x.toLowerCase().includes(disclaimerText.toLowerCase())) {
      found = true;
    }
  });

  if (found == true) {
    if (slideno == 2) {
      //displayOutput.push("Slide 2 - BB Disclaimer placement is compliant.");
      console.log("Slide 2 - BB Disclaimer placement is compliant.");
    }
    else {
      displayOutput.push("Slide " + slideno + " - BB Disclaimer is found, but must be placed as 2nd slide, behind cover page");
      console.log("Slide " + slideno + " - BB Disclaimer is found, but must be placed as 2nd slide, behind cover page");

    }
  }
  else {
    if (slideno == 2) {
      console.log("Slide 2 - BB Disclaimer needs to be on 2nd slide, behind cover page");
    }
    else {
      //displayOutput.push("Slide " + slideno + " - Warning: BB Disclaimer is not found.");
      console.log("Slide " + slideno + " - Warning: BB Disclaimer is not found.");

    }

  }



}

//function for check BB Disclaimer All
function checkBBAll() {
  const disclaimerText = "These materials have been prepared by one or more affiliates of Bank of America Corporation";
  var found = false;
  for (var i = 0; i < allSlideText.length; i++) {
    for (var j = 0; j < allSlideText[i].length; j++) {
      if (allSlideText[i][j].toLowerCase().includes(disclaimerText.toLowerCase())) {
        found = true;
        if ((i + 1) == 2) {
          //displayOutput.push("Slide 2 - BB Disclaimer placement is compliant.");
          console.log("Slide 2 - BB Disclaimer placement is compliant.");
        }
        else {
          displayOutput.push("Slide " + (i + 1) + " - BB Disclaimer is found, but must be placed as 2nd slide, behind cover page");
          console.log("Slide " + (i + 1) + " - BB Disclaimer is found, but must be placed as 2nd slide, behind cover page");
        }

      }
    }
  }

  if(found == false){
    displayOutput.push("Warning: BB Disclaimer is not found.");
    console.log("Warning: BB Disclaimer is not found.");
  }
}


//function for check MNPI single
function checkMNPISingle() {

  currSlideText.forEach(function (x) {
    for (let i = 0; i < accountNumbers.length; i++) {
      if (x.toLowerCase().includes(accountNumbers[i].toString())) {
        displayOutput.push("Slide " + slideno + " - Found account number " + accountNumbers[i].toString() + ". Please remove from slide immediately.");
        console.log("Slide " + slideno + " - Found account number " + accountNumbers[i].toString() + ". Please remove from slide immediately.");
      }
    }
  });

  currSlideText.forEach(function (x) {
    for (let i = 0; i < SSN.length; i++) {
      if (x.toLowerCase().includes(SSN[i])) {
        displayOutput.push("Slide " + slideno + " - Found SSN number " + SSN[i].toString() + ". Please remove from slide immediately.");
        console.log("Slide " + slideno + " - Found SSN number " + SSN[i].toString() + ". Please remove from slide immediately.");
      }
    }
  });

  currSlideText.forEach(function (x) {
    for (let i = 0; i < OtherBankProducts.length; i++) {
      if (x.includes(OtherBankProducts[i])) {
        displayOutput.push("Slide " + slideno + " - Found mention of competitor bank product: " + OtherBankProducts[i] + ". Please verify content of slide.");
        console.log("Slide " + slideno + " - Found mention of competitor bank product: " + OtherBankProducts[i] + ". Please verify content of slide.");
      }
    }
  });

  currSlideText.forEach(function (x) {
    for (let i = 0; i < MNPITriggerWords.length; i++) {
      if (x.toLowerCase().includes(MNPITriggerWords[i].toLowerCase())) {
        displayOutput.push("Slide " + slideno + " - Found indication of MNPI regarding " + MNPITriggerWords[i] + ". Please verify content of slide.");
        console.log("Slide " + slideno + " - Found indication of MNPI regarding " + MNPITriggerWords[i] + ". Please verify content of slide.");
      }
    }
  });


}

//function for check MNPI all
function checkMNPIAll() {
  for (var i = 0; i < allSlideText.length; i++) {
    for (var j = 0; j < allSlideText[i].length; j++) {
      //check for account numbers
      for (let x = 0; x < accountNumbers.length; x++) {
        if (allSlideText[i][j].toLowerCase().includes(accountNumbers[x].toString())) {
          displayOutput.push("Slide " + (i + 1) + " - Found account number " + accountNumbers[x].toString() + ". Please remove from slide immediately.");
          console.log("Slide " + (i + 1) + " - Found account number " + accountNumbers[x].toString() + ". Please remove from slide immediately.");
        }
      }
      //check for ssns
      for (let x = 0; x < SSN.length; x++) {
        if (allSlideText[i][j].toLowerCase().includes(SSN[x])) {
          displayOutput.push("Slide " + (i + 1) + " - Found SSN number " + SSN[x].toString() + ". Please remove from slide immediately.");
          console.log("Slide " + (i + 1) + " - Found SSN number " + SSN[x].toString() + ". Please remove from slide immediately.");
        }
      }
      //check for other bank products
      for (let x = 0; x < OtherBankProducts.length; x++) {
        if (allSlideText[i][j].includes(OtherBankProducts[x])) {
          displayOutput.push("Slide " + (i + 1) + " - Found mention of competitor bank product: " + OtherBankProducts[x] + ". Please verify content of slide.");
          console.log("Slide " + (i + 1) + " - Found mention of competitor bank product: " + OtherBankProducts[x] + ". Please verify content of slide.");
        }
      }
      //check for MNPI trigger words
      for (let x = 0; x < MNPITriggerWords.length; x++) {
        if (allSlideText[i][j].toLowerCase().includes(MNPITriggerWords[x].toLowerCase())) {
          displayOutput.push("Slide " + (i + 1) + " - Found indication of MNPI regarding " + MNPITriggerWords[x] + ". Please verify content of slide.");
          console.log("Slide " + (i + 1) + " - Found indication of MNPI regarding " + MNPITriggerWords[x] + ". Please verify content of slide.");
        }
      }
    }
  }

}

//function for check sources single
function checkSoucesSingle() {

  //If Source: Bofa
  currSlideText.forEach(function (x) {
    if (x.toLowerCase().includes("source: bofa") || x.toLowerCase().includes("source : bofa") || x.toLowerCase().includes("source- bofa") || x.toLowerCase().includes("source - bofa")) {
      if (x.toLowerCase().slice(-4) == "bofa") {
        displayOutput.push("Slide " + slideno + " - Insufficient citation found. Please replace with more specific citation.");
        console.log("Slide " + slideno + " - Insufficient citation found. Please replace with more specific citation.");
      }
    }
  });

  //If Source: Bofa Global Research
  currSlideText.forEach(function (x) {
    if (x.toLowerCase().includes("source: bofa global research") || x.toLowerCase().includes("source : bofa global research") || x.toLowerCase().includes("source- bofa global research") || x.toLowerCase().includes("source - bofa global research")) {
      displayOutput.push("Slide " + slideno + " - Insufficient citation found. Please replace with more specific citation.");
      console.log("Slide " + slideno + " - Insufficient citation found. Please replace with more specific citation.");
    }
  });

}

//function for check sources all
function checkSourcesAll() {
  //If Souce : Bofa
  for (var i = 0; i < allSlideText.length; i++) {
    for (var j = 0; j < allSlideText[i].length; j++) {
      if (allSlideText[i][j].toLowerCase().includes("source: bofa") || allSlideText[i][j].toLowerCase().includes("source : bofa") || allSlideText[i][j].toLowerCase().includes("source- bofa") || allSlideText[i][j].toLowerCase().includes("source - bofa")) {
        if (allSlideText[i][j].toLowerCase().slice(-4) == "bofa") {
          displayOutput.push("Slide " + (i + 1) + " - Insufficient citation found. Please replace with more specific citation.");
          console.log("Slide " + (i + 1) + " - Insufficient citation found. Please replace with more specific citation.");
        }
      }
    }
  }

  //If Source: Bofa Global Research
  for (var i = 0; i < allSlideText.length; i++) {
    for (var j = 0; j < allSlideText[i].length; j++) {
      if (allSlideText[i][j].toLowerCase().includes("source: bofa global research") || allSlideText[i][j].toLowerCase().includes("source : bofa global research") || allSlideText[i][j].toLowerCase().includes("source- bofa global research") || allSlideText[i][j].toLowerCase().includes("source - bofa global research")) {
        displayOutput.push("Slide " + (i + 1) + " - Insufficient citation found. Please replace with more specific citation.");
        console.log("Slide " + (i + 1) + " - Insufficient citation found. Please replace with more specific citation.");
      }
    }
  }



}

//function for everything check single
function checkEverythingSingle() {

}

//function for everything check all
function checkEverythingAll() {

}


//EXTRACT strings of CURRENT SLIDE
function extractCurrentSlideText() {
  selectedOption = "current";
  document.getElementById('region-search-label').innerText = 'Current Slide Only';
  resetGlobalVars();
  Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => {
    var s = "";
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log("Error getting Metadata: " + asyncResult.error.message);
      s = "Error getting Metadata: " + asyncResult.error.message;
    } else {
      console.log("Metadata for selected slides: " + JSON.stringify(asyncResult.value));
      s = JSON.stringify(asyncResult.value);
      let pos = s.indexOf("index\":") + 7;
      slideno = s.substring(pos, s.indexOf("}]}"));
    }

    //console.log("result: ", s);
    console.log("slideno: ", slideno);
    console.log("slideno as int: ", parseInt(slideno));
    getCurrentSlideStrings(parseInt(slideno));
  });
}

//Helper funtion of extract strings of current slide
async function getCurrentSlideStrings(n) {
  await PowerPoint.run(async (context) => {
    console.log("getting text from this slide: ", n);
    const sheet = context.presentation.slides.getItemAt(n - 1);
    const shapes = sheet.shapes;
    shapes.load("items");
    await context.sync();

    console.log("Number of shapes on this slide: ", shapes.items.length);

    for (let i = 0; i < shapes.items.length; i++) {
      const s = shapes.getItemAt(i);
      const t = s.textFrame.textRange;
      t.load();
      try {
        await context.sync();
        console.log(t.text);
        currSlideText.push(t.text);
      }
      catch (err) {
        console.log("Non-text shape skipped");
      }


    }

  });

}


//EXTRACT strings of ALL SLIDES
async function extractAllSlideText() {
  selectedOption = "all";
  document.getElementById('region-search-label').innerText = 'All Slides';
  resetGlobalVars();
  await PowerPoint.run(async (context) => {
    const sls = context.presentation.slides;
    sls.load("items");
    await context.sync();
    console.log("Number of slides: " + sls.items.length);

    for (let j = 0; j < sls.items.length; j++) {
      const sheet = context.presentation.slides.getItemAt(j);
      const shapes = sheet.shapes;
      shapes.load("items");
      await context.sync();

      console.log("Number of shapes on this slide: ", shapes.items.length);
      const slideStringsTemp = [];

      for (let i = 0; i < shapes.items.length; i++) {
        const s = shapes.getItemAt(i);
        const t = s.textFrame.textRange;
        t.load();
        try {
          await context.sync();
          console.log(t.text);
          slideStringsTemp.push(t.text);
        }
        catch (err) {
          console.log("Non-text shape skipped");
        }

      }

      allSlideText.push(slideStringsTemp);

    }

  });
}

//functions to print slide string arrays to console
function printCurrStrings() {
  currSlideText.forEach(function (x) {
    console.log(x);
  })
}

function printAllStrings() {
  for (var i = 0; i < allSlideText.length; i++) {
    for (var j = 0; j < allSlideText[i].length; j++) {
      console.log(allSlideText[i][j] + " from slide ", i + 1);
    }
  }
}

//output display test
function setOutputDisplayText(){
  var outputString = '';
  for(let i = 0; i < displayOutput.length; i++){
    outputString = outputString + "> " + displayOutput[i] + "\n";
  }
  document.getElementById('output-paragraph').innerText = outputString;
}

//functions to reset array and slide variables
function resetGlobalVars() {
  slideno = "n/a";
  currSlideText = [];

  allSlideText = [];
  displayOutput = [];

}

function resetDisplayOutput(){
  displayOutput = [];
  document.getElementById('output-paragraph').innerText = '';
}

function clearOutputButtonPressed(){
  resetDisplayOutput();
  selectedOption = "";
  document.getElementById('region-search-label').innerText = '';
  resetGlobalVars();

}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.log("Error: " + error.toString());
  }
}