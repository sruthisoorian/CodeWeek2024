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
  if (selectedOption == "current") {
    // call the singleBB function here
  } else if (selectedOption == "all") {
    // call the allBB function
  }
}

function checkMNPI() {
  console.log("Running " + selectedOption + " MNPI check");
  if (selectedOption == "current") {
    // call the singleMNPIfunction here
    checkMNPISingle();
  } else if (selectedOption == "all") {
    // call the allMNPI function
    checkMNPIAll();
  }
}

function checkSource() {
  if (selectedOption == "current") {
    // call the singleSource function here
  } else if (selectedOption == "all") {
    // call the allSource function
  }
}

function checkAll() {
  if (selectedOption == "current") {
    // call the singleCheckAll function here
  } else if (selectedOption == "all") {
    // call the allCheckAll function
  }
}


//BUTTON ACTION FUNCTIONS HERE

//Function for single BB
function checkBBSingle() {

}

//function for all BB
function checkBBAll() {

}

//function for check MNPI single
function checkMNPISingle() {

  currSlideText.forEach(function (x) {
    for(let i = 0; i < accountNumbers.length; i++){
      if (x.toLowerCase().includes(accountNumbers[i].toString())){
        displayOutput.push("Slide " + slideno + " - Found account number " + accountNumbers[i].toString() + ". Please remove from slide immediately.");
        console.log("Slide " + slideno + " - Found account number " + accountNumbers[i].toString() + ". Please remove from slide immediately.");
      }
    }
  });

  currSlideText.forEach(function (x) {
    for(let i = 0; i < SSN.length; i++){
      if (x.toLowerCase().includes(SSN[i])){
        displayOutput.push("Slide " + slideno + " - Found SSN number " + SSN[i].toString() + ". Please remove from slide immediately.");
        console.log("Slide " + slideno + " - Found SSN number " + SSN[i].toString()  + ". Please remove from slide immediately.");
      }
    }
  });

  currSlideText.forEach(function (x) {
    for(let i = 0; i < OtherBankProducts.length; i++){
      if (x.includes(OtherBankProducts[i])){
        displayOutput.push("Slide " + slideno + " - Found mention of competitor bank product: " + OtherBankProducts[i] + ". Please verify content of slide.");
        console.log("Slide " + slideno + " - Found mention of competitor bank product: " + OtherBankProducts[i] + ". Please verify content of slide.");
      }
    }
  });

  currSlideText.forEach(function (x) {
    for(let i = 0; i < MNPITriggerWords.length; i++){
      if (x.toLowerCase().includes(MNPITriggerWords[i].toLowerCase())){
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
      for(let x = 0; x < accountNumbers.length; x++){
        if (allSlideText[i][j].toLowerCase().includes(accountNumbers[x].toString())){
          displayOutput.push("Slide " + (i+1) + " - Found account number " + accountNumbers[x].toString() + ". Please remove from slide immediately.");
          console.log("Slide " + (i+1) + " - Found account number " + accountNumbers[x].toString() + ". Please remove from slide immediately.");
        }
      }
      //check for ssns
      for(let x = 0; x < SSN.length; x++){
        if (allSlideText[i][j].toLowerCase().includes(SSN[x])){
          displayOutput.push("Slide " + (i+1) + " - Found SSN number " + SSN[x].toString() + ". Please remove from slide immediately.");
          console.log("Slide " + (i+1) + " - Found SSN number " + SSN[x].toString()  + ". Please remove from slide immediately.");
        }
      }
      //check for other bank products
      for(let x = 0; x < OtherBankProducts.length; x++){
        if (allSlideText[i][j].includes(OtherBankProducts[x])){
          displayOutput.push("Slide " + (i+1) + " - Found mention of competitor bank product: " + OtherBankProducts[x] + ". Please verify content of slide.");
          console.log("Slide " + (i+1) + " - Found mention of competitor bank product: " + OtherBankProducts[x] + ". Please verify content of slide.");
        }
      }
      //check for MNPI trigger words
      for(let x = 0; x < MNPITriggerWords.length; x++){
        if (allSlideText[i][j].toLowerCase().includes(MNPITriggerWords[x].toLowerCase())){
          displayOutput.push("Slide " + (i+1) + " - Found indication of MNPI regarding " + MNPITriggerWords[x] + ". Please verify content of slide.");
          console.log("Slide " + (i+1) + " - Found indication of MNPI regarding " + MNPITriggerWords[x] + ". Please verify content of slide.");
        }
      }
    }
  }

}

//function for check sources single
function checkSoucesSingle() {

}

//function for check sources all
function checkSourcesAll() {

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

//functions to reset array and slide variables
function resetGlobalVars() {
  slideno = "n/a";
  currSlideText = [];

  allSlideText = [];
  displayOutput = [];

}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.log("Error: " + error.toString());
  }
}