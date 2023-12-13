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
const SNN = ["738-26-3677", "145-44-7809", "288-49-1174"];
const OtherBankProducts = ["DreaMaker", "Eagle Community Home Loan"];
const MNPITriggerWords = ["account", "SNN", "Legal Disputes", "M&A", "Hiring Plans"];

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
    document.getElementById("check-bb-button").onclick = () => checkBBDisclaimer();
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

function checkBBDisclaimer(){
  if(selectedOption == "current"){
    checkBBSingle();
  } else if(selectedOption == "all"){
    checkBBAll();
  }
}

function checkMNPI() {
  if(selectedOption == "current"){
    // call the singleMNPIfunction here
    console.log("SelectedOption is: ", selectedOption);
    console.log("curr MNPI Selected");
    printCurrStrings();
  } else if(selectedOption == "all"){
    // call the allMNPI function
    console.log("SelectedOption is: ", selectedOption);
    console.log("all MNPI Selected");
    printAllStrings();
  }
}

function checkSource() {
  if(selectedOption == "current"){
    // call the singleSource function here
  } else if(selectedOption == "all"){
    // call the allSource function
  }
}

function checkAll(){
  if(selectedOption == "current"){
    // call the singleCheckAll function here
  } else if(selectedOption == "all"){
    // call the allCheckAll function
  }
}


//BUTTON ACTION FUNCTIONS HERE

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


// Function to check if the disclaimer is present in any of the slides
async function checkBBAll() {
  await PowerPoint.run(async (context) => {
      const sls = context.presentation.slides;
      sls.load("items");
      await context.sync();
      //console.log("Number of slides: " + sls.items.length);

      for (let j = 0; j < sls.items.length; j++) {
          const sheet = context.presentation.slides.getItemAt(j);
          const shapes = sheet.shapes;
          shapes.load("items");
          await context.sync();

          //console.log("Number of shapes on this slide: ", shapes.items.length);

          for (let i = 0; i < shapes.items.length; i++) {
              const s = shapes.getItemAt(i);
              const t = s.textFrame.textRange;
              t.load();
              try {
                  await context.sync();

                  // Check if the disclaimer is present in the current text
                  if (await hasDisclaimer(t.text)) {
                      console.log("Disclaimer found on slide " + (j + 1));
                      return true;
                  }

              } catch (err) {
                  console.log("Non-text shape skipped");
              }
          }
      }

      // If no disclaimer is found in any slide, log the result
      console.log("Disclaimer not found in any slide");
      return false;
  });
}

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

//function for check MNPI single
function checkMNPISingle(){
  

}

//function for check MNPI all
function checkMNPIAll(){

}

//function for check sources single
function checkSoucesSingle(){

}

//function for check sources all
function checkSourcesAll(){

}

//function for everything check single
function checkEverythingSingle(){

}

//function for everything check all
function checkEverythingAll(){
  
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
          console.log(allSlideText[i][j] + " from slide ", i+1);
      }
  }
}

//functions to reset array and slide variables
function resetGlobalVars(){
  slideno = "n/a";
  currSlideText = [];

  allSlideText = []; 

}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      console.log("Error: " + error.toString());
  }
}