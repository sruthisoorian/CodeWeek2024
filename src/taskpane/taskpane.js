/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

var selectedOption;

Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) { 

    document.getElementById('view1').style.display = "block";
    const infoIcons = document.getElementsByClassName("info-icon");

    // Show tooltip for the corresponding checklist item
    for (const icon of infoIcons) {
      icon.addEventListener("mouseover", function() {
          const lines = JSON.parse(this.getAttribute("data-info"));
          const text = lines.join('<br>');
          showTooltip(text, icon);
      });

      icon.addEventListener("mouseout", function() {
          hideTooltip();
      });
    }

    // Assign event handlers to the tabs
    document.getElementById("tabs").addEventListener("click", function (event) {
      // TODO
      // Reset - hide the labels when user switches between tabs
      labels.forEach(function(label) {
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
    buttons.forEach(function(button, index) {
      button.addEventListener("click", function() {
        // Hide all the labels
        labels.forEach(function(label) {
          label.style.display = "none";
        });

        // Show the message on selected button click
        labels[index].style.display = "block";
      });
  });
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

function selectRadioButton() {
  // Check which radio button is selected and set the selectedOption variable to the selected radio button
  if (document.getElementById('current-slide').checked) {
    selectedOption = 'current';
  } else if (document.getElementById('all-slides').checked) {
    selectedOption = 'all';
  }
}

function checkBBDisclaimer(){
  if(selectedOption === 'current'){
    // call the singleBB function here
  } else if(selectedOption === 'all'){
    // call the allBB function
  }
}

function checkMNPI() {
  if(selectedOption === 'current'){
    // call the singleMNPIfunction here
  } else if(selectedOption === 'all'){
    // call the allMNPI function
  }
}

function checkSource() {
  if(selectedOption === 'current'){
    // call the singleSource function here
  } else if(selectedOption === 'all'){
    // call the allSource function
  }
}

function checkAll(){
  if(selectedOption === 'current'){
    // call the singleCheckAll function here
  } else if(selectedOption === 'all'){
    // call the allCheckAll function
  }
}