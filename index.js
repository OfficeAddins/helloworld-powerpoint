window.onload = function Start() { 
   console.log('hello world'); 

   var app_1 = document.getElementById("app"); 
   app_1.innerHTML = '<b>hello world</b>'; 

   app_1.innerHTML = app_1.innerHTML + 
      '<br><input type="button" value="Add Data" onclick="loadPowerPointData();" />'; 
} 

Office.initialize = function (reason) { 
} 

window.loadPowerPointData = loadPowerPointData; 
function loadPowerPointData() { 
   console.log('powerpoint data loaded'); 

   PowerPoint.run(function (ctx) { 
      Office.context.document.setSelectedDataAsync("BetterSolutions.com", { coercionType: Office.CoercionType.Text }); 

       return ctx.sync(); 
    }); 
} 