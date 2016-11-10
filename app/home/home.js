(function(){
  'use strict';
  var searchResults;
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      jQuery('#get-data-from-selection').click(getDataFromSelection);
      jQuery('#display-range-objects').click(displayObjects);
    });
  };

  // Reads data from current document selection and displays a notification
  function getDataFromSelection(){
    Word.run(function(context) {
      searchResults = context.document.body.search('abc@gmail.com');
      context.load(searchResults);
    return context.sync().then(function(){
      for (var i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].insertContentControl();
      }
      app.showNotification("Found Count:" + searchResults.items.length);
      return context.sync();
    });
    }).catch(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          app.showNotification('The selected text is:', '"' + result.value + '"');
        } else {
          app.showNotification('Error:', result.error.message);
        }
      }
    );
  }
  function displayObjects() {
    Word.run(function(context) {
      var contentControls = context.document.contentControls;
      context.load(contentControls);
      return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            contentControls.items[0].load('text');
            // Queue a command to clear the contents of the first content control.
            app.showNotification("found text: " + contentControls.items[0].text);
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });
        } 
      }).catch(function (result) {
        app.showNotification("Error: " + result);
      });
 
    });
  }

})();
