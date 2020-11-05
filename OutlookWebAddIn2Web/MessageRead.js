(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      var element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();
    });
  };

  $(document).ready(function () {
    $.ajax({
      url: '/api/Web',
      accepts: 'application/json'
    })
      .done((response) => {
        var selDiv = document.getElementById('selCustomer');
        var sel = document.createElement('select');
        response.forEach((val) => {
          var opt = document.createElement("option");
          opt.value = val.ID;
          opt.text = val.Name;
          sel.options.add(opt);        
        });
        selDiv.appendChild(sel);
      });
  });
  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();