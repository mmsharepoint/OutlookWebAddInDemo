(function () {
  "use strict";

  var messageBanner;

  Office.onReady(function () {
    // Office is ready
    $(document).ready(function () {
      var element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();

      var btn = document.getElementById('graphBtn');
      btn.addEventListener('click', accessMicrosoftGraph);
      getCustomers();
    });
  });

  async function getCustomers() {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken();
    $.ajax({
      url: '/api/Web',
      accepts: 'application/json',
      headers: {
        "Authorization": "Bearer " + bootstrapToken // Used here to pass authorization in WebController
      },
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
  }

  async function accessMicrosoftGraph() {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken();
    const mailID = Office.context.mailbox.item.itemId;
    const requestBody = { MessageID: mailID };
    $.ajax({
      type: "POST",
      url: '/api/Web/GetMimeMessage',
      headers: {
        "Authorization": "Bearer " + bootstrapToken
      },
      data: JSON.stringify(requestBody),
      contentType: "application/json; charset=utf-8"
    }).done(function (data) {
      console.log(data);
      // renderItems(data);
    }).fail(function (error) {
      console.log(error);
    }).always(function () {
      // Cleanup
    });
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();