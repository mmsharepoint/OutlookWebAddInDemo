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
    renderAttachments(bootstrapToken);
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

  async function renderAttachments(bootstrapToken) {
    const mailID = Office.context.mailbox.item.itemId;
    const restMailID = Office.context.mailbox.convertToRestId(mailID, Office.MailboxEnums.RestVersion.v2_0);
    const requestBody = { MessageID: restMailID };
    $.ajax({
      url: '/api/Web/GetAttachments',
      type: 'POST',
      accepts: 'application/json',
      headers: {
        "Authorization": "Bearer " + bootstrapToken // Used here to pass authorization in WebController
      },
      data: JSON.stringify(requestBody),
      contentType: "application/json; charset=utf-8"
    }).done(function (data) {
      console.log(data);
      var list = document.getElementById('attachmentsList');
      data.forEach((doc) => {
        var listItem = document.createElement('li');
        listItem.innerHTML = '<input type="checkbox" data-docID="' + doc.id + '" data-docName="' + doc.name + '" /> ' + doc.name;
        list.appendChild(listItem);
      });
      var savBtn = document.getElementById('saveAttachments');
      savBtn.addEventListener('click', saveAttachments);
    }).fail(function (error) {
      console.log(error);
    });
  }
  async function saveAttachments() {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken();
    var attachments = document.querySelectorAll('.mmAttachmentList input[type="checkbox"]:checked');
    var attArr = Array.from(attachments);
    var selectedDocs = [];
    attArr.forEach((sel) => {
      selectedDocs.push({ id: sel.getAttribute('data-docID'), filename: sel.getAttribute('data-docName') });
    });
    const mailID = Office.context.mailbox.item.itemId;
    const restMailID = Office.context.mailbox.convertToRestId(mailID, Office.MailboxEnums.RestVersion.v2_0);
    const requestBody = { Attachments: selectedDocs, MessageID: restMailID };
    $.ajax({
      url: '/api/Web/SaveAttachments',
      type: 'POST',
      accepts: 'application/json',
      headers: {
        "Authorization": "Bearer " + bootstrapToken // Used here to pass authorization in WebController
      },
      data: JSON.stringify(requestBody),
      contentType: "application/json; charset=utf-8"
    }).done(function (data) {
      console.log(data);
    }).fail(function (error) {
      console.log(error);
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