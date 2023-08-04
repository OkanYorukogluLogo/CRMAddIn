/* eslint-disable no-inner-declarations */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(function (info) {
  // Office.js tam olarak yüklendiğinde buradaki kod çalışacak.
  if (info.host === Office.HostType.Outlook) {
    // Butonu seçin ve tıklama işlemini Office.onReady() içinde tanımlayın
    const loginButton = document.getElementById("loginBtn");

      // Ayarlar butonuna tıklama olayı ekleme
      const settingsBtn = document.getElementById("settingsBtn");
      settingsBtn.addEventListener("click", function() {
        const modal = document.getElementById("settingsModal");
        modal.style.display = "block";
      });
      // Modal kapatma butonuna tıklama olayı ekleme
      const closeModalBtn = document.querySelector(".close");
      closeModalBtn.addEventListener("click", function() {
        const modal = document.getElementById("settingsModal");
        modal.style.display = "none";
      });
      // Kaydet butonuna tıklama olayı ekleme
      const saveSettingsBtn = document.getElementById("saveSettingsBtn");
      saveSettingsBtn.addEventListener("click", function() {
        const serviceUrlInput = document.getElementById("serviceUrlInput");
        const serviceUrl = serviceUrlInput.value;

        // Ayarları localStorage'e kaydetme
        localStorage.setItem("serviceUrl", serviceUrl);
        

        // Modalı kapatma
        const modal = document.getElementById("settingsModal");
        modal.style.display = "none";
      });
      
    function onButtonClick() {
      console.log("Button clicked CRM!");

      const username = document.getElementById('username').value;

      getUserGists(username, function(gists, error) {
          const resultElement = document.getElementById('result');
          
          if (error) {
              resultElement.innerHTML = 'Error occurred: ' + error.statusText;
          } else {
              if (gists.length > 0) {
                  let gistList = '<ul>';
                  gists.forEach(function(gist) {
                      gistList += '<li>' + gist.html_url + '</li>';
                  });
                  gistList += '</ul>';
                  resultElement.innerHTML = gistList;
              } else {
                  resultElement.innerHTML = 'No gists found for the user.';
              }
          }
      });
      
    }

    // Butona tıklama işlemi event ekleme
    loginButton.addEventListener("click", onButtonClick);


    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    function getUserGists(user, callback) {
      // eslint-disable-next-line no-debugger
      debugger;
      const requestUrl ='https://localhost:3001/login?authorization=QXBwbGU6QXBwbGU=';  //'http://democrm.logo.com.tr/LogoCRMRest/api/v1.0/login?authorization=QXBwbGU6QXBwbGU=';
      //Office.context.ui.displayDialogAsync('taskpane.html', { height: 50, width: 50 });

      //window.open("taskpane.html", '_blank');




      $.ajax({
          url: requestUrl,
          dataType: 'json',
          type: 'POST',
          rejectUnauthorized: false,
          requestCert: false,
          agent: false
      }).done(function(response) {
          // Ajax isteği başarılı oldu, burada yönlendirmeyi yapabilirsiniz.
          if (response) {
            $("#username").val(response.SessionId);   
          } else {
            // İstek başarılı olsa bile, sunucudan gelen cevaba göre başka bir işlem yapabilirsiniz.
            // Örneğin, hata mesajını kullanıcıya gösterebilirsiniz.
          }
        }).fail(function(error) {
          // İstek başarısız oldu, hata mesajını kullanıcıya gösterebilirsiniz.
          console.log('İstek başarısız:', error);           
        });

  }
  }
});



