/* eslint-disable no-inner-declarations */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
Office.onReady(function (info) {
    // Office.js tam olarak yüklendiğinde buradaki kod çalışacak.
    if (info.host === Office.HostType.Outlook) {
      // Butonu seçin ve tıklama işlemini Office.onReady() içinde tanımlayın

      const sessionID = localStorage.getItem("sessionId");
      if(sessionID)
      {        
        window.location.href = "home.html";
      }

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


          // Info Modal kapatma butonuna tıklama olayı ekleme
          const closeInfoModalBtn = document.querySelector(".infoClose");
          closeInfoModalBtn.addEventListener("click", function() {
            const infomodal = document.getElementById("infoModal");
            infomodal.style.display = "none";
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

        // Ayarları silme işlevi
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        function clearSettings() {
          // localStorage'de "serviceUrl" adında saklanan değeri sil
          localStorage.removeItem("serviceUrl");
          
          // Eğer başka ayarları da saklıyorsanız, onları da aynı şekilde silebilirsiniz
        }
/*
      if (isPersistenceSupported()) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, setUserNameInfo);
      }
            
      function isPersistenceSupported() {        
        return Office.context.mailbox.addHandlerAsync !== undefined;
      }

      // eslint-disable-next-line @typescript-eslint/no-unused-vars
      function setUserNameInfo(eventArgs) {
        const item = Office.context.mailbox.item;
        if(item){
          var emailAddress = item.from.emailAddress;
          document.getElementById('username').value = emailAddress;
        }        
      }      
*/
      function onButtonClick() {       
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

        //const username = document.getElementById('username').value;
        //const password = document.getElementById('password').value;
        //const base64Credentials = (username +":" + password).toString('base64');//Buffer.from(`${username}:${password}`).toString('base64');


        const username = document.getElementById('username').value;
        const password = document.getElementById('password').value;
        const credentials = `${username}:${password}`;
        const base64Credentials = btoa(credentials);


        // Hedef URL'yi her istek için belirleyin
        //const targetURL = 'http://democrm.logo.com.tr/LogoCRMRest/api/v1.0/login?authorization=' + base64Credentials;



        const requestUrl ='https://localhost:3001/login?authorization=' + base64Credentials; //QXBwbGU6QXBwbGU=';  //'http://democrm.logo.com.tr/LogoCRMRest/api/v1.0/login?authorization=QXBwbGU6QXBwbGU=';
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
            if (response && response.Message =="") {
              //$("#username").val(response.SessionId);   
              localStorage.setItem("sessionId", response.SessionId);
              window.location.href = "home.html";
            } else {
              const infomodal = document.getElementById("infoModal");
              infomodal.style.display = "block"; 
              document.getElementById('infoValue').innerHTML = response.Message;//"Giriş başarısız oldu; sunucuda bir hata oluştu.";  
            }
          }).fail(function(error) {
            // İstek başarısız oldu, hata mesajını kullanıcıya gösterebilirsiniz.
            const infomodal = document.getElementById("infoModal");
            infomodal.style.display = "block";  
            document.getElementById('infoValue').innerHTML = "Giriş başarısız oldu; sunucuda bir hata oluştu. " + error.statusText;          
          });
 
    }

    }
  });



