/* eslint-disable no-inner-declarations */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(function (info) {
    // eslint-disable-next-line no-empty
    if (info.host === Office.HostType.Outlook) {
      const exitButton = document.getElementById("exitBtn");
      const firmaButton = document.getElementById("firmaBtn");
      // Butona tıklama işlemi event ekleme
      exitButton.addEventListener("click", onButtonClick);
      firmaButton.addEventListener("click", onButtonClick2);
      
      function onButtonClick() {   
        localStorage.removeItem("sessionId");
        window.location.href = "login.html";
      }
      function onButtonClick2() {   
        window.open("http://democrm.logo.com.tr/LOGOCRM/default.aspx#ViewID=userStartScreen", '_blank');  
      }
      // eslint-disable-next-line @typescript-eslint/no-unused-vars
      function myFirma() {
        window.open("http://democrm.logo.com.tr/LOGOCRM/default.aspx#ViewID=userStartScreen", '_blank');  
        // Burada yapmak istediğiniz işlemi gerçekleştirebilirsiniz
      }
      if (isPersistenceSupported()) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, setUserNameInfo);
      }
            
      function isPersistenceSupported() {        
        return Office.context.mailbox.addHandlerAsync !== undefined;
      }
  
      // eslint-disable-next-line @typescript-eslint/no-unused-vars
      function setUserNameInfo(eventArgs) {
        const item = Office.context.mailbox.item;
        debugger;
        const item2 = Office.context.mailbox.userProfile;
        if(item){
          var emailAddress = item.from.emailAddress;
   
          var linkElement = document.getElementById("userMail");
          debugger;
          // Öğenin içeriğini değiştir
          linkElement.innerHTML = '<img src="https://www.logo.com.tr/_next/image?url=%2Flogo.webp&w=1920&q=75" style="border-radius: 10%;" alt="" class="img-circle" width="44" />' + emailAddress;        
  
        }        
      }  

    }
   

   
  });
  
   
 