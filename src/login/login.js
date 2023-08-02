/* eslint-disable no-inner-declarations */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(function (info) {
    // Office.js tam olarak yüklendiğinde buradaki kod çalışacak.
    if (info.host === Office.HostType.Outlook) {
      // Butonu seçin ve tıklama işlemini Office.onReady() içinde tanımlayın
      const myButton = document.getElementById("myButton");

      function onButtonClick() {
        console.log("Button clicked CRM!");
      }

      // Butona tıklama işlemi event ekleme
      myButton.addEventListener("click", onButtonClick);
    }
  });