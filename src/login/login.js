// Buton elementini yakalayalım
// eslint-disable-next-line no-undef
const myButton = document.getElementById("myButton");

/* eslint-disable */
function onButtonClick() {
    console.log("Button clicked!");
}

// Butona tıklama olayını dinleyelim ve onButtonClick fonksiyonunu çağıralım
myButton.addEventListener("click", onButtonClick);