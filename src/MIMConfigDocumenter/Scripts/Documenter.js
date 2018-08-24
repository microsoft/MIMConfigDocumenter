window.onload = function (e) {
    document.getElementById("OnlyShowChanges").disabled = false;
}

function ToggleVisibility() {
    var x = document.getElementById("OnlyShowChanges");
    var elements = document.getElementsByClassName("CanHide");
    for (var i = 0; i < elements.length; ++i) {
        if (x.checked == true) {
            elements[i].style.display = "none";
        }
        else {
            elements[i].style.display = "";
        }
    }
}
