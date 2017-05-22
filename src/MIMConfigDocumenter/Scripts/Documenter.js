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

    var downloadLink = document.getElementById("DownloadLink");
    if (x.checked == true) {
        downloadLink.style.display = "";
        DownloadScript(downloadLink);
    }
    else {
        downloadLink.style.display = "none";
    }
}
