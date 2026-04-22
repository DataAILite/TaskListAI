function showSpinner() {

    setTimeout(function () { document.getElementById("spinner").style.display = ""; document.images["imgSpinner"].src = "Controls/Images/WaitImage2.gif"; }, 200);
}
    
function ClearTextbox(ctlId) {
    //var txt = document.getElementById(ctlId);
    //if (txt != null)
    //  txt.value = '';
    SetText(ctlId, '');
}
function ClearNewPswd()
{
    ClearTextbox("txtNew");
    ClearTextbox("txtRepeat");
}
function ClearCurrentPswd()
{
    ClearTextbox("txtCurrent")
}
function SetText(ctlId,text)
{
    var txt = document.getElementById(ctlId);
    if (txt != null)
        txt.value = text;
}

function clickFileUpload() {
    var fileUpload = document.getElementById("FileRDL");
    if (fileUpload != null) fileUpload.click();
}
function getAttachedFile() {
    var fileUpload = document.getElementById("FileRDL");
    if (fileUpload != null) {
        var lblAttached = document.getElementById("lblFileChosen");
        lblAttached.innerText = fileUpload.files[0].name;
        
    }
}