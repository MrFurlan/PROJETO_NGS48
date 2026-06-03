/* UpdatePanel animation scripts */
/* using loadmask to animate panel div */
var animatePanels;
function addPanelToAnimate(id) {
    if (!animatePanels) {
        animatePanels = new Array();
    }
    animatePanels[animatePanels.length] = id;
}

// add handlers
$(document).ready(function () {
    //Sys.WebForms.PageRequestManager.getInstance().add_beginRequest(BeginRequestHandler);
    //Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);
});

// show animation mask
function BeginRequestHandler(sender, args) {
    AddLoadmask(args.get_postBackElement());
}

// remove animation maks
function EndRequestHandler(sender, args) {
    var mask = $(".loadmask");
    if (mask) {
        mask.remove();
    }

    var loadMask = $(".loadmask-msg");
    if (loadMask) {
        loadMask.remove();
    }
}

// check if posted controls contains an
// updatepanel from the list and add mask
function AddLoadmask(elem) {
    if (elem) {
        var childs = elem.parentNode.children;
        if (animatePanels) {
            for (i = 0; i < animatePanels.length; i++) {
                // get panel-id
                //var pnlID = replaceAll(animatePanels[i], "$", "_");
                // posted element-collection contains panel?
                //if (Contains(childs, pnlID)) {
                $(".divLoading").mask("Carregando...");
                //}
            }
        }
    }
}

// recursive replace
function replaceAll(text, strOld, strNew) {
    while (text.indexOf(strOld) > 0) {
        text = text.replace(strOld, strNew);
    }
    return text;
}

// contains function for array
function Contains(arr, elm) {
    if (arr && elm) {
        for (i = 0; i < arr.length; i++) {
            if (arr[i] == elm) {
                return i;
            }
        }
    }
    return -1;
}