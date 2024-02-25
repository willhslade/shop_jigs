javascript: function muteMe(elem) {
    elem.muted = true;
    elem.pause(); 
} 
(() => {
    var elems = document.querySelectorAll("video,audio");
    [].forEach.call(elems,function(elem){muteMe(elem);});
})()
