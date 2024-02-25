javascript: function url() {
    var date = new Date();
    var y = date.getFullYear();
    var m = date.getMonth() + 1;
    if (m < 10) {
        m = '0' + m;
    }
    var d = date.getDate();
    if (d < 10) {
        d = '0' + d;
    }
    var date = y + m + d;
    return 'https://www.wsj.com/print-edition/' + date + '/frontpage';
}
window.location.href=url();
