javascript:function url() {
 var date = new Date();
 var backup = 1; if (date.getDay() === 0) {backup = 2;} if (date.getDay() === 1) {backup = 3;}
 date.setDate(date.getDate() - backup);
 var y = date.getFullYear(); 
 var m = date.getMonth() + 1; if(m < 10){m = '0' + m;}
 var d = date.getDate(); if(d < 10){d = '0' + d;}
 return 'https://www.google.com/search?q=jeopardy+full+episode+youtube+' + m + '/' + d + '/' + y;
} window.location.href = url();
