(function(){'use strict';
const target=new URL('./app/',window.location.href),params=new URLSearchParams(window.location.search),sheet=params.get('sheet')||'',choir=(params.get('choir')||'').trim();
if(/^[A-Za-z0-9_-]{20,}$/.test(sheet)&&choir&&choir.length<=80){target.searchParams.set('sheet',sheet);target.searchParams.set('choir',choir);}
document.getElementById('open-web-app').href=target.href;
window.location.replace(target.href);
})();
