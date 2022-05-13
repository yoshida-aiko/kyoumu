<script language="javascript1.2"> <!--

window.onload=init

ie=false
nn=false
if(document.all){ie=true}
if(navigator.appName=="Netscape"){nn=true}

function init(){
if(ie){
frames[0].document.body.onscroll=scrollie
frames[2].document.body.onscroll=scrollie
}
if(nn){
scroll=new Array(0,0)
scrollnn()
}
}

function scrollie(){
if(frames[0].event){
frames[2].scrollTo(frames[0].document.body.scrollLeft,frames[2].document.body.scrollTop)
}
if(frames[2].event){
frames[0].scrollTo(frames[2].document.body.scrollLeft,frames[0].document.body.scrollTop)
}
}

function scrollnn(){
var scr0=frames[0].pageXOffset
var scr1=frames[2].pageXOffset
var y0=frames[0].pageYOffset
var y1=frames[2].pageYOffset
if(scr0!=scroll[0]){
//左がスクロール
frames[2].scrollTo(scr0,y1)
scroll[0]=scr0
scroll[2]=scr0
}else{
if(scr1!=scroll[2]){
//右がスクロール
frames[0].scrollTo(scr1,y0)
scroll[0]=scr1
scroll[2]=scr1
}}
setTimeout("scrollnn()",500)
}

//--> </script>
