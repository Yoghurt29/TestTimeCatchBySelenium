var username="Maple_Zeng";
var password="performance201705";

function login() {
    document.querySelectorAll("input[name='UserName']")[0].value = username;
    document.querySelectorAll("input[name='Password']")[0].value = password;
    document.querySelector("input[name='3.7.5.13']").click();
}

function goProductHistory() {
    document.querySelector("#product_history").click();
}

function clickSerch(sn) {
    document.querySelector("input[name='7.1.1']").value = sn;
    document.querySelector("input[name='7.1.3']").click();
}

function getLastTestTimestamp() {
    var timestamp = document.querySelector("form[name='f_13_1_5_15_1_3_0_1'] tbody:nth-child(1) tr:nth-child(2)  td:nth-child(17) font:nth-child(1)").innerHTML;
    return timestamp;
}

function getSnsFromPEMainBoardTrack(){
    var request=new XMLHttpRequest();
    request.open("get","http://172.28.136.19:8099/PEMainboardTrack/inOut/lendingBoardReport?isDelay=allDay&pageIndex=1&capacity=-1",false);
    request.send(null);
    return request.responseText;
}

function uploadResultByAjaxPost(data){
    var request=new XMLHttpRequest();
    request.open("post","http://172.28.136.19:8099/PEMainboardTrack/inOut/uploadActiveTimeByGETRequestPlayloadJSONData",false);
    request.setRequestHeader("Content-type","application/x-www-form-urlencoded");  
    request.send('activeTimeJSONData='+data);
    return request.responseText;
}
