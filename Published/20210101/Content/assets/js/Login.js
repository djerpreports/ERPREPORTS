/*****************************************
A. Name: Login Script
B. Date Created: Nov 15, 2020
C. Author: Jay-R A. Magdaluyo
D. Modification History:
E. Synopsis: For User login
***********************************************/
'use strict';
(function () {
    var Login = $D();
    $(document).ready(function () {
        $("#UserName").focus();
        $("#frmLogin").submit(function (e) {
            e.preventDefault();
            LoginMeIn();
        });
        //All Function --------------------------------------------------------------------------------
        function LoginMeIn() {
            Login.formData = $('#frmLogin').serializeArray();
            Login.setJsonData();
            Login.formAction = '/Login/LoginEntry';
            Login.sendData().then(function () {
                var login = Login.responseData;
                if (login.error) {
                    Login.showError(login.errmsg);
                    $("#UserName").addClass("input-error");
                    $("#Password").addClass("input-error");
                    $("#UserName").addClass("parsley-success");
                    $("#Password").addClass("parsley-success");
                    $("#Password").val("");
                } else {
                    $("#UserName").removeClass("input-error");
                    $("#Password").removeClass("input-error");
                    $("#frmLogin > div.login-buttons > button").attr("disabled", true);
                    window.location = "/";
                }
            });
        }
    });
})();


window.history.forward();
function noBack() {
    window.history.forward();
}
setTimeout("noBack()", 0);
window.onunload = function () { null };
