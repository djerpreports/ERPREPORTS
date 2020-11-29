/*****************************************
A. Name: Message Class
B. Date Created: Aug 09, 2020
C. Author: Jay-R A. Magdaluyo
D. Modification History:
E. Synopsis: Class Module used for showing messages
***********************************************/
;
(function (global, $) {
    const Message = function (msgType, msgTitle, msg, func, trxId) {
        return new Message.init(msgType, msgTitle, msg, func, trxId);
    }
    Message.init = function (msgType, msgTitle, msg, func, trxId) {
        this.msgType = msgType || "";
        this.msgTitle = msgTitle || "";
        this.msg = msg || "";
        this.trxId = trxId || "";
        this.func = func || "";

        this.msgTimer = 5000;
        this.msgToastrType = {
            "info": "fa fa-info-circle",
            "warning": "fas fa-lg fa-fw m-r-10 fa-exclamation-triangle",
            "success": "fa fa-check-square-o",
            "error": "fa fa-times-circle",
        };
        this.txtColor = {
            "info": "blue",
            "warning": "yellow",
            "success": "green",
            "error": "red",
        };
    }
    Message.prototype = {
        getError: function () {
            var errorList = "";
            if (Array.isArray(this.msg) || typeof this.msg === "object") {
                $.each(this.msg, function (key, value) {
                    errorList += '<div>' + (key + 1) + ". " + value + '</div>';
                });
            } else {
                errorList += '<div>' + this.msg + '</div>';
            }
            this.msg = errorList;
        },
        showToastrMsg: function () {
            if (this.msgType === "error") {
                this.getError();
            }
            iziToast.show({
                title: this.msgTitle,
                message: this.msg,
                icon: this.msgToastrType[this.msgType],
                position: 'topRight',
                backgroundColor: '',
                theme: 'light', // dark
                color: this.txtColor[this.msgType], // blue, red, green, yellow
                timeout: this.msgTimer,
            });
            return this;
        },
        showToastrMsgNoClose: function () {
            if (this.msgType === "error") {
                this.getError();
            }
            iziToast.show({
                title: this.msgTitle,
                message: this.msg,
                icon: this.msgToastrType[this.msgType],
                position: 'topRight',
                backgroundColor: '',
                theme: 'light', // dark
                color: this.txtColor[this.msgType], // blue, red, green, yellow
                timeout: 0,
            });
            return this;
        },
        showInfo: function (msg) {
            this.msgType = "info";
            this.msgTitle = "Message!";
            this.msg = msg;
            this.showToastrMsg();
        },
        showError: function (msg) {
            this.msgType = "error";
            this.msgTitle = "Error!";
            this.msg = msg;
            this.showToastrMsg();
        },
        showErrorNoClose: function (msg) {
            this.msgType = "error";
            this.msgTitle = "Error!";
            this.msg = msg;
            this.showToastrMsgNoClose();
        },
        showSuccess: function (msg) {
            this.msgType = "success";
            this.msgTitle = "Success!";
            this.msg = msg;
            this.showToastrMsg();
        },
        showWarning: function (msg) {
            this.msgType = "warning";
            this.msgTitle = "Warning!";
            this.msg = msg;
            this.showToastrMsg();
        },
        confirmAction: function () {
            var self = this;
            var promiseObj = new Promise(function (resolve, reject) {
                iziToast.question({
                    timeout: 20000,
                    close: false,
                    overlay: true,
                    displayMode: 'once',
                    id: 'question',
                    zindex: 99999999,
                    title: 'Confirm: ',
                    message: self.msg,
                    position: 'topCenter',
                    timeout: 0,
                    buttons: [
                        ['<button>Yes</button>', function (instance, toast) {

                            instance.hide({ transitionOut: 'fadeOut' }, toast, 'button');
                            resolve(true);

                        }],
                        ['<button><b>NO</b></button>', function (instance, toast) {
                            instance.hide({ transitionOut: 'fadeOut' }, toast, 'button');
                            resolve(false);
                        }, true],
                    ],
                });
            });
            return promiseObj;
        },
        confirmActionNo: function () {
            var self = this;
            var promiseObj = new Promise(function (resolve, reject) {
                iziToast.question({
                    timeout: 20000,
                    close: false,
                    overlay: true,
                    displayMode: 'once',
                    id: 'question',
                    zindex: 99999999,
                    title: 'Confirm: ',
                    message: self.msg,
                    position: 'topCenter',
                    timeout: 0,
                    buttons: [
                        ['<button>Yes</button>', function (instance, toast) {

                            instance.hide({ transitionOut: 'fadeOut' }, toast, 'button');
                            resolve(true);

                        }],
                        ['<button><b>NO</b></button>', function (instance, toast) {
                            resolve(false);
                            instance.hide({ transitionOut: 'fadeOut' }, toast, 'button');
                        }, true],
                    ],
                });
            });
            return promiseObj;
        },
        showSessionError: function (response) {
            $("#iziModalError").iziModal({
                title: "Attention",
                subtitle: response.errors,
                icon: 'icon-power_settings_new',
                headerColor: '#BD5B5B',
                width: 600,
                timeout: 0,
                timeoutProgressbar: true,
                transitionIn: 'fadeInDown',
                transitionOut: 'fadeOutDown',
                pauseOnHover: true
            });
            $('#iziModalError').iziModal('open');
        },
    }
    Message.init.prototype = Message.prototype;
    return global.Message = global.$M = Message;
}(window, $));
