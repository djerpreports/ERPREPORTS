/*****************************************
A. Name: Format Data Class
B. Date Created: Aug 09, 2020
C. Author: Jay-R A. Magdaluyo
D. Modification History:
E. Synopsis: Class Module used to format data
***********************************************/
; (function () {
    const FormatterClass = function (dataToFormat) {
        return new FormatterClass.init(dataToFormat);
    }
    FormatterClass.init = function (dataToFormat) {
        this.dataToFormat = dataToFormat;
        this.months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    }
    FormatterClass.prototype = {
        formatDate: function (dateFormat) {
            if (this.dataToFormat) {
                if (dateFormat) {
                    if (dateFormat === "dd/mm/yyyy") {
                        var today = new Date(this.dataToFormat);
                        var dd = today.getDate();
                        var mm = today.getMonth() + 1; //January is 0!

                        var yyyy = today.getFullYear();
                        if (dd < 10) {
                            dd = '0' + dd;
                        }
                        if (mm < 10) {
                            mm = '0' + mm;
                        }
                        var today = dd + '/' + mm + '/' + yyyy;
                    }
                    if (dateFormat === "mm/dd/yyyy") {
                        var today = new Date(this.dataToFormat);
                        var dd = today.getDate();
                        var mm = today.getMonth() + 1; //January is 0!

                        var yyyy = today.getFullYear();
                        if (dd < 10) {
                            dd = '0' + dd;
                        }
                        if (mm < 10) {
                            mm = '0' + mm;
                        }
                        var today = mm + '/' + dd + '/' + yyyy;
                    }
                    if (dateFormat === "mmyyyy") {
                        var today = new Date(this.dataToFormat);
                        var dd = today.getDate();
                        var mm = today.getMonth() + 1; //January is 0!

                        var yyyy = today.getFullYear();
                        if (dd < 10) {
                            dd = '0' + dd;
                        }
                        if (mm < 10) {
                            mm = '0' + mm;
                        }
                        var today = mm + '' + yyyy;
                    }
                    return today;
                } else {
                    return this.dataToFormat;
                }

            } else {
                return "";
            }
        },
        getFullMotnName: function (m) {
            return this.months[m];
        },
        formatMoney: function (decimalPlaces) {
            this.addZeros();
            return this.dataToFormat.toFixed(2).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
            return this.formatDecimal(2);
        },
        formatDecimal: function (decimalPlaces) {

            if (decimalPlaces > 0) {
                this.addZeros();
                this.dataToFormat = this.dataToFormat.toFixed(decimalPlaces);
                var arrParts = this.dataToFormat.split('.');
                var p1 = arrParts[0].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
                return p1 + "." + arrParts[1];
            } else {
                return this.dataToFormat.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
            }


        },
        addZeros: function () {
            var value = Number(this.dataToFormat);
            var res = ("'" + this.dataToFormat + "'").search(".");
            if (res.length < 0) {
                value = value + "00";
            }
            return value;
        },
        setInputFilter: function (textbox, inputFilter) {
            var id = "";
            if (textbox.length) {
                ["input", "keydown", "keyup", "mousedown", "mouseup", "select", "contextmenu", "drop"].forEach(function (event) {
                    id = textbox[0].id;
                    if (id) {
                        el = document.getElementById(id);
                        el.addEventListener(event, function () {
                            if (inputFilter(this.value)) {
                                this.oldValue = this.value;
                                this.oldSelectionStart = this.selectionStart;
                                this.oldSelectionEnd = this.selectionEnd;
                            } else if (this.hasOwnProperty("oldValue")) {
                                this.value = this.oldValue;
                                this.setSelectionRange(this.oldSelectionStart, this.oldSelectionEnd);
                            }
                        });
                    }
                });
            }
        },
        ConvertMinutes: function (num) {
            d = Math.floor(num / 1440); // 60*24
            h = Math.floor((num - (d * 1440)) / 60);
            m = Math.round(num % 60);

            if (d > 0) {
                return (d + " days, " + h + " hours, " + m + " minute" + ((m > 1) ? "s" : ""));
            } else {
                return (h + " hours, " + m + " minute" + ((m > 1) ? "s" : ""));
            }
        }
    }
    FormatterClass.init.prototype = FormatterClass.prototype;

    FormatterClass.init.prototype = FormatterClass.prototype;
    return window.FormatterClass = window.$F = FormatterClass;
}());
$(document).ready(function () {
    $("body").on("keypress", ".decimal-only, .money-only, .decimal", function (evt) {
        var charCode = (evt.which) ? evt.which : evt.keyCode;
        if (charCode != 46 && charCode > 31
          && (charCode < 48 || charCode > 57))
            return false;

        return true;
    });
    $(".decimal-only").change(function (evt) {
        return $(this).val($F(+$(this).val()).formatDecimal(4));
    });
    $(".money-only").change(function (evt) {
        return $(this).val($F(+$(this).val()).formatMoney(2));
    });
    $("body").on("change, focusout", ".amountInputValidate", function () {
        var val = $(this).val();
        if (val.search(",") >= 0) {
            val = val.replace(/,/g, '')
        }
        if (+val <= 0) {
            $(this).addClass("input-error");
            $("#err-" + $(this).attr("id")).text("This value should be greater than 0.");
            $("#err-" + $(this).attr("id")).addClass("text-danger");
            $("#err-" + $(this).attr("id")).addClass("text-right");
            $(this).val("");
        } else {
            $(this).removeClass("input-error");
            $("#err-" + $(this).attr("id")).text("");
            $("#err-" + $(this).attr("id")).removeClass("text-danger")
        }
    });
    $F().setInputFilter(document.getElementsByClassName("posNegOnly"), function (value) {
        return /^-?\d*$/.test(value);
    });
    $F().setInputFilter(document.getElementsByClassName("posOnly"), function (value) {
        return /^\d*$/.test(value);
    });
    $F().setInputFilter(document.getElementsByClassName("intLimit"), function (value) {
        return /^\d*$/.test(value) && (value === "" || parseInt(value) <= 500);
    });
    $F().setInputFilter(document.getElementsByClassName("floatNum"), function (value) {
        return /^-?\d*[.,]?\d*$/.test(value);
    });
    $F().setInputFilter(document.getElementsByClassName("currency2Dec"), function (value) {
        return /^-?\d*[.,]?\d{0,2}$/.test(value);
    });
    $F().setInputFilter(document.getElementsByClassName("hexa"), function (value) {
        return /^[0-9a-f]*$/i.test(value);
    });

});
