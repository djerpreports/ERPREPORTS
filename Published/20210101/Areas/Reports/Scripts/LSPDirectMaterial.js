"use strict";
(function () {
    const LSPDirectMaterial = function () {
        return new LSPDirectMaterial.init();
    }
    LSPDirectMaterial.init = function () {
        $D.init.call(this);
        this.$tblLSPDirectMaterial = "";
        this.ID = 0;
    }
    LSPDirectMaterial.prototype = {
        drawDatatables: function () {
            var self = this;
            //if (!$.fn.DataTable.isDataTable('#tblLSPDirectMaterial')) {
            //    self.$tblLSPDirectMaterial = $('#tblLSPDirectMaterial').DataTable({
            //        processing: true,
            //        serverSide: true,
            //        "order": [[0, "asc"]],
            //        "pageLength": 25,
            //        "ajax": {
            //            "url": "/MasterMaintenance/LSPDirectMaterial/GetLSPDirectMaterialList",
            //            "type": "POST",
            //            "datatype": "json",
            //            "data": function (d) {
            //                $('#tblLSPDirectMaterial thead #trSearch th').each(function () {
            //                    var field = $(this).data("field");
            //                    d[field] = $(this).find('select').val();
            //                });
            //            }
            //        },
            //        dataSrc: "data",
            //        scrollY: '100%', scrollX: '100%',
            //        select: true,
            //        columns: [
            //            { title: "LSPDirectMaterialName", data: "LSPDirectMaterialName" },
            //            { title: "First Name", data: "FirstName" },
            //            { title: "Middle Name", data: "MiddleName" },
            //            { title: "Last Name", data: "LastName" },
            //            { title: "Email Address", data: 'EmailAddress' },
            //        ],
            //        "createdRow": function (row, data, dataIndex) {
            //            $(row).attr('data-id', data.ID);
            //            $(row).attr('data-username', data.LSPDirectMaterialName);
            //        },
            //    })
            //}
            return this;
        },
    }
    LSPDirectMaterial.init.prototype = $.extend(LSPDirectMaterial.prototype, $D.init.prototype);
    LSPDirectMaterial.init.prototype = LSPDirectMaterial.prototype;

    $(document).ready(function () {
        var LSPDM = LSPDirectMaterial();
        LSPDM.drawDatatables();

        $(".ReportType").click(function () {
            if ($('.ReportTypeG1:checked').length) {
                //$("#StartDate,#EndDate").attr("required",true);
                if ($("#StartDate").val() && $("#EndDate").val()) {
                    $("#btnPrint").prop("disabled", false);
                } else {
                    $("#btnPrint").prop("disabled", true);
                }
            } else if ($('.ReportTypeG2:checked').length) {
                $("#btnPrint").prop("disabled", false);
            } else {
                $("#btnPrint").prop("disabled", true);
            }
        });
        $("#StartDate,#EndDate").change(function () {
            if ($("#StartDate").val() && $("#EndDate").val()) {
                $("#btnPrint").prop("disabled", false);
            } else {
                $("#btnPrint").prop("disabled", true);
            }
        });
        $(".ReportTypeG1").click(function () {
            if ($('.ReportTypeG1:checked').length) {
                $("#StartDate,#EndDate").prop("disabled", false);
            } else {
                $("#StartDate,#EndDate").val("").prop("disabled", true);
                $("#ProductCode1,#ProductCode2,#Model1,#Model2").val("").trigger("change.select2");
            }
            if ($('#DMAndLaborPercentageReport').is(":checked")) {
                $("#ProductCode1,#ProductCode2,#Model1,#Model2").prop("disabled", false);
            } else {
                $("#ProductCode1,#ProductCode2,#Model1,#Model2").prop("disabled", true);
            }
        });
        $("#StartDate").datepicker({
            todayHighlight: true,
            autoclose: true,
        });
        $("#StartDate").change(function () {
            var minDate = new Date($(this).val());
            var lastDay = new Date(minDate.getFullYear(), minDate.getMonth() + 1, 0);
            minDate.setDate(minDate.getDate() + 1)
            $("#EndDate").prop("disabled", false).val('');
            $("#EndDate").datepicker('destroy');
            $("#EndDate").datepicker({
                startDate: minDate,
                endDate: lastDay,
                todayHighlight: true,
                autoclose: true,
            });
        });
        $("#btnPrint").click(function (e) {
            var checkedCount = $(".ReportType:checked").length;
            var myInterval = setInterval(submitForm, 1000);
            var arrSubmittedURL = [];
            var intervalCounter = 0;
            function submitForm() {
                intervalCounter++;
                if (intervalCounter == checkedCount) {
                    clearInterval(myInterval);
                }
                $(".ReportType").each(function () {
                    if ($(this).is(":checked")) {
                        var url = $(this).attr("data-url");
                        var foundURL = arrSubmittedURL.indexOf(url);
                        if (foundURL < 0) {
                            $("#frmReport").attr("action", url);
                            $("#frmReport").submit();
                            arrSubmittedURL.push(url);
                        }
                    }
                })
            }
        });
        $("#StartDate,#EndDate").prop("disabled", true);
        $('#ProductCode1,#ProductCode2').select2({
            ajax: {
                url: "/General/GetSelect2Data",
                data: function (params) {
                    var search = params.term || "";
                    return {
                        q: params.term,
                        id: 'product_code',
                        text: 'description',
                        table: 'prodcode',
                        db: 'LSPI803_App',
                        display: 'id&id-text',
                        query: "SELECT product_code, description FROM prodcode WHERE product_code LIKE 'FG-%' AND (product_code like '%" + search + "%' OR description like '%" + search + "%')",
                    };
                },
            },
            placeholder: '--Please Select--',
        }).prop("disabled", true);
        $("#Model1,#Model2").select2({
            ajax: {
                url: "/Reports/LSPDirectMaterial/GetSelect2DataModel",
                data: function (params) {
                    var search = params.term || "";
                    var ProductCode1 = $('#ProductCode1').val() || "";
                    var ProductCode2 = $('#ProductCode2').val() || "";
                    return {
                        q: params.term,
                        id: 'item',
                        text: 'description',
                        table: 'prodcode',
                        db: 'LSPI803_App',
                        display: 'id&text',
                        sp: "LSP_GetFGItemListPerProdCodeWihtNullSp",
                        StartProdCode: ProductCode1,
                        EndProdCode: ProductCode2,
                    };
                },
            },
            placeholder: '--Please Select--',
        }).prop("disabled", true);
    });
})();