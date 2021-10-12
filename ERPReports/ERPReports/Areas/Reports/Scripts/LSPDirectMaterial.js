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
        validateRMBreakdownPerJOReport: function () {
            var self = this;
            var PONumber = $("#PONumber").val() || "";
            var JONumber = $("#JONumber").val() || "";

            if (PONumber == "" && JONumber == "") {
                self.showError("Please enter PO Number or JO NUmber. Thank you.");
                return false;
            }
            else {
                return true;
            }
        },
        validatePrint: function () {
            var PONumber = $("#PONumber").val() || "";
            var JONumber = $("#JONumber").val() || "";
            var ProductCode = $("#ProductCode").val() || "";
            var TransactionDate = $("#TransactionDate").val() || "";
            var isValid = true;
            if ($('.ReportTypeG1:checked').length) {
                //$("#StartDate,#EndDate").attr("required",true);
                if ($("#StartDate").val() && $("#EndDate").val()) {
                    isValid = true;
                } else {
                    isValid = false;
                }
            }
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

            if ($('.ReportTypeG2:checked').length && isValid) {
                isValid = true;
            }
            if ($('.ReportTypeG3:checked').length) {
                if ($('#Month').val() && isValid)
                    isValid = true;
                else {
                    isValid = false;
                }

                if ($("#SlowMonitoringAnalysisReport").is(":checked") && $("#Month").val() && isValid) {
                    isValid = true;
                } else {
                    isValid = false;
                }
            }
            if ($('.ReportTypeG4:checked').length) {
                if ((PONumber != "" || JONumber != "") && isValid) {
                    isValid = true;
                }
                else {
                    isValid = false;
                }

                $("#JONumber,#PONumber").prop("disabled", false);
            } else {
                $("#JONumber,#PONumber").val("").prop("disabled", true);
            }
            if ($('.ReportTypeG5:checked').length) {
                if (ProductCode != "" && TransactionDate != "" && isValid) {
                    isValid = true;
                }
                else {
                    isValid = false;
                }

                $("#ProductCode,#TransactionDate").prop("disabled", false);
            } else {
                $("#ProductCode,#TransactionDate").val("").prop("disabled", true);
                $("#ProductCode").trigger("change.select2");
            }
            if ($('.ReportTypeG6:checked').length && isValid) {
                $("#ShowDetailedTransaction").prop("disabled", false);
                isValid = true;
            } else {
                $("#ShowDetailedTransaction").prop("disabled", true);
            }
            if (isValid) {
                $("#btnPrint").prop("disabled", false);
            } else {
                $("#btnPrint").prop("disabled", true);
            }
        },
    }
    LSPDirectMaterial.init.prototype = $.extend(LSPDirectMaterial.prototype, $D.init.prototype);
    LSPDirectMaterial.init.prototype = LSPDirectMaterial.prototype;

    $(document).ready(function () {
        var LSPDM = LSPDirectMaterial();
        LSPDM.drawDatatables();

        $(".ReportType").click(function () {
            LSPDM.validatePrint();
        });
        $("#StartDate,#EndDate").change(function () {
            LSPDM.validatePrint();
        });
        $(".ReportTypeG1").click(function () {
            LSPDM.validatePrint();
        });
        $("#StartDate").datepicker({
            todayHighlight: true,
            autoclose: true,
            format: 'MM dd, yyyy',
        });
        $("#StartDate").change(function () {
            var formattedStartDate = $F($(this).val()).formatDate("mm/dd/yyyy");
            var minDate = new Date(formattedStartDate);
            var lastDay = new Date(minDate.getFullYear(), minDate.getMonth() + 1, 0);
            minDate.setDate(minDate.getDate())
            $("#EndDate").prop("disabled", false).val('');
            $("#EndDate").datepicker('destroy');
            $("#EndDate").datepicker({
                startDate: minDate,
                endDate: lastDay,
                todayHighlight: true,
                autoclose: true,
                format: 'MM dd, yyyy',
            });
        });
        $("#TransactionDate").datepicker({
            todayHighlight: true,
            autoclose: true,
            format: 'MM dd, yyyy',
        });
        $("#btnPrint").click(function (e) {
            var isValid = true;
            if ($("#RMBreakdownPerJOReport").is(":checked"))
                isValid = LSPDM.validateRMBreakdownPerJOReport();

            if (!isValid) {
                return;
            }

            var checkedCount = $(".ReportType:checked").length;
            var myInterval = setInterval(submitForm, 1);
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
        $("#Month").change(function () {
            LSPDM.validatePrint();
        });
        $("#SlowMonitoringAnalysisReport").change(function () {
            if ($("#SlowMonitoringAnalysisReport").is(":checked")) {
                $("#Month").prop("disabled", false);
            } else {
                $("#Month").prop("disabled", true);
            }
        });
        $("#JONumber,#PONumber").change(function () {
            LSPDM.validatePrint();

        });
        $("#JONumber").change(function () {
            var JONumber = $(this).val() || "";

            if (JONumber != "") {
                var arrJO = JONumber.split("-");
                var isFormatCorrect = arrJO.length == 2;
                if (isFormatCorrect) {
                    var JOPrefix = arrJO[0];
                    var JOSuffix = arrJO[1];
                    if (JOSuffix.length) {
                        var JONumberLn = JOPrefix.length;
                        var remainingCount = 9 - JONumberLn;

                        var newJONumber = JOPrefix + "-" + padLeadingZeros(JOSuffix, remainingCount);
                        $("#JONumber").val(newJONumber.toUpperCase());
                    } else {
                        $("#btnPrint").prop("disabled", true);
                    }
                } else {
                    $("#btnPrint").prop("disabled", true);
                }
            }

            function padLeadingZeros(num, size) {
                var s = num + "";
                while (s.length < size) s = "0" + s;
                return s;
            }
        });
        $("#ProductCode,#TransactionDate").change(function () {
            LSPDM.validatePrint();

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
                        display: 'id&id-text',
                        sp: "LSP_GetFGItemListPerProdCodeWihtNullSp",
                        StartProdCode: ProductCode1,
                        EndProdCode: ProductCode2,
                    };
                },
            },
            placeholder: '--Please Select--',
        }).prop("disabled", true);
        $("#ProductCode").select2({
            ajax: {
                url: "/General/GetSelect2Data",
                data: function (params) {
                    var search = params.term || "";
                    return {
                        q: params.term,
                        id: 'ProductCode',
                        text: 'Description',
                        table: 'prodcode',
                        db: 'ERPReports',
                        display: 'id&text',
                        query: "EXEC LSPI803_App.dbo.LSP_NewDM_GetAllRMProductCodesGroupedSp",
                    };
                },
            },
            placeholder: '--Please Select--',
        }).prop("disabled", true);
    });
})();
