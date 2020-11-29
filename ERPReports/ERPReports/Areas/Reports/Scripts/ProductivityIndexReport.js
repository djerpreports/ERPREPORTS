"use strict";
(function () {
    const ProductivityIndexReport = function () {
        return new ProductivityIndexReport.init();
    }
    ProductivityIndexReport.init = function () {
        $D.init.call(this);
        this.$tblProductivityIndexReport = "";
        this.ID = 0;
    }
    ProductivityIndexReport.prototype = {
        drawDatatables: function () {
            var self = this;
            //if (!$.fn.DataTable.isDataTable('#tblProductivityIndexReport')) {
            //    self.$tblProductivityIndexReport = $('#tblProductivityIndexReport').DataTable({
            //        processing: true,
            //        serverSide: true,
            //        "order": [[0, "asc"]],
            //        "pageLength": 25,
            //        "ajax": {
            //            "url": "/MasterMaintenance/ProductivityIndexReport/GetProductivityIndexReportList",
            //            "type": "POST",
            //            "datatype": "json",
            //            "data": function (d) {
            //                $('#tblProductivityIndexReport thead #trSearch th').each(function () {
            //                    var field = $(this).data("field");
            //                    d[field] = $(this).find('select').val();
            //                });
            //            }
            //        },
            //        dataSrc: "data",
            //        scrollY: '100%', scrollX: '100%',
            //        select: true,
            //        columns: [
            //            { title: "ProductivityIndexReportName", data: "ProductivityIndexReportName" },
            //            { title: "First Name", data: "FirstName" },
            //            { title: "Middle Name", data: "MiddleName" },
            //            { title: "Last Name", data: "LastName" },
            //            { title: "Email Address", data: 'EmailAddress' },
            //        ],
            //        "createdRow": function (row, data, dataIndex) {
            //            $(row).attr('data-id', data.ID);
            //            $(row).attr('data-username', data.ProductivityIndexReportName);
            //        },
            //    })
            //}
            return this;
        },
    }
    ProductivityIndexReport.init.prototype = $.extend(ProductivityIndexReport.prototype, $D.init.prototype);
    ProductivityIndexReport.init.prototype = ProductivityIndexReport.prototype;

    $(document).ready(function () {
        var PIR = ProductivityIndexReport();
        PIR.drawDatatables();
        $("#MonthYear").change(function(){
            if($("#MonthYear").val()){
                $("#btnDownloadProductivityIndexReport").prop("disabled",false);
            } else {
                $("#btnDownloadProductivityIndexReport").prop("disabled", true);
            }
        });
        $("#MonthYear").datepicker({
            format: "M-yyyy",
            viewMode: "months",
            minViewMode: "months",
            autoclose:true,
        });
    });
})();