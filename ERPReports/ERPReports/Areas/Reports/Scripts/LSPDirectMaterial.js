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

        $("#StartDate,#EndDate").datepicker({
            autoclose: true,
            todayHighlight: true
        });
        $("#ProductCode1,#ProductCode2,#Model1,#Model2").select2({
            placeholder: "--Please Select--"
        });
    });
})();