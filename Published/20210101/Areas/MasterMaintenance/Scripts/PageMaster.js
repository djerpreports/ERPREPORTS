"use strict";
/*****************************************
A. Name: Page Master
B. Date Created: Aug 09, 2020
C. Author: Jay-R A. Magdaluyo
D. Modification History:
E. Synopsis: Module used to create Menu and URLs for  the users to access. Use this module after
                creating a new module to create its menu and add to access list of user at User Master Modules.
***********************************************/

(function () {
    const PageMaster = function () {
        return new PageMaster.init();
    }
    PageMaster.init = function () {
        $D.init.call(this);
        this.$tblPage = "";
        this.ID = 0;
        this.pageName = "";
    }
    PageMaster.prototype = {
        drawDatatables: function () {
            var self = this;
            if (!$.fn.DataTable.isDataTable('#tblPage')) {
                //For Footer Search
                //$('#tblPage tfoot th').each(function () {
                //    var title = $(this).text();
                //    $(this).html("<input type='text' class='form-control form-control-sm' placeholder='Search " + title + "' />");
                //});
                //End For Footer Search
                $('#tblPage thead tr').clone(true).appendTo('#tblPage thead');
                $('#tblPage thead tr:eq(1) th').each(function (i) {
                    var title = $(this).text();
                    $(this).html("<input type='text' class='form-control form-control-sm' placeholder='" + title + "' />");

                    $('input', this).on('keyup change', function () {
                        if (self.$tblPage.column(i).search() !== this.value) {
                            self.$tblPage
                                .column(i)
                                .search(this.value)
                                .draw();
                        }
                    });
                });

                self.$tblPage = $('#tblPage').DataTable({
                    processing: true,
                    serverSide: true,
                    select: true,
                    "order": [[0, "asc"]],
                    "pageLength": 25,
                    "ajax": {
                        "url": "/MasterMaintenance/PageMaster/GetPageList",
                        "type": "POST",
                        "datatype": "json",
                    },
                    orderCellsTop: true,
                    dataSrc: "data",
                    scrollY: '100%', scrollX: '100%',
                    columns: [
                        { title: "Group Label", data: "GroupLabel" },
                        { title: "Page Name", data: "PageName" },
                        { title: "Page Label", data: "PageLabel" },
                        { title: "URL", data: "URL" },
                        { title: "Has Sub", data: "HasSub" },
                        { title: "Parent Menu", data: "ParentMenu" },
                        { title: "ParentOrder", data: "ParentOrder" },
                        { title: "Order", data: "Order" },
                        { title: "Icon", data: "Icon" }
                    ],
                    "createdRow": function (row, data, dataIndex) {
                        $(row).attr('data-id', data.ID);
                        $(row).attr('data-PageName', data.PageName);
                    },
                    initComplete: function () {
                        //For Footer Search
                        //    this.api().columns().every(function () {
                        //        var that = this;
                        //        $('input', this.footer()).on('keyup change clear', function () {
                        //            if (that.search() !== this.value) {
                        //                that
                        //                    .search(this.value)
                        //                    .draw();
                        //            }
                        //        });
                        //    });
                        //End For Footer Search
                    }
                });
            }
            return this;
        },
        validatePageName: function (el) {
            var self = this;
            var pageName = $(el).val().trim();
            if (pageName) {
                self.formAction = '/MasterMaintenance/PageMaster/ValidatePageName';
                self.jsonData = { PageName: pageName }
                self.sendData().then(function () {
                    var validPageName = self.responseData.isValid;
                    if (!validPageName) {
                        self.showError("PageName already exists. Please try another PageName");
                        $(el).val("");
                        $(el).focus();
                    } else {
                        $("#Email").focus();
                    }
                });
            }
            return this;
        },
        savePage: function () {
            var self = this;
            self.formData = $('#frmPage').serializeArray();
            self.formAction = '/MasterMaintenance/PageMaster/SavePage';
            self.setJsonData().sendData().then(function () {
                self.$tblPage.ajax.reload(null, false);
                self.cancelPageTbl();
                self.cancelPageForm();
            });
            return this;
        },
        cancelPageForm: function () {
            var self = this;
            self.ID = 0;
            self.pageName = "";
            self.clearFromData("frmPage");
            $('#PageName').prop('readonly', false);
            $("#btnSavePage .btnLabel").text(" Save");
            $("#mdlPageTitle").text(" Create Page");
            $("#lblIcon").text("");
            $("#mdlPage").modal("hide");
            return this;
        },
        cancelPageTbl: function () {
            var self = this;
            $('#btnEditPage').attr("disabled", "disabled");
            $('#btnDeletePage').attr("disabled", "disabled");
            return this;
        },
        editPage: function () {
            var self = this;
            self.pageName = self.$tblPage.rows({ selected: true }).data()[0].PageName;
            self.jsonData = { PageName: self.pageName };
            self.formAction = '/MasterMaintenance/PageMaster/GetPageDetails';
            self.sendData().then(function () {
                self.populatePageData(self.responseData.pageDetails);
            });
            return this;
        },
        populatePageData: function (pageDetails) {
            var self = this;
            $('#PageName').prop('readonly', true);
            $("#frmPage").parsley().reset();
            $("#btnSavePage .btnLabel").text(" Update");
            $("#mdlPageTitle").text(" Edit Page");
            $("#mdlPage").modal("show");
            self.populateToFormInputs(pageDetails, "#frmPage");
            return this;
        },
        deletePage: function () {
            var self = this;
            self.pageName = self.$tblPage.rows({ selected: true }).data()[0].PageName;
            self.formAction = '/MasterMaintenance/PageMaster/DeletePage';
            self.jsonData = { PageName: self.pageName };
            self.sendData().then(function () {
                self.$tblPage.ajax.reload(null, false);
                self.cancelPageTbl();
                self.cancelPageForm();
            });
            return this;
        },
    }
    PageMaster.init.prototype = $.extend(PageMaster.prototype, $D.init.prototype);
    PageMaster.init.prototype = PageMaster.prototype;

    $(document).ready(function () {
        var Page = PageMaster();
        Page.drawDatatables();
        $("#btnNewPage").click(function () {
            $("#mdlPage").modal("show");
        });
        $("#PageName").change(function () {
            Page.validatePageName(this);
        });
        $("#Icon").change(function () {
            $("#lblIcon").attr("class", "");
            $("#lblIcon").addClass("float-right " + $(this).val());
        });
        $('#btnCancelPage').click(function () {
            Page.cancelPageForm();
        });
        $("#btnGeneratePassword").click(function () {
            $("#Password").val(generatePassword());
            Page.parsleyValidate("frmPage");
        });
        $("#btnSavePage").click(function () {
            var mode = $("#btnSavePage > span.btnLabel").text().toLowerCase();
            Page.msg = "Are you sure you want to " + mode + " this Page?";
            Page.confirmAction().then(function (approve) {
                if (approve)
                    $("#frmPage").submit();
            });
        });
        $("#frmPage").submit(function (e) {
            e.preventDefault();
            Page.savePage();
        });
        $('#btnEditPage').click(function () {
            Page.editPage();
        });
        $('#btnDeletePage').click(function () {
            Page.msg = "Are you sure you want to delete this Page?";
            Page.confirmAction().then(function (approve) {
                if (approve)
                    Page.deletePage();
            });
        });

        //#Special Events
        Page.$tblPage.on('draw.dt', function () {
            CUI.dataTableID = "#tblPage";
            CUI.setDatatableMaxHeightFixed();
        });
        Page.$tblPage.on('select', function (e, dt, type, indexes) {
            $('#btnEditPage').removeAttr("disabled");
            $('#btnDeletePage').removeAttr("disabled");
        });
        Page.$tblPage.on('deselect', function (e, dt, type, indexes) {
            $('#btnEditPage').attr("disabled", "disabled");
            $('#btnDeletePage').attr("disabled", "disabled");
        });
    });
})();
