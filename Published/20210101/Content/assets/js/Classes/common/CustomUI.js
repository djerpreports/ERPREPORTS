/*****************************************
A. Name: Custom UI Class
B. Date Created: Aug 09, 2020
C. Author: Jay-R A. Magdaluyo
D. Modification History:
E. Synopsis: Class Module used to Create custom UI features
***********************************************/
; (function () {
    const CustomUI = function () {
        return new CustomUI.init();
    }
    CustomUI.init = function () {
        this.dataTableID = "";
        this.notifTable = "";
    }
    CustomUI.prototype = {
        setDatatableMaxHeight: function () {
            if (this.dataTableID) {
                var $closestDiv = $(this.dataTableID).closest(".col-md-12");
                if ($closestDiv.length) {
                    var eTop = $($closestDiv).offset().top;
                    var windowHeight = $(window).height();
                    var containerID = this.dataTableID + "-container";
                    var adjust = $(this.dataTableID).data("adjust") || 0;
                    var pagingHeight = $(this.dataTableID + "_wrapper .pagination").height();
                    var divHeight = windowHeight - eTop - 70 - pagingHeight + adjust;
                    var containerHeight = divHeight;

                    $($closestDiv).attr("id", containerID.replace('#', ''));
                    $($closestDiv).css("max-height", containerHeight + "px");
                    $($closestDiv).addClass("fixed-tblcontainer");
                } else if ($("#tblPriceEstimation_filter").length) {
                    var eTop = $("#tblPriceEstimation_filter").offset().top;
                    var windowHeight = $(window).height();
                    var containerID = this.dataTableID + "-container";
                    var adjust = $(this.dataTableID).data("adjust") || 0;
                    var pagingHeight = $(this.dataTableID + "_wrapper .pagination").height();
                    var divHeight = windowHeight - eTop - 70 - pagingHeight + adjust;
                    var containerHeight = divHeight;
                    if (!$(".tbl-container-here").length) $(this.dataTableID).wrap("<div class='tbl-container-here'></div>");
                    $(".tbl-container-here").attr("id", containerID.replace('#', ''));
                    $(".tbl-container-here").css("max-height", containerHeight + "px");
                    $(".tbl-container-here").addClass("fixed-tblcontainer");
                }
            }
        },
        setDatatableMaxHeightFixed: function () {
            if (this.dataTableID) {
                var $closestDiv = $(this.dataTableID + "_wrapper");
                if ($closestDiv.length) {
                    var eTop = $($closestDiv).offset().top || 0;
                    var windowHeight = $(window).height() || 0;
                    var containerID = this.dataTableID + "-container" || 0;
                    var adjust = $(this.dataTableID).data("adjust") || 0 || 0;
                    var pagingHeight = $(this.dataTableID + "_wrapper .pagination").height() || 0;
                    var divHeight = 0;
                    var docWidth = $(document).width() || 0;

                    if (docWidth >= 767)
                        divHeight = windowHeight - eTop - 70 - pagingHeight + adjust;
                    else if (docWidth >= 576 && docWidth < 767)
                        divHeight = windowHeight - eTop - pagingHeight + adjust;
                    else if (docWidth < 576)
                        divHeight = windowHeight - eTop - pagingHeight + adjust;
                    if ($(this.dataTableID + " > tfoot").length || $(this.dataTableID + " > thead").length)
                        divHeight = divHeight - 31;
                    $(this.dataTableID + "_wrapper > div:nth-child(2) > div > div.dataTables_scroll > div.dataTables_scrollBody").css("height", "100%");
                    $(this.dataTableID + "_wrapper > div:nth-child(2) > div > div.dataTables_scroll > div.dataTables_scrollBody").css("max-height", divHeight + "px");
                }
            }
        },
        setDivMaxHeight: function (ID) {
            if (ID) {
                if (ID.length) {
                    var eTop = $(ID).offset().top;
                    var windowHeight = $(window).height();
                    var adjust = $(ID).data("adjust") || 0;
                    var divHeight = windowHeight - eTop - 70 + adjust;
                    var containerHeight = divHeight;
                    $(ID).css("max-height", containerHeight + "px");
                    $(ID).css("overflow-y", "auto");
                }
            }
        },
        createSelectOption: function (selectOptionsList) {
            var options = "<option value=''>--Please Select--</option>";
            if (selectOptionsList.length) {
                $.each(selectOptionsList, function (i, x) {
                    options += '<option value="' + x.value + '">' + x.text + '</option>';
                });
            }
            return options;
        },
        createSelect2: function (arrID, arrList) {
            var self = this;
            $.each(arrID, function (i, val) {
                if (arrList[i].length > 0) {
                    $(val).html(self.createSelectOption(arrList[i]));
                }
                $(val).select2({
                    placeholder: '--Please Select--',
                    allowClear: true
                });
            })
        },
        clearCustomError: function (id) {
            $("#" + id).removeClass("input-error");
            $("#err-" + id).text("");
            $("#err-" + id).removeClass("text-danger")
        },
        openCreatePanel: function () {
            var bodyID = $(".btnCreateData").data("panelbodyid");
            $(bodyID).show();
            $(bodyID).removeClass("tago");
            $(".btnCreateData")[0].children[0].className = "fa fa-minus";
            $(".btnCreateData").prop("title", "Collapse");
            this.setDatatableMaxHeight();
        },
        closeCreatePanel: function () {
            var bodyID = $(".btnCreateData").data("panelbodyid");
            $(bodyID).hide();
            $(bodyID).addClass("tago");
            $(".btnCreateData")[0].children[0].className = "fa fa-plus";
            $(".btnCreateData").prop("title", "Create");
            this.setDatatableMaxHeight();
        },
    }
    CustomUI.init.prototype = CustomUI.prototype;
    return window.CustomUI = window.$UI = CustomUI;
}());

const CUI = $UI();
$(document).ready(function () {
    $(window).resize(function () {
        CUI.setDatatableMaxHeightFixed();
    });
    $('.tabs-with-datatable .nav-tabs a').on('shown.bs.tab', function (event) {
        CUI.dataTableID = "#tbl" + $(this).attr("href").replace("#", "");
        if ($.fn.DataTable.isDataTable(CUI.dataTableID)) {
            CUI.setDatatableMaxHeightFixed();
        }
    });
    if ($("#iziModalError").length) {
        $(document).on('closing', '#iziModalError', function (e) {
            window.location.href = "/login";
        });
    }
    $('#loading_modal').on('hidden.bs.modal', function () {
        if ($('.modal:visible').length) {
            $('body').addClass('modal-open');
        }
    });
});
