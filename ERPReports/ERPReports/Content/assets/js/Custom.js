"use strict";
(function () {
    $(document).ready(function () {
        $(".btnCreateData").click(function () {
            var bodyID = $(this).data("panelbodyid");
            var isVisible = $(bodyID + ":visible").length;
            $('[data-toggle="tooltip"]').tooltip('dispose');
            if (isVisible) {
                $(this)[0].children[0].className = "fa fa-plus";
                $(this).prop("title", "Create");
            } else {
                $(this)[0].children[0].className = "fa fa-minus";
                $(this).prop("title", "Collapse");
            }
            $('[data-toggle="tooltip"]').tooltip();
        });
        $('[data-toggle="tooltip"]').tooltip();

    });
    //#Functions Here-------------------------------------------------------------------------------------------------------------------

})();