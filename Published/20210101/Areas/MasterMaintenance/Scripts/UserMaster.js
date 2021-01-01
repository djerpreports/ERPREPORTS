"use strict";
/*****************************************
A. Name: Page Master
B. Date Created: Aug 09, 2020
C. Author: Jay-R A. Magdaluyo
D. Modification History:
E. Synopsis: Module used to set User Access of Modules
***********************************************/
(function () {
    const UserMaster = function () {
        return new UserMaster.init();
    }
    UserMaster.init = function () {
        $D.init.call(this);
        this.$tblUser = "";
        this.ID = 0;
        this.populateMode = false;
        this.costCenterSelectElementCounter = 1;
    }
    UserMaster.prototype = {
        drawDatatables: function () {
            var self = this;
            if (!$.fn.DataTable.isDataTable('#tblUser')) {
                self.$tblUser = $('#tblUser').DataTable({
                    processing: true,
                    serverSide: true,
                    "order": [[0, "asc"]],
                    "pageLength": 25,
                    "ajax": {
                        "url": "/MasterMaintenance/UserMaster/GetUserList",
                        "type": "POST",
                        "datatype": "json",
                        "data": function (d) {
                            $('#tblUser thead #trSearch th').each(function () {
                                var field = $(this).data("field");
                                d[field] = $(this).find('select').val();
                            });
                        }
                    },
                    dataSrc: "data",
                    scrollY: '100%', scrollX: '100%',
                    select: true,
                    columns: [
                        { title: "UserName", data: "UserName" },
                        { title: "First Name", data: "FirstName" },
                        { title: "Middle Name", data: "MiddleName" },
                        { title: "Last Name", data: "LastName" },
                        { title: "Email Address", data: 'EmailAddress' },
                    ],
                    "createdRow": function (row, data, dataIndex) {
                        $(row).attr('data-id', data.ID);
                        $(row).attr('data-username', data.UserName);
                    },
                })
            }
            return this;
        },
        validateUserName: function (el) {
            var self = this;
            var username = $(el).val().trim();
            if (username) {
                self.formAction = '/MasterMaintenance/UserMaster/ValidateUserName';
                self.jsonData = { UserName: username }
                self.sendData().then(function () {
                    var validUserName = self.responseData.isValid;
                    if (!validUserName) {
                        self.showError("UserName already exists. Please try another UserName");
                        $(el).val("");
                        $(el).focus();
                    } else {
                        $("#EmailAddress").focus();
                    }
                });
            }
            return this;
        },
        validateEmailAddress: function (el) {
            var self = this;
            var email = $(el).val().trim();
            if (email) {
                self.formAction = '/MasterMaintenance/UserMaster/ValidateEmailAddress';
                self.jsonData = { EmailAddress: email }
                self.sendData().then(function () {
                    var validEmailAddress = self.responseData.isValid;
                    if (!validEmailAddress) {
                        self.showError("EmailAddress already exists. Please try another EmailAddress");
                        $(el).val("");
                        $(el).focus();
                    } else {
                        $("#Section").focus();
                    }
                });
            }
            return this;
        },
        generatePassword: function () {
            var self = this;
            var length = 8,
                charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
                retVal = "";
            for (var i = 0, n = charset.length; i < length; ++i) {
                retVal += charset.charAt(Math.floor(Math.random() * n));
            }
            return retVal;
        },
        saveUser: function () {
            var self = this;
            var oldPassword = $("#Password").attr("data-oldpassword");
            var didPasswordChange = $("#Password").attr("data-didpasswordchange") || false;
            self.formData = $('#frmUser').serializeArray();
            self.formAction = '/MasterMaintenance/UserMaster/SaveUser';
            self.setJsonData();
            self.jsonData = { User: self.jsonData, OldPassword: oldPassword, DidPasswordChange: didPasswordChange }
            self.sendData().then(function () {
                self.$tblUser.ajax.reload(null, false);
                self.cancelUserTbl();
                self.cancelUserForm();
            });
            return this;
        },
        cancelUserForm: function () {
            var self = this;
            self.populateMode = false;
            self.clearFromData("frmUser");
            $("#UserID").val(0);
            $("#Password").attr("type", "text").attr("data-didpasswordchange", false).attr("readonly", false).val("");
            $('#UserName').prop('readonly', false);
            $("#mdlUserTitle").text(" Create User");
            $("#btnSaveUser .btnLabel").text(" Save");
            $("#Role option[value='Custom']").remove();
            $("#mdlUser").modal("hide");
            self.costCenterSelectElementCounter = 1;
            return this;
        },
        cancelUserTbl: function () {
            var self = this;
            $('#btnEditUser').attr("disabled", "disabled");
            $('#btnDeleteUser').attr("disabled", "disabled");
            $('#btnUserAccess').attr("disabled", "disabled");
            return this;
        },
        editUser: function () {
            var self = this;
            self.populateMode = true;
            var ID = self.$tblUser.rows({ selected: true }).data()[0].ID;
            self.formAction = '/MasterMaintenance/UserMaster/GetUserDetails';
            if (self.formAction) {
                self.jsonData = { ID: ID };
                self.sendData().then(function () {
                    self.populateUserData(self.responseData.userData);
                });
            } else {
                self.showError("Please try again.");
            }
            return this;
        },
        populateUserData: function (user) {
            var self = this;
            var sectionOption = "";
            var positionOption = "";
            $('#UserName').prop('readonly', true);
            $("#frmUser").parsley().reset();
            $("#mdlUserTitle").text(" Update User");
            $("#btnSaveUser .btnLabel").text(" Update");
            self.populateToFormInputs(user, "#frmUser");
            sectionOption = new Option(user.SectionDesc, user.Section, true, true);
            $('#Section').append(sectionOption).trigger('change');
            positionOption = new Option(user.PositionDesc, user.Position, true, true);
            $('#Position').append(positionOption).trigger('change');
            $("#Password").attr("type", "password").attr("data-oldpassword", user.Password).attr("readonly", true);
            $("#IsAdmin").val(user.IsAdmin ? "true" : "false");
            $("#IsObserver").val(user.IsObserver ? "true" : "false");
            $("#mdlUser").modal("show");
            return this;
        },
        deleteUser: function () {
            var self = this;
            var ID = self.$tblUser.rows({ selected: true }).data()[0].ID;
            self.formAction = '/MasterMaintenance/UserMaster/DeleteUser';
            self.jsonData = { ID: ID };
            self.sendData().then(function () {
                self.$tblUser.ajax.reload(null, false);
                self.cancelUserTbl();
                self.cancelUserForm();
            });
            return this;
        },
        getUserAccess: function () {
            var self = this;
            var ID = self.$tblUser.rows({ selected: true }).data()[0].ID;
            self.formAction = '/MasterMaintenance/UserMaster/GetUserAccess';
            self.jsonData = { ID: ID };
            self.sendData().then(function () {
                self.drawUserMenu(self.responseData);
            });
            return this;
        },
        drawUserMenu: function (menu) {
            var self = this;
            const tmpForParentMenu = menu;
            const tmpForSubMenu = menu;
            //Grouping menus into GroupLabel
            var arrGroupedData = _.mapValues(_.groupBy(menu, 'GroupLabel'), (function (clist) {
                return clist.map(function (byGroupLabel) {
                    return _.omit(byGroupLabel, 'GroupLabel');
                });
            }));
            //End Grouping menus into GroupLabel 
            //Creating GroupLabel Tab Link
            var menuTab = "<ul class='nav nav-tabs'>";
            $.each(arrGroupedData, function (i) {
                var GroupLabelID = i == "" ? "Common" : i;
                var active = i == "" ? "active" : "";
                menuTab += "<li class='nav-items'>\
                                    <a href='#tab" + GroupLabelID + "' data-toggle='tab' class='nav-link " + active + "'>\
                                        <span class='d-sm-none'>" + GroupLabelID + "</span>\
                                        <span class='d-sm-block d-none'> " + GroupLabelID + "</span>\
                                    </a>\
                                 </li>";
            });
            menuTab += "</ul>";
            //End Creating GroupLabel Tab Link

            //Creating GroupLabel Tab Content
            menuTab += "<div class='tab-content'>";
            $.each(arrGroupedData, function (i, groupedMainMenu) {
                var GroupLabelID = i == "" ? "Common" : i;
                var active = i == "" ? "active" : "";
                //Creating GroupLabel Individual Tab Content
                menuTab += "<div class='tab-pane show " + active + "' id='tab" + GroupLabelID + "'>";
                //Creating Parent Menu Tab Link
                menuTab += "    <ul class='nav nav-tabs'>";
                var parentMenu = _.filter(groupedMainMenu, function (obj) { return obj.Icon != 0; });
                var sortedParentMenu = _.sortBy(parentMenu, function (o) { return new Number(o.ParentOrder); }, ['asc']);
                $.each(sortedParentMenu, function (ii) {
                    var active1 = (ii == 0 && GroupLabelID == "Common") || (ii == 0 && GroupLabelID != "Common") ? 'active' : '';
                    if (this.URL != "/") {
                        menuTab += "<li class='nav-items'>\
                                    <a href='#tab" + (this.PageName).replace(/\s/g, "") + "' data-toggle='tab' class='nav-link " + active1 + "'>\
                                        <span class='d-sm-none'><span class='" + this.Icon + "'></span> " + this.PageLabel + "</span>\
                                        <span class='d-sm-block d-none'><span class='" + this.Icon + "'></span> " + this.PageLabel + "</span>\
                                    </a>\
                                 </li>";
                    }
                });
                menuTab += "    </ul>";
                //End Creating Parent Menu Tab Link
                //Creating Parent Menu Tab Content
                menuTab += "<div class='tab-content'>";
                $.each(sortedParentMenu, function (iii, m) {
                    if (this.URL != "/") {
                        var PageName = this.PageName;
                        var PageLabel = this.PageLabel;
                        var submenus = _.filter(groupedMainMenu, function (obj) {
                            return obj.ParentMenu == PageName && obj.Icon == 0;
                        });
                        var sortedSubmenu = _.sortBy(submenus, function (o) { return new Number(o.Order); }, ['asc']);

                        //var show = i == 1 ? 'fade active' : '';
                        var show = (iii == 0 && GroupLabelID == "Common") || (iii == 0 && GroupLabelID != "Common") ? 'fade active' : '';
                        var checked = this.Status != '' ? 'checked' : '';

                        var readOnly = this.ReadAndWrite ? '' : 'checked';
                        var readAndWrite = this.ReadAndWrite ? 'checked' : '';
                        var delChecked = this.Delete ? 'checked' : '';
                        //Creating indivdial menu list content
                        menuTab += "<div class='tab-pane show " + show + "' id='tab" + (PageName).replace(/\s/g, '') + "'>";
                        menuTab += "    <ul>";
                        menuTab += "        <li>";
                        menuTab += "            <div class='checkbox checkbox-css checkbox-inline'>\
                                                    <input type='checkbox'  id='chk" + PageName.replace(/\s/g, '') + "' value='" + this.ID + "' " + checked + " data-pagename='" + PageName.replace(/\s/g, '') + "'  data-hassub='" + this.HasSub + "'  data-readandwrite='" + this.ReadAndWrite + "' data-delete='" + this.Delete + "'   data-childclass='" + PageName.replace(/\s/g, '') + "' class='parentMenu chkUserAccess clickable'>\
                                                    <label class='clickable' for='chk" + PageName.replace(/\s/g, '') + "'>" + PageLabel + "</label>\
                                                </div>";
                        if (!this.HasSub) {
                            menuTab += "            <div class='radio radio-css radio-inline m-l-20'>\
                                                        <input class='form-check-input clickable' type='radio' name='rdoReadAndWrite" + PageName.replace(/\s/g, '') + "' id='rdoRead" + PageName.replace(/\s/g, '') + "' value='0' " + readOnly + ">\
                                                        <label class='form-check-label clickable' for='rdoRead" + PageName.replace(/\s/g, '') + "'>Read Only</label>\
                                                    </div>\
                                                    <div class='radio radio-css radio-inline '>\
                                                        <input class='form-check-input clickable' type='radio' name='rdoReadAndWrite" + PageName.replace(/\s/g, '') + "' id='rdoReadWrite" + PageName.replace(/\s/g, '') + "' value='1' " + readAndWrite + ">\
                                                        <label class='form-check-label clickable' for='rdoReadWrite" + PageName.replace(/\s/g, '') + "'>Read & Write</label>\
                                                    </div>";
                            menuTab += "            <div class='checkbox checkbox-css checkbox-inline m-l-20'>\
                                                        <input type='checkbox'  id='chkDel" + PageName.replace(/\s/g, "") + "' value='" + this.ID + "' " + delChecked + " data-pagename='" + PageName.replace(/\s/g, "") + "' class='clickable'>\
                                                        <label class='clickable' for='chkDel" + PageName.replace(/\s/g, "") + "'>Delete</label>\
                                                    </div>";
                        }
                        menuTab += "        </li>";

                        if (sortedSubmenu.length) {
                            if (PageName == sortedSubmenu[0].ParentMenu) {
                                menuTab += "        <li>";
                                menuTab += "            <ul>";
                                $.each(sortedSubmenu, function (index) {
                                    var subChecked = this.Status != '' ? 'checked' : '';
                                    var readOnlySub = this.ReadAndWrite ? '' : 'checked';
                                    var delChecked = this.Delete ? 'checked' : '';
                                    var readAndWriteSub = this.ReadAndWrite ? 'checked' : '';
                                    menuTab += "            <li>";
                                    menuTab += "                <div class='row'>";
                                    menuTab += "                    <div class='col-md-2'>";
                                    menuTab += "                        <div class='checkbox checkbox-css checkbox-inline'>\
                                                                            <input type='checkbox'  id='chkSub" + this.PageName.replace(/\s/g, "") + "' value='" + this.ID + "' " + subChecked + " data-pagename='" + this.PageName.replace(/\s/g, "") + "' data-parent='chk" + this.ParentMenu.replace(/\s/g, "") + "' data-myclass='" + this.ParentMenu.replace(/\s/g, "") + "' class='subMenu chkUserAccess " + this.ParentMenu.replace(/\s/g, "") + " clickable'>\
                                                                            <label class='clickable' for='chkSub" + this.PageName.replace(/\s/g, "") + "'>" + this.PageLabel + "</label>\
                                                                        </div>";
                                    menuTab += "                    </div>";
                                    menuTab += "                    <div class='col-md-2'>";
                                    menuTab += "                        <div class='radio radio-css radio-inline'>\
                                                                            <input class='form-check-input clickable' type='radio' name='rdoReadAndWrite" + this.PageName.replace(/\s/g, '') + "' id='rdoRead" + this.PageName.replace(/\s/g, '') + "' value='0' " + readOnlySub + ">\
                                                                            <label class='form-check-label  clickable' for='rdoRead" + this.PageName.replace(/\s/g, '') + "'>Read Only</label>\
                                                                        </div>\
                                                                    </div>";
                                    menuTab += "                    <div class='col-md-2'>\
                                                                        <div class='radio radio-css radio-inline'>\
                                                                            <input class='form-check-input clickable' type='radio' name='rdoReadAndWrite" + this.PageName.replace(/\s/g, '') + "' id='rdoReadWrite" + this.PageName.replace(/\s/g, '') + "' value='1' " + readAndWriteSub + ">\
                                                                            <label class='form-check-label  clickable' for='rdoReadWrite" + this.PageName.replace(/\s/g, '') + "'>Read & Write</label>\
                                                                        </div>\
                                                                    </div>";
                                    menuTab += "                    <div class='col-md-2'>";
                                    menuTab += "                        <div class='checkbox checkbox-css checkbox-inline'>\
                                                                            <input type='checkbox'  id='chkDel" + this.PageName.replace(/\s/g, "") + "' value='" + this.ID + "' " + delChecked + " class='clickable'/>\
                                                                            <label class='clickable' for='chkDel" + this.PageName.replace(/\s/g, "") + "'>Delete</label>\
                                                                        </div>";
                                    menuTab += "                    </div>";
                                    menuTab += "                </div>";
                                    menuTab += "            </li>";
                                });
                                menuTab += "            </ul>";
                                menuTab += "        </li>";
                            }
                        }

                        menuTab += "    </ul>";
                        menuTab += "</div>";
                        //End Creating indivdial menu list content
                    }
                });

                menuTab += "</div>";
                //Creating Parent Menu Tab Content
                menuTab += "</div>";
                //End Creating GroupLabel Individual Tab Content
            });
            menuTab += "</div>";
            //End Creating GroupLabel Tab Content
            $("#mdlBodyUserAccess").html(menuTab);
            $("#mdlUserAccess").modal("show");
            return this;
        },
        saveUserAccess: function () {
            var self = this;
            var ID = self.$tblUser.rows({ selected: true }).data()[0].ID;
            var arrAccessList = [];
            $(".chkUserAccess").each(function (i, val) {
                var userAccessList = {};
                var status = "";
                var pagename = $(this).data("pagename");
                var readAndWrite = "0";
                var deleteFunction = "0";
                if ($(this).is(":checked")) {
                    status = true;
                } else {
                    status = false;
                }
                if ($("input[name=rdoReadAndWrite" + pagename + "]").length) {
                    readAndWrite = $("input[name=rdoReadAndWrite" + pagename + "]:checked").val();
                } else {
                    if ($(this).data('hassub') == '1')
                        readAndWrite = $(this).data('readandwrite');
                }
                if ($("#chkDel" + pagename).is(":checked")) {
                    deleteFunction = "1"
                } else {
                    if ($(this).data('hassub') == '1')
                        deleteFunction = $(this).data('delete');
                }

                userAccessList = {
                    UserID: ID,
                    PageID: $(this).val(),
                    Status: status,
                    ReadAndWrite: readAndWrite,
                    Delete: deleteFunction
                }
                arrAccessList.push(userAccessList);
            });

            var AdminLength = _.filter(arrAccessList, function (o) { if (o.Delete == '1' && o.ReadAndWrite == '1') return o }).length
            var EncoderLength = _.filter(arrAccessList, function (o) { if (o.Delete == '0' && o.ReadAndWrite == '1') return o }).length
            var ViewerLength = _.filter(arrAccessList, function (o) { if (o.Delete == '0' && o.ReadAndWrite == '0') return o }).length

            self.formAction = '/MasterMaintenance/UserMaster/SaveUserAccess';
            self.jsonData = { userAccessList: arrAccessList, UserID: ID };
            self.sendData().then(function (response) {
                self.$tblUser.ajax.reload(null, false);
                self.cancelUserTbl();
                $("#mdlUserAccess").modal("hide");
            });
            return this;
        }
    }
    UserMaster.init.prototype = $.extend(UserMaster.prototype, $D.init.prototype);
    UserMaster.init.prototype = UserMaster.prototype;

    $(document).ready(function () {
        var User = UserMaster();
        User.drawDatatables();
        $("#btnAddUser").click(function () {
            User.populateMode = false;
            $("#mdlUser").modal("show");
        });
        $("#UserName").change(function () {
            User.validateUserName(this);
        });
        $("#EmailAddress").change(function () {
            User.validateEmailAddress(this);
        });
        $("#Password").change(function () {
            if ($("#Password").attr("type") == "text")
                $(this).attr("data-didpasswordchange", true);
        });
        $('.btnCancelUser').click(function () {
            User.cancelUserForm();
        });
        $("#btnGeneratePassword").click(function () {
            $("#Password").attr("readonly", false).attr("type", "text").val(User.generatePassword()).trigger("change");
            User.parsleyValidate("frmUser");
        });
        $("#frmUser").submit(function (e) {
            e.preventDefault();
            User.saveUser();
        });
        $('#btnEditUser').click(function () {
            User.editUser();
        });
        $('#btnDeleteUser').click(function () {
            User.msg = "Are you sure you want to delete this User?";
            User.confirmAction().then(function (approve) {
                if (approve)
                    User.deleteUser();
            });
        });
        $('#btnUserAccess').click(function () {
            User.getUserAccess();
        });
        $("#mdlBodyUserAccess").on("click", ".parentMenu", function () {
            var childClass = $(this).data("childclass");
            $("." + childClass).prop("checked", $(this).is(":checked"));
        });
        $("#mdlBodyUserAccess").on("click", ".subMenu", function () {
            var parentID = $(this).data("parent");
            var myClass = $(this).data("myclass");
            $("#" + parentID).prop("checked", $("." + myClass + ":checked").length);
        });
        $("#btnSaveUserAccess").click(function () {
            User.saveUserAccess();
        });
        $("#tblUser").on("change", '.columnSearch', function () {
            User.$tblUser.ajax.reload(null, false);
        });
        $('#btnPrint').on('click', function (e) {
            User.$tblUser.ajax.reload(null, false);
            window.location = "/MasterMaintenance/UserMaster/DownloadUser";
        });

        //#Special Events
        User.$tblUser.on('draw.dt', function () {
            CUI.dataTableID = "#tblUser";
            CUI.setDatatableMaxHeightFixed();
        });
        User.$tblUser.on('select', function (e, dt, type, indexes) {
            var ID = User.$tblUser.rows({ selected: true }).data()[0].ID;
            $(".tbl-tr-btns").prop("disabled", false);
            if (ID == $("#img-user").attr("data-id")) {
                $("#btnDeleteUser").prop("disabled", true);
            }
        });
        User.$tblUser.on('deselect', function (e, dt, type, indexes) {
            $(".tbl-tr-btns").prop("disabled", true);
        });
        $('#mdlUser').on('shown.bs.modal', function (e) {
            $('#Section').select2({
                ajax: {
                    url: "/General/MySQLGetSelect2Data",
                    data: function (params) {
                        var search = params.term || "";
                        return {
                            q: params.term,
                            id: 'ID',
                            text: 'SectionName',
                            table: 'sectionmaster',
                            db: 'ERPReports',
                            display: 'id&text',
                            query: "SELECT ID, CONCAT(`Section`,'-',`Name`) as SectionName FROM sectionmaster WHERE IsDeleted=0 AND CONCAT(`Section`,'-',`Name`) LIKE '%" + search + "%';",
                        };
                    },
                },
                placeholder: '--Please Select--',
            });
            $('#Position').select2({
                ajax: {
                    url: "/General/MySQLGetSelect2Data",
                    data: function (params) {
                        var search = params.term || "";
                        return {
                            q: params.term,
                            id: 'ID',
                            text: 'PositionDescription',
                            table: 'sectionmaster',
                            db: 'ERPReports',
                            display: 'id&text',
                            query: "SELECT ID, CONCAT(`Position`,'-',`Description`) as PositionDescription FROM positionmaster WHERE IsDeleted=0 AND CONCAT(`Position`,'-',`Description`) LIKE '%" + search + "%';",
                        };
                    },
                },
                placeholder: '--Please Select--',
            });
        });
    });
})();
