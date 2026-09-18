var transaction_id;
var currentfile = [];
var sendMsg = "update_payment_";
var sendMsgQ = "update_queue_";
var reqChanel;
var roundedNumber;
var list_alien_id = [];
var objDataPayment;
var user_reject;
var form_reject;
var group_reject;
var paymentobj = "";
var checkage = true;
var statuscheck;
var form_type_nameTH;
var form_type_nameEN;
var formtypeID;
var recruitment_agency;
var bus_type_id;
var DataImmigrationCheckpoint;
var fsc_id;
var draftfile;
var name_address_work;
var jsonStr;
var NoHave_insurance = true;


const postRequest = (url, data = null) => {
    return new Promise((resolve, reject) => {
        $.ajax({
            type: "GET",
            url: url,
            data: data,
            dataType: "json",
            success: resolve,
            error: reject
        });
    });
};


async function GetDataNationalityInfo() {
    let objRequest = {
        is_mou: "Y",
    };
    var request = await jQuery.post(
        url_GetDataNationalityInfo,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;

                var setData = "";
                setData +=
                    "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    setData +=
                        "<option value='" +
                        objData[i].nationality_id +
                        "' data_th='" +
                        objData[i].nationality_th?.toSafeJsonHtml() +
                        "' data_en='" +
                        objData[i].nationality_en +
                        "'></option>";
                }
                jQuery("#nationality_id_mou").html(setData);
            }
        },
        "json"
    );
    await changeDatalang();
}
async function GetDataPermitCateWorkInfo() {
    let objRequest = {
        is_mou: "Y",
    };

    var request = await jQuery.post(
        url_GetDataPermitCateWorkInfo,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                var setData = "";

                for (var i = 0; i < objData.checklist.length; i++) {
                    if (
                        objData.checklist[i].active == "Y" ||
                        objData.checklist[i].active == "y"
                    ) {
                        let _permitValue = parseIfJson(
                            objData.checklist[i].permit_catework_name_th
                        );

                        if (typeof _permitValue == "object") {
                            _permitValue =
                                _permitValue["TH"] || _permitValue["Th"] || _permiValue["th"];
                        }

                        setData += '<div class="col-lg-12 mt-3">';
                        setData += '<div class="payment">';
                        setData += '<label class="cursor-pointer w-100 h-100">';
                        setData +=
                            '<input class="card-pay style-pay control__indicator d-none" type="radio" data-sale="' +
                            objData.checklist[i].is_saleperson +
                            '"  data-value="' +
                            _permitValue +
                            '" data_th="' +
                            objData.checklist[i].permit_catework_name_th?.toSafeJsonHtml() +
                            '" data_en="' +
                            objData.checklist[i].permit_catework_name_en +
                            '"  name="permit_type" id="checkboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '" value="' +
                            objData.checklist[i].permit_catework_id +
                            '" onclick="initializepermit()">';
                        setData += ' <div class="control__indicator radio1" ></div>';
                        setData +=
                            '<div class="border rounded pd border-gray-300 bg-white pd-l-6 select">';
                        setData += '<div class="row">';
                        setData +=
                            '<div class="col-lg-4 col-md-4 col-12 pr-md-4 pr-auto"><h4 class="head-card-select head-card-text" data_th=' +
                            objData.checklist[i].permit_catework_name_th?.toSafeJsonHtml() +
                            "  data_en=" +
                            objData.checklist[i].permit_catework_name_en +
                            "></h4></div>";
                        setData +=
                            ' <div class="col-lg-4 col-md-4 col-12 pr-md-4 pr-auto"><div class="input-group form-group"><div class="input-group-append"><span class="input-group-text2 span-job lang_lbl_male1">เพศชาย</span></div>';

                        setData +=
                            '<input type="text" class="form-control right-text-job check_number_lang_specify_quantity' +
                            objData.checklist[i].permit_catework_id +
                            '" name="check_number_checkboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '" placeholder="ระบุจำนวน" onchange="initializepermit()" oninput="this.value = this.value.replace(/[^0-9]/g,\'\'); if(this.value.startsWith(\'0\')) this.value = \'\';" id="numbermen_checkboxjob_' + objData.checklist[i].permit_catework_id +
                            '" style="border-left:0;text-align: right;" disabled></div>';
                        setData +=
                            '<label id="numbermen_checkboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '-error" class="error" for="numbermen_checkboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '"></label></div>';

                        setData +=
                            '<div class="col-lg-4 col-md-4 col-12 pr-md-4 pr-auto"><div class="input-group form-group"><div class="input-group-append"><span class="input-group-text2 span-job lang_lbl_female">เพศหญิง</span></div>';

                        setData +=
                            '<input type="text" class="form-control right-text-job check_numberWomen_ lang_specify_quantity' +
                            objData.checklist[i].permit_catework_id +
                            '" name="check_numberWomen_checkboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '" placeholder="ระบุจำนวน" onchange="initializepermit()" oninput="this.value = this.value.replace(/[^0-9]/g,\'\'); if(this.value.startsWith(\'0\')) this.value = \'\';"  id="numberwomen_checkboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '" style="border-left:0;text-align: right;" disabled></div>';
                        setData +=
                            '<label id="numberwomen_checkboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '-error" class="error" for="numberwomen_checkboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '"></label>';
                        setData +=
                            '<span id="validateinputnumbercheckboxjob_' +
                            objData.checklist[i].permit_catework_id +
                            '" class="error" style="display:none; text-align:end;color: #E10000 !important;font-weight: 500;font-size: 12px;">กรุณากรอกข้อมูล</span>';
                        setData += "</div></div></div></div></label></div></div>";
                    }
                }
                jQuery("#permit_catework_id").html(setData);
                var setDataDropdown = "";
                setDataDropdown +=
                    '<option class="lang_opt_pleaseselect" selected="" value="">กรุณาเลือก</option>';
                for (var o = 0; o < objData.otherlist.length; o++) {
                    if (
                        objData.otherlist[o].active == "Y" ||
                        objData.otherlist[o].active == "y"
                    ) {
                        setDataDropdown +=
                            '<option value="' +
                            objData.otherlist[o].permit_catework_id +
                            '" data-value="' +
                            objData.otherlist[o].is_saleperson +
                            '" data_th=' +
                            objData.otherlist[o].permit_catework_name_th?.toSafeJsonHtml() +
                            "  data_en=" +
                            objData.otherlist[o].permit_catework_name_en +
                            ">" +
                            objData.otherlist[o].permit_catework_name_th +
                            "</option>";
                    }
                }
                jQuery("#permit_catework_Dropdown").html(setDataDropdown);

            }
        },
        "json"
    );
}
async function GetDataDemandLetter() {
    try {

        const response = await postRequest(url_GetDataDemandLetter);

        if (response.status === "success") {
            const data = response.data[0];

            formtypeID = data.form_type_id;
            form_type_nameTH = data.form_mt_th;
            form_type_nameEN = data.form_mt_en;

            $("#text_name_section, #text_header").attr({
                "data_th": data.form_type_name_th,
                "data_en": data.form_type_name_en
            });

            $("#card_title_sub").attr({
                "data_th": data.form_desc_th,
                "data_en": data.form_desc_en
            });

            populateDropdowns(
                data.req_permit_max_year,
                data.req_permit_max_month,
                data.req_permit_max_day
            );
        }
    } catch (err) {
        console.error("Error fetching Demand Letter data:", err);
    }
};


async function GetDataImmigrationCheckpoint() {
    $("#DataImmigrationCheckpoint").removeAttr("disabled");
    $("#fsc_id").removeAttr("disabled");

    let objRequest = {
        nationality_id: $("#nationality_id_mou :selected").val(),
    };

    var request = await jQuery.post(
        url_GetDataImmigrationCheckpoint,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;

                var setData = "";
                setData +=
                    "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    if (objData[i].active == "Y" || objData[i].active == "y") {
                        setData +=
                            "<option value='" +
                            objData[i].immigration_checkpoint_id +
                            "' data_th='" +
                            objData[i].immigration_checkpoint_name_th?.toSafeJsonHtml() +
                            "' data_en='" +
                            objData[i].immigration_checkpoint_name_en +
                            "'></option>";
                    }
                }

                jQuery("#DataImmigrationCheckpoint").html(setData);
                if (DataImmigrationCheckpoint) {
                    $("#DataImmigrationCheckpoint").val(DataImmigrationCheckpoint);
                }

                GetDataFirstServiceCenter("", $("#nationality_id_mou :selected").val());
                GetDataRecruitmentAgency($("#nationality_id_mou :selected").val());
                $("#DataImmigrationCheckpoint").on("change", function () {
                    GetDataFirstServiceCenter(
                        this.value,
                        $("#nationality_id_mou :selected").val()
                    );
                });
            }
        },
        "json"
    );
    await changeDatalang();
}

async function GetDataFirstServiceCenter(value_select, value_nationality) {
    let objRequest = {
        nationality_id: value_nationality,
    };

    var request = await jQuery.post(
        url_GetDataFirstServiceCenter,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;

                var setData = "";
                setData +=
                    "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    if (objData[i].active == "Y" || objData[i].active == "y") {
                        setData +=
                            "<option value='" +
                            objData[i].fsc_id +
                            "' data_th='" +
                            objData[i].fsc_name_th?.toSafeJsonHtml() +
                            "' data_en='" +
                            objData[i].fsc_name_en +
                            "'></option>";
                    }
                }

                jQuery("#fsc_id").html(setData);
                if (value_select != "") {
                    $("#fsc_id").val(value_select);
                }
                if (fsc_id) {
                    $("#fsc_id").val(fsc_id);
                }
            }
        },
        "json"
    );
    await changeDatalang();
}
async function GetDataRecruitmentAgency(value_nationality) {
    $("#recruitment_agency").removeAttr("disabled");
    let objRequest = {
        nationality_id: value_nationality,
    };

    var request = await jQuery.post(
        url_GetDataRecruitmentAgency,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                var setData = "";
                setData +=
                    "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    if (objData[i].active == "Y" || objData[i].active == "y") {
                        setData +=
                            "<option data-value='" +
                            objData[i].recruitment_agency_address +
                            "' value='" +
                            objData[i].recruitment_agency_id +
                            "' data_th='" +
                            objData[i].recruitment_agency_name_th +
                            "' data_en='" +
                            objData[i].recruitment_agency_name_en +
                            "'>" +
                            objData[i].recruitment_agency_name_th +
                            "</option>";
                    }
                }

                jQuery("#recruitment_agency").html(setData);

                if (recruitment_agency) {
                    $("#recruitment_agency").val(recruitment_agency);
                }
            }
        },
        "json"
    );
    await changeDatalang();
}
function getaddress_recruitment_agency() {
    var selectedData = $("#recruitment_agency :selected").data("value");
    $("#address_recruitment").val(selectedData);
}
function get_select_addwork() {
    var selectedData = $("#name_address_work :selected").data("value");

    var no = selectedData.no;
    var village_building = selectedData.village_building;
    var soi = selectedData.soi;
    var road = selectedData.road;
    var sub_dist_nmae = selectedData.sub_dist_nmae;
    var dist_name = selectedData.dist_name;
    var province_name = selectedData.province_name;

    if (village_building == "" || village_building == "-") {
        village_building = "";
    } else {
        village_building = "หมู่ " + village_building;
    }

    if (soi == "" || soi == "-") {
        soi = "";
    } else {
        soi = "ซอย " + soi;
    }

    if (road == "" || road == "-") {
        road = "";
    } else {
        road = "ถนน " + road;
    }
    $("#address_work").val(
        no +
        " " +
        village_building +
        " " +
        soi +
        " " +
        road +
        " " +
        sub_dist_nmae +
        " " +
        dist_name +
        " " +
        province_name
    );

    $("#workplace_no").val(selectedData.no);
    $("#workplace_village_building").val(selectedData.village_building);
    $("#workplace_soi").val(selectedData.soi);
    $("#workplace_road").val(selectedData.road);
    $("#workplace_province").val(selectedData.province);
    $("#workplace_district").val(selectedData.dist);
    $("#workplace_sub_district").val(selectedData.sub_dist);
    $("#workplace_postal_code").val(selectedData.postal_code);
    var data_bus = $("#name_address_work :selected").attr("data-value-add");
    var obj_bus = JSON.parse(data_bus);
    var setdatabus = "";


    for (var i = 0; i < obj_bus.length; i++) {
        setdatabus += "<option value='" + obj_bus[i].bus_type_id + "' data-sale='" + obj_bus[i].is_saleperson + "' data_th='" + obj_bus[i].bus_type_name_th?.toSafeJsonHtml() + "' data_en='" + obj_bus[i].bus_type_name_en + "'>" + obj_bus[i].bus_type_name_th + "</option>";
    }
    $("#bus_type_id").html(setdatabus);
    if (bus_type_id) {
        $("#bus_type_id").val(bus_type_id);
    }
    console.log(bus_type_id)

    changeDatalang();
}
async function GetDataAcceptWorkInfo() {
    var request = await jQuery.post(
        url_GetDataAcceptWorkInfo,
        null,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                var setData = "";

                RejectAcceptWork = $(`#RejectAcceptWork`).DataTable({
                    paging: true,
                    lengthChange: false,
                    searching: false,
                    ordering: true,
                    destroy: true,
                    columns: [
                        { data: "accept_work_id" },
                        { data: "accept_work_id" },
                        { data: "accept_work_name_th", className: "text-left" },
                    ],
                    columnDefs: [
                        {
                            targets: 0,
                            width: "5%",
                            render: function (data, type, row, meta) {
                                return `<div class="checkbox mb-0"> <label class="mb-0"><input type="checkbox" name="acceptwork" value="${row.accept_work_id}" id="${row.accept_work_id}" onclick="checkselectacceptwork(${row.accept_work_id})"><span class="custom-checkbox"><i class="icon-check"></i></span></label> </div>`;
                            },
                        },

                        {
                            targets: 1,
                            render: function (data, type, row, meta) {
                                return `${meta.row + meta.settings._iDisplayStart + 1}`;
                            },
                        },
                        {
                            targets: 2,
                            render: function (data, type, row, meta) {
                                if (row.accept_work_id == "0") {
                                    return `<span data_th=' ${row.accept_work_name_th?.toSafeJsonHtml()}' data_en=' ${row.accept_work_name_en
                                        }'></span><textarea type="text" class="form-control mt-2" name="employerID_59_check" id="resson_acceptwork" placeholder="กรุณากรอก" disabled></textarea><label id="resson_acceptwork-error" class="error d-none lang_please_input_info" for="resson_acceptwork" >กรุณากรอกข้อมูล</label>`;
                                } else {
                                    return `<span data_th=' ${row.accept_work_name_th?.toSafeJsonHtml()}' data_en=' ${row.accept_work_name_en
                                        }'></span>`;
                                }
                            },
                        },
                    ],

                    data: objData,
                    drawCallback: function (settings) {

                        changeDatalang();
                    },
                });
            }
            $("#loading").delay().fadeOut("");
        },
        "json"
    );
}

async function buttonclickacceptwork() {
    var checkdataacceptwork = true;
    var arrayaccept = [];
    var arrayorther = [];
    $('input[name="acceptwork"]:checked').each(function () {

        var ID_checked = this.id;
        arrayaccept.push(ID_checked);

        if (ID_checked == "0") {
            if ($("#resson_acceptwork").val() == "") {
                checkdataacceptwork = false;
                $("#resson_acceptwork-error").removeClass("d-none");
            } else {
                arrayorther.push($("#resson_acceptwork").val());
                $("#resson_acceptwork-error").addClass("d-none");
            }
        } else {
            arrayorther.push("0");
        }
    });
    if (checkdataacceptwork) {
        const isTokenValid = await validate_token_presubmit();
        if (!isTokenValid) {
            return false;
        }
        let objRequest = {
            user_id: user_reject,
            group_id: group_reject,
            form_id: form_reject,
            accept_work_id: arrayaccept.toString(),
            accept_work_other: arrayorther.toString(),
            relate_account_id: core_datas.current_profile.profile_id,
        };
        var request = jQuery.post(
            url_UpdateRejectAcceptWork,
            objRequest,
            function (data) {
                if (data.status == "success") {
                    Swal.fire({
                        title: '<span class="lang_reject_work">ปฎิเสธการจ้างงาน</span>',
                        html: '<span class="lang_Employment_rejection_completed_successfully">ปฏิเสธการจ้างงานสำเร็จ</span>',
                        imageUrl: important_data.successIconUrl,
                        showCancelButton: false,
                        confirmButtonText: '<span data_th="ตกลง" data_en="OK"></span>',
                        reverseButtons: true,
                    }).then((result) => {
                        location.reload();

                    });
                    changeDatalang();
                } else {
                    Swal.fire({
                        title: '<span class="lang_reject_work">ปฎิเสธการจ้างงาน</span>',
                        html: '<span class="lang_Failed_reject_employment">ปฏิเสธการจ้างงานไม่สำเร็จ</span>',
                        imageUrl: important_data.errorIconUrl,
                        showCancelButton: false,
                        confirmButtonText: '<span data_th="ตกลง" data_en="OK"></span>',
                        reverseButtons: true,
                    }).then((result) => {
                        location.reload();

                    });
                    changeDatalang();
                }
            },
            "json"
        );
    }
}
function checkselectacceptwork() {
    const checkboxes = document.querySelectorAll('input[name="acceptwork"]');
    const buttonselectacceptwork = document.getElementById(
        "buttonclickacceptwork"
    );
    const ressonAcceptwork = $("#resson_acceptwork");


    let isAnyCheckboxChecked = Array.from(checkboxes).some(
        (checkbox) => checkbox.checked
    );


    let isValueZeroChecked = Array.from(checkboxes).some(
        (checkbox) => checkbox.checked && checkbox.value === "0"
    );

    if (checkboxes.length > 0) {
        if (isAnyCheckboxChecked) {

            buttonselectacceptwork.disabled = false;

            if (isValueZeroChecked) {

                ressonAcceptwork.attr("disabled", false);
            } else {

                ressonAcceptwork.attr("disabled", true);
            }
        } else {

            buttonselectacceptwork.disabled = true;
            ressonAcceptwork.attr("disabled", true);
        }
    } else {

        buttonselectacceptwork.disabled = true;
        ressonAcceptwork.attr("disabled", true);
    }
}
async function GetDataPrefixInfo() {
    var request = await jQuery.post(
        url_GetDataPrefixInfo,
        null,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                var setData = "";
                setData +=
                    "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    setData +=
                        "<option value='" +
                        objData[i].pfid +
                        "'>" +
                        objData[i].prefix_name_en +
                        "</option>";
                }
                jQuery("#prefixNameEnDropdown").html(setData);
            }
            $("#loading").delay().fadeOut("");
        },
        "json"
    );
}
async function GetDataCountryInfo() {

    var request = await jQuery.post(
        url_GetDataCountryInfo,
        null,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                var setData = "";
                setData +=
                    "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    if (objData[i].active == "Y" || objData[i].active == "y") {
                        setData +=
                            "<option value='" +
                            objData[i].country_id +
                            "' data_th='" +
                            objData[i].country_name_th?.toSafeJsonHtml() +
                            "' data_en='" +
                            objData[i].country_name_en +
                            "'></option>";
                    }
                }
                jQuery("#country_al").html(setData);

            }
            $("#loading").delay().fadeOut("");
        },
        "json"
    );
    await changeDatalang();
}
async function GetDataStayPermissionTypeInfo() {
    var request = await jQuery.post(
        url_GetDataStayPermissionTypeFormInfo,
        null,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                var setData = "";

                for (var i = 0; i < objData.length; i++) {
                    setData += '<div class="col-md-4">';
                    setData +=
                        '<label class="container_2" style="color: #11111F !important; font-weight: 500;"><span data_th=' +
                        objData[i].stay_permis_type_name_th?.toSafeJsonHtml() +
                        " data_en=" +
                        objData[i].stay_permis_type_name_en +
                        "></span>";
                    setData +=
                        '<input type="radio" data-value="' +
                        objData[i].stay_permis_type_name_th +
                        '" name="stay_permission_type46_59" value=' +
                        objData[i].stay_permis_type_name_th +
                        " id=" +
                        objData[i].stay_permis_type_id +
                        ">";
                    setData += '<span class="checkmark"></span>';
                    setData += "</label>";
                    setData += "</div>";
                }

                jQuery("#stay_permission_type46_59").html(setData);

                var radios = document.getElementsByName("stay_permission_type46_59");


                if (radios.length > 0) {
                    radios[0].checked = true;
                }
            }
        },
        "json"
    );
    await changeDatalang();
}

function initializepermit() {
    $('input[name="permit_type"]').each(function () {
        var ID_checked = this.id;

        if ($("#" + ID_checked).is(":checked")) {
            $("#numbermen_" + ID_checked).attr("disabled", false);
            $("#numberwomen_" + ID_checked).attr("disabled", false);
            $("#number_of_alien_male").val($("#numbermen_" + ID_checked).val());
            $("#number_of_alien_female").val($("#numberwomen_" + ID_checked).val());
            var women = $("#numberwomen_" + ID_checked).val();
            var men = $("#numbermen_" + ID_checked).val();
            if (women != "" && men != "") {
                var sum_women = parseInt($("#numberwomen_" + ID_checked).val());
                var sum_men = parseInt($("#numbermen_" + ID_checked).val());
                $("#number_of_alien_all").val(sum_women + sum_men);
            } else if (women == "" && men != "") {
                var sum_men = parseInt($("#numbermen_" + ID_checked).val());
                $("#number_of_alien_all").val(sum_men);
            } else if (women != "" && men == "") {
                var sum_women = parseInt($("#numberwomen_" + ID_checked).val());
                $("#number_of_alien_all").val(sum_women);
            } else {
                $("#number_of_alien_all").val("");
            }
            if ($("#checkboxjob_4").is(":checked")) {
            } else {
                $("#permit_catework_Dropdown").val("").trigger("change");
            }
        } else {
            $("#numbermen_" + ID_checked).attr("disabled", true);
            $("#numberwomen_" + ID_checked).attr("disabled", true);
            $("#numbermen_" + ID_checked).val("");
            $("#numberwomen_" + ID_checked).val("");
        }

        if ($("#checkboxjob_4").is(":checked")) {


            $(".checkboxjob_4").css({
                border: "1px solid #E24C86",
                background: "#fcf8ff",
            });

            $("#permit_catework_Dropdown").attr("disabled", false);
            $("#numbermen_checkboxjob_4").attr("disabled", false);
            $("#numberwomen_checkboxjob_4").attr("disabled", false);
            $("#number_of_alien_male").val($("#numbermen_checkboxjob_4").val());
            $("#number_of_alien_female").val($("#numberwomen_checkboxjob_4").val());
            var women = $("#numberwomen_checkboxjob_4").val();
            var men = $("#numbermen_checkboxjob_4").val();
            if (women != "" && men != "") {
                var sum_women = parseInt($("#numberwomen_checkboxjob_4").val());
                var sum_men = parseInt($("#numbermen_checkboxjob_4").val());
                $("#number_of_alien_all").val(sum_women + sum_men);
            } else if (women == "" && men != "") {
                var sum_men = parseInt($("#numbermen_checkboxjob_4").val());
                $("#number_of_alien_all").val(sum_men);
            } else if (women != "" && men == "") {
                var sum_women = parseInt($("#numberwomen_checkboxjob_4").val());
                $("#number_of_alien_all").val(sum_women);
            } else {
                $("#number_of_alien_all").val("");
            }
        } else {
            $(".checkboxjob_4").addClass("bg-white");
            $(".checkboxjob_4").css("border", "1px solid #eaeaed");

            $("#permit_catework_Dropdown").attr("disabled", true);
            $("#numbermen_checkboxjob_4").attr("disabled", true);
            $("#numberwomen_checkboxjob_4").attr("disabled", true);
            $("#numbermen_checkboxjob_4").val("");
            $("#numberwomen_checkboxjob_4").val("");


        }
    });
}

function initializePageCheckbox() {
    var checkbox = document.getElementById("check_truth_46_59");
    var button = document.getElementById("NextStepFour46_59");

    if (checkbox.checked) {
        button.disabled = false;
    } else {
        button.disabled = true;
    }
}
$("#is_accept_work_Y").on("click", function () {
    if ($("#is_accept_work_Y").is(":checked") == true) {
        $(".work_info").hide();
        $("#uploadfilealien").show();
        $(".work_come_in").show();
        $(".work_info_detail").hide();
        $("#bus_type_id").val($("#hidden_bus_type_text").val());
        $("#bus_type_id_value").val($("#hidden_bus_type_id").val());
    }
});
$("#is_accept_work_N").on("click", function () {
    if ($("#is_accept_work_N").is(":checked") == true) {
        $(".work_info").show();
        $("#uploadfilealien").hide();
        $(".work_come_in").hide();
    }
});
$("#checkAll").on("click", function () {
    if ($("#checkAll").is(":checked")) {
        let isChecked = $(this).is(":checked");


        $("#AppointmentMoreDocForSend")
            .DataTable()
            .rows()
            .every(function () {
                let row = this.node();
                $(row).find("[name='form_sendDoc']").prop("checked", isChecked);
            });


        $("#confirm_send_btn").attr("disabled", false);
    } else {

        $("#AppointmentMoreDocForSend")
            .DataTable()
            .rows()
            .every(function () {
                let row = this.node();
                $(row).find("[name='form_sendDoc']").prop("checked", false);
            });

        $("#confirm_send_btn").attr("disabled", true);
    }
});
function GetWorkID(value) {
    if (value == "0") {
        $(".work_info_detail").toggle();
    }
}
function selectdate_age() {
    const minDate = new Date();
    minDate.setFullYear(minDate.getFullYear() - 100);
    const minDateStr = minDate.toISOString().split("T")[0];
    const today = new Date();
    const todayStr = today.toISOString().split("T")[0];

    if ($("#birthdate_al").val() != "") {
        var birthdate_al = $("#birthdate_al").val().split("/");
        const birthdateInput =
            birthdate_al[2] + "-" + birthdate_al[1] + "-" + birthdate_al[0];
        calculateAge(birthdateInput);
    }
}
function calculateAge(birthdateInput) {
    const ageInput = document.getElementById("age_al");

    if (birthdateInput) {
        const birthdate = new Date(birthdateInput);
        const today = new Date();

        const age =
            today.getFullYear() -
            birthdate.getFullYear() -
            (today.getMonth() < birthdate.getMonth() ||
                (today.getMonth() === birthdate.getMonth() &&
                    today.getDate() < birthdate.getDate()));

        const intAge = parseInt(age);


        const minAge = Number(important_data.alienAge.minAge) || 18;
        const maxAge = Number(important_data.alienAge.maxAge) || 55;

        if (intAge >= minAge && intAge <= maxAge) {
            $("#age_al-error").hide();
            ageInput.value = intAge;
            checkage = true;
        } else {
            checkage = false;
            ageInput.value = "";
            $("#age_al-error")
                .show()
                .html(`<span>อายุระหว่าง ${minAge} ถึง ${maxAge} ปี</span>`);
        }
    } else {
        ageInput.value = "";
    }
}


function ConvertDate(dateValue) {
    var sp_dateValue = dateValue.split("/");
    var resultDate =
        sp_dateValue[2] + "-" + sp_dateValue[1] + "-" + sp_dateValue[0];
    return resultDate;
}

$("#NextStepTwoPageOneMOU").on("click", function (event) {
    event.preventDefault();
    validateForm();
});
function validateForm() {
    let isAnyCheckboxChecked = false;
    let isAnyInputEntered = false;
    let check = true;
    let arrayMen = [];
    let arraySave_Men = [];
    let arrayWomen = [];
    let arraySave_Women = [];
    let arrayData = [];

    $('input[name="permit_type"]:checked').each(function () {
        isAnyCheckboxChecked = true;
        var ID_checked = this.id;
        var menInput = $("#numbermen_" + ID_checked);
        var womenInput = $("#numberwomen_" + ID_checked);

        if (menInput.val() === "" && womenInput.val() === "") {
            $("#validateinputnumber" + ID_checked).css("display", "block");
            check = false;
        } else if (menInput.val() === "" && womenInput.val() === "0") {
            $("#validateinputnumber" + ID_checked).css("display", "block");
            check = false;
        } else if (menInput.val() === "0" && womenInput.val() === "") {
            $("#validateinputnumber" + ID_checked).css("display", "block");
            check = false;
        } else if (menInput.val() === "0" && womenInput.val() === "0") {
            $("#validateinputnumber" + ID_checked).css("display", "block");
            check = false;
        } else {
            $("#validateinputnumber" + ID_checked).css("display", "none");
            isAnyInputEntered = true;
        }

        if ($("#checkboxjob_4").is(":checked")) {
            if ($("#permit_catework_Dropdown").val() == "") {
                $("#validateinputnumber" + ID_checked).css("display", "block");
                check = false;
            }
        }

        var checksale = $("#" + ID_checked).attr("data-sale");
        var bus_type_id = $("#bus_type_id :selected").attr("data-sale");

        if (bus_type_id === "Y" && checksale === "Y") {
            $("#issale").val(bus_type_id);
            $(".div_head_sell").css("display", "block");
            $(".listuploadfile_sell").css("display", "contents");
            OnloadFileSell();
        } else {
            $("#issale").val("N");
            $(".div_head_sell").css("display", "none");
            $(".listuploadfile_sell").css("display", "none");
        }

        if (menInput.val() === "") {
            arrayMen.push("-");
            arraySave_Men.push("0");
        } else {
            arrayMen.push(menInput.val());
            arraySave_Men.push(menInput.val());
        }

        if (womenInput.val() === "") {
            arrayWomen.push("-");
            arraySave_Women.push("0");
        } else {
            arrayWomen.push(womenInput.val());
            arraySave_Women.push(womenInput.val());
        }

        var dataValue = $(this).attr("data-value");
        arrayData.push(dataValue);
    });

    $("#All_Men").val(arraySave_Men);
    $("#All_Women").val(arraySave_Women);

    if (!isAnyCheckboxChecked) {
        $("#typejob_error").show();
        check = false;
    } else {
        $("#typejob_error").hide();
    }

    if (!isAnyInputEntered) {
        const focusedElement = document.getElementById("head_type_job");
        if (focusedElement) {
            focusedElement.focus();
            focusedElement.scrollIntoView({
                behavior: "smooth",
                block: "center",
                inline: "center",
            });
        }
        check = false;
    }

    if ($("#msform").valid()) {
        if (check) {
            $("#doc").addClass("active");
            $("#stepOnePageTwo").css("display", "none");
            $("#stepTwoPageOne").css("display", "block");
            $("#stepTwoPageOne").css("opacity", "1");
            showdetail();
        }
    } else {
        const inputElements = document.querySelectorAll(
            "select.error, input.error"
        );
        if (inputElements.length > 0) {
            const focusedElement = document.getElementById(inputElements[0].id);
            if (focusedElement) {
                focusedElement.focus();
                focusedElement.scrollIntoView({
                    behavior: "smooth",
                    block: "center",
                    inline: "center",
                });
            }
        }
    }
}


async function GetDataAddNewRequestFormSuccessMOU(demand_letter_id) {
    var splitdemand = demand_letter_id.split("|");
    let objRequest = {
        group_id: splitdemand[0],
        user_id: $("#user_id").val(),
    };

    var request = await jQuery.post(
        url_GetDataAddNewRequestFormDLSuccess,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                $("#show_createdtimestamp").attr(
                    "data_th",
                    convert_to_thai_date(objData.created_timestamp)
                );
                $("#show_createdtimestamp").attr(
                    "data_en",
                    convert_to_eng_date(objData.created_timestamp)
                );
                $("#show_number_of_alien_all").html(
                    objData.number_of_alien_all +
                    " " +
                    "<span class='lang_person'></span>"
                );
                $("#show_form_type_name_th").attr("data_th", objData.form_type_name_th);
                $("#show_form_type_name_th").attr("data_en", objData.form_type_name_en);
                $("#show_demand_letter_id").text(objData.group_id);

                if (objData.companyname_th == "") {
                    $("#show_employer_full_name").attr("data_th", objData.fullname_th);
                    $("#show_employer_full_name").attr("data_en", objData.fullname_en);
                } else {
                    $("#show_employer_full_name").attr("data_th", objData.companyname_th);
                    $("#show_employer_full_name").attr("data_en", objData.companyname_en);
                }

                $("#time_ap").html(
                    objData.work_duration_year +
                    " " +
                    "<span class='lang_lbl_year'><span>" +
                    " " +
                    objData.work_duration_month +
                    " " +
                    "<span class='lang_lbl_month'><span>" +
                    " " +
                    objData.work_duration_day +
                    " " +
                    "<span class='lang_lbl_day'><span>"
                );

                if (formtypeID == checkmt) {
                    deleteDraftFromIndexedDB();
                }
                if ($("#successed").hasClass("active")) {
                    $("#div_draft").css("display", "none");
                }
            }
        },
        "json"
    );
    await changeDatalang();
}
function openPayment() {
    $("#loading").css("display", "block");
    var status = sessionStorage.getItem("group_status");
    if (
        status == "NCOS2" ||
        status == "NCOC2" ||
        status == "NCOCR2" ||
        status == "NCOCRD2"
    ) {
        $("#loading").delay().fadeOut("");
        "#tabShowTableEmployee".css("display", "none");
        $("#tabShowPayment").css("display", "none");
        $("#tabShowPaymentSuccess").css("display", "block");
        $("#Alien_Table").hide();
        $("#divpayment").hide();
    } else {
        CheckDataNameListQuotaComplete();
        CalculatePaymentBeforeRequestSubmitMOU(
            url_CalculateDataPayment,
            roundedNumber,
            list_alien_id
        );
    }
}


async function GetDataFormAppointmentByGroupID() {
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
        };
    }
    var request = await jQuery.post(
        url_GetDataFormAppointmentByGroupID,
        objRequest,
        function (data) {
            var objData = data.data;

            var setData = "";
            if (data.status == "success") {
                for (var i = 0; i < objData.length; i++) {
                    var appointment_datetime = objData[i].appointment_datetime.split("-");
                    $("#date_ap").text(
                        appointment_datetime[2] +
                        " " +
                        MonthStringSub(appointment_datetime[1]) +
                        " " +
                        appointment_datetime[0]
                    );
                    $("#location_ap").attr("data_th", objData[i].locatename_th);
                    $("#location_ap").attr("data_en", objData[i].locatename_en);
                    setData += '<div class="col-12 mt-4">';
                    setData += '<div class="payment">';
                    setData += '<label class="cursor-pointer w-100">';
                    setData += '<div class="border rounded p-4 bg-white shadow-primary">';
                    setData += '<div class="row">';
                    setData += '<div class="col-lg-9 col-md-8 col-12 pr-md-4 pr-auto">';
                    setData +=
                        '<h4 class="mb-3" style="font-size: 18px; font-weight: bold;">' +
                        appointment_datetime[2] +
                        " " +
                        MonthStringSub(appointment_datetime[1]) +
                        " " +
                        appointment_datetime[0] +
                        "</h4>";
                    setData += '<h5 class="text-left">';
                    setData +=
                        '<span style="color: #6F6F85;">จำนวน :</span> ' +
                        objData[i].num_alien_appointment +
                        ' <span class="lang_person"></span> &nbsp;';
                    setData +=
                        '<span style="color: #EAEAED;">| &nbsp; </span> <span class="lang_location" style="color: #6F6F85;">สถานที่</span> <span>:</span> ' +
                        objData[i].locatename_th +
                        " &nbsp;";
                    setData += "</h5>";
                    setData += "</div>";
                    setData +=
                        '<div class="col-lg-3 col-md-4 col-12 d-flex justify-content-lg-end justify-content-sm-start justify-content-center align-items-center pt-md-0 pt-3 column-gap-20">';
                    setData +=
                        '<a href="javascript:void(0)" class="add-btn lang_watch_detail" value="" onclick="GetDataFormAppointmentByAppointmentDate(&quot;' +
                        objData[i].appointment_datetime +
                        '&quot;)">ดูรายละเอียด</a>';
                    setData += "</div>";
                    setData += "</div>";
                    setData += "</div>";
                    setData += "</label>";
                    setData += "</div>";
                    setData += "</div>";
                }
                $("#alien_appointment").html(setData);

                var urlParams = new URLSearchParams(window.location.search);
                if (urlParams.has("status") && urlParams.get("status") === "AP") {
                    urlParams.set("status", "APSS");
                    window.location.href =
                        window.location.pathname + "?" + urlParams.toString();
                    $("#ModalConfirmAp").modal("hide");

                    $("#div_appointment1").css("display", "none");
                    $("#div_appointment2").css("display", "none");
                    $("#div_appointment").css("display", "none");
                    $("#div_appointment0").css("display", "none");
                    $("#div_appointment4").css("display", "block");
                    $("#div_tab1").removeClass("active");
                    $("#tab_default_1").removeClass("active");

                    $("#div_tab5").addClass("active");
                    $("#tab_default_5").addClass("active");
                }
            }
        },
        "json"
    );
    await changeDatalang();
}
async function GetDataFormAppointmentByAppointmentDate(appointment_datetime) {
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            appointment_datetime: appointment_datetime,
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            appointment_datetime: appointment_datetime,
        };
    }
    var request = await jQuery.post(
        url_GetDataFormAppointmentByAppointmentDate,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;

                $("#text_num_ap5").text(objData.length);
                $("#text_num_ap5_2").text(objData.length);
                $("#div_appointment4").addClass("d-none");
                $("#div_appointment5").removeClass("d-none");
                datatable_Appointment_Detail = $(
                    `#datatable_Appointment_Detail`
                ).DataTable({
                    paging: true,
                    lengthChange: false,
                    searching: false,
                    ordering: true,
                    destroy: true,
                    columns: [
                        { data: "form_id", className: "text-left" },
                        { data: "alien_name_th", className: "text-left" },
                        { data: "alien_nationality_th", className: "text-left" },
                        { data: "alien_date_of_birth", className: "text-left" },
                        { data: "stay_permis_no", className: "text-left" },
                    ],
                    columnDefs: [
                        {
                            targets: 1,
                            render: function (data, type, row, meta) {
                                return `${row.alien_name_en}`;
                            },
                        },
                        {
                            targets: 2,
                            render: function (data, type, row, meta) {
                                return `<span data_th='${row.alien_nationality_th?.toSafeJsonHtml()}' data_en='${row.alien_nationality_en
                                    }'></span>`;
                            },
                        },
                    ],

                    data: objData,
                    drawCallback: function (settings) {

                        changeDatalang();
                    },
                });
                $("#loading").delay().fadeOut("");
            }
        },
        "json"
    );
    await changeDatalang();
}


async function GetDataRequestFormNameListByDemandLetterID(textcheck) {
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            relate_account_id: core_datas.current_profile.profile_id,
        };
    } else {
        var userid = core_datas.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            relate_account_id: core_datas.current_profile.profile_id,
        };
    }
    var request = await jQuery.post(
        url_GetDataRequestFormEtrackingMOU_ByGroupID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;

                $("#numberform").text(objData[0].formlist.length);
                $("#numberpayment").text(objData[0].formlist.length);
                $("#numberpayment2").text(objData[0].formlist.length);
                $("#alien_show_num").text(objData[0].formlist.length);
                var status = objData[0].group_status_id;

                if (objData[0].formlist.length != 0) {
                    if (
                        status == "WCOSNA2" ||
                        status == "WCOSAS2" ||
                        status == "WCOCNA2" ||
                        status == "WCOCAS2" ||
                        status == "WCOCOC2" ||
                        status == "AP" ||
                        status == "SS" ||
                        status == "WCOCOCNA2"
                    ) {
                        $("#divpayment").css("display", "none");
                        $("#divpayment_2").css("display", "none");
                        $("#btnaddEmp").css("display", "none");
                        $("#Alien_Table").show();
                        $("#Alien_Table_None").hide();
                    } else if (status == "WA") {
                        $("#btnpayment").css("display", "block");
                        $("#head_step_2").css("display", "flex");
                        var statusDefult = "WA";
                        var found = false;

                        if (objData[0].formlist.length > 0) {
                            for (var f = 0; f < objData[0].formlist.length; f++) {
                                if (objData[0].formlist[f].form_status_id === statusDefult) {
                                    found = true;
                                    break;
                                }
                            }
                            if (found) {
                                $("#divpayment").css("display", "block");
                                $("#divpayment_2").css("display", "none");
                            } else {
                                if (
                                    status == "NCOS2" ||
                                    status == "NCOC2" ||
                                    status == "NCOCR2" ||
                                    status == "NCOCRD2"
                                ) {
                                    $("#divpayment").css("display", "none");
                                    $("#divpayment_2").css("display", "none");
                                } else {
                                    $("#divpayment").css("display", "block");
                                    $("#divpayment_2").css("display", "none");
                                }
                            }
                        }
                        var filecheck = "N";
                        jQuery("[name='file_name_doc']").each(function (index) {
                            var validatecheck = jQuery("#validate_" + this.id).val();
                            arraycheck.push(validatecheck);
                        });

                        for (var i = 0; i < arraycheck.length; i++) {
                            if (filecheck == arraycheck[i]) {
                                $("#head_step_2").hide();
                                $("#Alien_Table_None").css("display", "none");
                                $("#Alien_Table").hide();
                            } else {
                                if (arraycheck[i] == "Y") {
                                    $("#head_step_2").show();

                                    if (objData[0].formlist.length > 0) {
                                        $("#Alien_Table_None").css("display", "none");
                                        $("#Alien_Table").show();
                                    } else {
                                        $("#Alien_Table_None").css("display", "block");
                                        $("#Alien_Table").hide();
                                    }
                                } else {
                                    $("#head_step_2").show();
                                    $("#Alien_Table_None").css("display", "block");
                                }
                            }
                        }
                    } else if (status == "APSS") {
                        $("#Alien_Table").css("display", "none");
                        $("#Alien_Table_None").hide();
                        $("#divpayment").css("display", "none");
                        $("#head_step_2").css("display", "none");
                    } else if (
                        status == "NCOS2" ||
                        status == "NCOC2" ||
                        status == "NCOCR2" ||
                        status == "NCOCRD2"
                    ) {
                        $("#Alien_Table_None").hide();
                        $("#Alien_Table").css("display", "inline-table");
                        $("#divpayment").css("display", "none");
                        $("#btnpayment").css("display", "none");
                        $("#head_step_2").css("display", "flex");
                        $("#btnaddEmp").css("display", "none");
                        if (textcheck == "update") {
                            $("#btnpayment_2").css("display", "none");
                        }
                        var statusDefult = "WA";
                        var found = false;

                        if (objData[0].formlist.length > 0) {
                            for (var f = 0; f < objData[0].formlist.length; f++) {
                                if (objData[0].formlist[f].form_status_id === statusDefult) {
                                    found = true;
                                    break;
                                }
                            }
                            if (found) {
                                $("#divpayment").css("display", "none");
                                $("#divpayment_2").css("display", "block");
                            } else {
                                if (
                                    status == "NCOS2" ||
                                    status == "NCOC2" ||
                                    status == "NCOCR2" ||
                                    status == "NCOCRD2"
                                ) {
                                    $("#divpayment").css("display", "none");
                                    $("#divpayment_2").css("display", "none");
                                } else {
                                    $("#divpayment").css("display", "block");
                                    $("#divpayment_2").css("display", "none");
                                }
                            }
                            $("#Alien_Table").show();
                            $("#Alien_Table_None").hide();
                        }
                    } else if (status == "NCOS3") {
                        $("#Alien_Table").css("display", "none");
                        $("#Alien_Table_None").hide();

                        $("#divpayment").css("display", "none");
                        $("#btnpayment").css("display", "none");

                        $("#btnaddEmp").css("display", "none");

                        $("#head_step_2").css("display", "none");
                        if (textcheck == "update") {
                            $("#btnpayment_2").css("display", "none");
                        }
                    }

                    datatableE_Tracking = $(`#datatableE_Tracking`).DataTable({
                        paging: true,
                        lengthChange: false,
                        searching: false,
                        ordering: true,
                        destroy: true,
                        columns: [
                            { data: "form_id" },
                            { data: "form_id" },
                            { data: "form_type_name_th", className: "text-left" },
                            { data: "alien_first_name", className: "text-left" },
                            { data: "createdtimestamp" },
                            { data: "", className: "text-center" },
                            { data: "" },
                        ],
                        columnDefs: [
                            {
                                targets: 0,
                                width: "5%",
                                render: function (data, type, row, meta) {
                                    return `${meta.row + meta.settings._iDisplayStart + 1}`;
                                },
                            },
                            {
                                targets: 1,
                                render: function (data, type, row, meta) {
                                    return `${row.stay_permis_no}`;
                                },
                            },
                            {
                                targets: 2,
                                render: function (data, type, row, meta) {
                                    return `<span>${row.alien_fullname_en}</span>`;
                                },
                            },
                            {
                                targets: 3,
                                render: function (data, type, row, meta) {
                                    return `<span data_th='${row.nationality_th?.toSafeJsonHtml()}' data_en='${row.nationality_en
                                        }'></span>`;
                                },
                            },


                            {
                                targets: 4,
                                width: "20%",
                                render: function (data, type, row, meta) {
                                    if (
                                        row.form_status_id == "NCOS2" ||
                                        row.form_status_id == "NCOC2" ||
                                        row.form_status_id == "CCOS2" ||
                                        row.form_status_id == "NCOS3" ||
                                        row.form_status_id == "CC" ||
                                        row.form_status_id == "NCOCR2" ||
                                        row.form_status_id == "NCOCRD2"
                                    ) {
                                        return `<div class='badge badge-pill status12' style='word-wrap: break-word;white-space: normal;' data_th='${row.transection_status_th?.toSafeJsonHtml()}' data_en='${row.transection_status_en
                                            }'></div>`;
                                    } else if (
                                        row.form_status_id == "WA" ||
                                        row.form_status_id == "WCOSNA2" ||
                                        row.form_status_id == "WCOSAS2" ||
                                        row.form_status_id == "WCOCAS2" ||
                                        row.form_status_id == "WCOCOC2" ||
                                        row.form_status_id == "WCOSAS3" ||
                                        row.form_status_id == "WCOSNA3" ||
                                        row.form_status_id == "WCOCOCNA2" ||
                                        row.form_status_id == "WP3"
                                    ) {
                                        return `<div class='badge badge-pill status11-appoint' style='word-wrap: break-word;white-space: normal;' data_th='${row.transection_status_th?.toSafeJsonHtml()}' data_en='${row.transection_status_en
                                            }'></div>`;
                                    } else {
                                        return `<div class='badge badge-pill status13' style='word-wrap: break-word;white-space: normal;' data_th='${row.transection_status_th?.toSafeJsonHtml()}' data_en='${row.transection_status_en
                                            }'></div>`;
                                    }
                                },
                            },

                            {
                                targets: 5,
                                render: function (data, type, row, meta) {
                                    var objData_data = objData[0];
                                    if (
                                        row.form_status_id == "WA" ||
                                        row.form_status_id == "NCOS2" ||
                                        row.form_status_id == "NCOC2" ||
                                        row.form_status_id == "NCOCR2" ||
                                        row.form_status_id == "NCOCRD2" ||
                                        row.form_status_id == "NCOS3"
                                    ) {
                                        return `<a href="javascript:void(0)" class="lang_btn_edit_data btn btn-outline-primary action-button offcanvas-toggle text-center" style="font-size:14px;" onclick='OpenModalAddNameList(&quot;${row.form_id}&quot;,&quot;${row.form_status_id}&quot;,&quot;${objData_data.demand_letter_data.demand_letter_id}&quot;)'>แก้ไขข้อมูล</a>`;
                                    } else {
                                        return ``;
                                    }
                                },
                            },
                            {
                                targets: 6,
                                render: function (data, type, row, meta) {
                                    var getstatus = objData[0].group_status_id;


                                    var objData_data = objData[0];
                                    if (row.form_status_id == "WA") {
                                        return "";


                                    } else if (
                                        row.form_status_id == "NCOS2" ||
                                        row.form_status_id == "NCOC2" ||
                                        row.form_status_id == "NCOCR2" ||
                                        row.form_status_id == "NCOCRD2" ||
                                        row.form_status_id == "NCOS3" ||
                                        row.form_status_id == "AP"
                                    ) {
                                        return `<div class="dropdown">
                                                            <button class="btn btn-outline-primary" type="button" onclick="clickdropdown(event)" data-toggle="dropdown" aria-expanded="false" style="width: 32px;height: 32px;padding: 0px;">
                                                                <i class="fa-solid fa-ellipsis-vertical"></i>
                                                            </button>
                                                            <div class="dropdown-menu" style="">
                                                                <a href="javascript:void(0)" class="dropdown-item offcanvas-edit-toggle lang_watch_detail" style="font-weight:700" onclick="UploadFileMOU(&quot;${row.form_id}&quot;,&quot;${row.form_status_id}&quot;,&quot;${getstatus}&quot;)">ดูรายละเอียด</a>
                                                                <a href="javascript:void(0)" class="dropdown-item lang_reject_work" id="cancelAssign"  style="color:#FF3B30;font-weight:700" onclick="UpdateRejectAcceptWork(&quot;${objData_data.group_id}&quot;,&quot;${row.form_id}&quot;,&quot;${objData_data.request_by_user_id}&quot;)">ปฏิเสธการจ้างงาน</a>
                                                            </div>
                                                        </div>`;
                                    } else {
                                        return `<a href='javascript:void(0)' onclick="UploadFileMOU(&quot;${row.form_id}&quot;,&quot;${row.form_status_id}&quot;,&quot;${getstatus}&quot;)"><svg width="32" height="32" viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M0.5 6C0.5 2.96243 2.96243 0.5 6 0.5H26C29.0376 0.5 31.5 2.96243 31.5 6V26C31.5 29.0376 29.0376 31.5 26 31.5H6C2.96243 31.5 0.5 29.0376 0.5 26V6Z" stroke="#E24C86"/>
<path d="M16.0007 12.6654C16.9173 12.6654 17.6673 11.9154 17.6673 10.9987C17.6673 10.082 16.9173 9.33203 16.0007 9.33203C15.084 9.33203 14.334 10.082 14.334 10.9987C14.334 11.9154 15.084 12.6654 16.0007 12.6654ZM16.0007 14.332C15.084 14.332 14.334 15.082 14.334 15.9987C14.334 16.9154 15.084 17.6654 16.0007 17.6654C16.9173 17.6654 17.6673 16.9154 17.6673 15.9987C17.6673 15.082 16.9173 14.332 16.0007 14.332ZM16.0007 19.332C15.084 19.332 14.334 20.082 14.334 20.9987C14.334 21.9154 15.084 22.6654 16.0007 22.6654C16.9173 22.6654 17.6673 21.9154 17.6673 20.9987C17.6673 20.082 16.9173 19.332 16.0007 19.332Z" fill="#E24C86"/>
</svg></a >`;
                                    }
                                },
                            },
                        ],

                        data: objData[0].formlist,
                    });

                } else {
                    $("#divpayment").css("display", "none");
                    $("#divpayment_2").css("display", "none");

                    if (status == "APSS") {
                        $("#Alien_Table").css("display", "none");
                        $("#Alien_Table_None").hide();
                        $("#head_step_2").css("display", "none");
                    }
                    $("#div_wait_tap_three").removeClass("d-none");
                }

                $("#loading").delay().fadeOut("");
            }
        },
        "json"
    );
    await changeDatalang();

}


async function GetDataDemandLetterJobPermitCateList(
    form,
    demand_letter_id,
    form_status
) {
    let objRequest = {
        demand_letter_id: sessionStorage.getItem("demand_id"),
    };
    var request = await jQuery.post(
        url_GetDataDemandLetterJobPermitCateList,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                var setData = "";

                for (var i = 0; i < objData.length; i++) {
                    setData +=
                        "<option value='" +
                        objData[i].permit_cate_id +
                        "' data_th='" +
                        objData[i].permit_catework_name_th?.toSafeJsonHtml() +
                        "' data_en='" +
                        objData[i].permit_catework_name_en +
                        "'>" +
                        objData[i].permit_catework_name_th +
                        "</option>";
                }
                jQuery("#permit_cate_id").html(setData);
            }
            if (form) {
                AlienEdit(form, form_status);
            }
        },
        "json"
    );
    await Promise.all([  changeDatalang()]);
}
async function AlienEdit(form_id, form_status) {
    if (form_id) {
        if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
            var userid = getParameterByName("user_id");
            var group_id = getParameterByName("group_id");
            objRequest = {
                user_id: userid,
                group_id: group_id,
                form_id: form_id,
            };
        } else {
            var userid = core_datas.default_profile.user_id;
            var group_id = getParameterByName("group_id");
            objRequest = {
                user_id: userid,
                group_id: group_id,
                form_id: form_id,
            };
        }
        var request = await jQuery.post(
            url_GetDataRequestFormEtrackingMOU_ByFormID,
            objRequest,
            function (data) {
                if (data.status == "success") {
                    if (data.data.length != 0) {
                        var dataObj = data.data[0];
                        var selectedAlien = dataObj.alien_data;

                        if (form_status != "WA") {
                            if (dataObj.is_update_upload == "N") {
                                Swal.fire({
                                    imageUrl: important_data.warningIconUrl,
                                    title:
                                        '<span class="lang_Unable_to_edit_data">ไม่สามารถแก้ไขข้อมูลได้</span>',
                                    html: '<span class="lang_data_already_been_updated">เนื่องจากพบว่ามีการเเก้ไขข้มูลเรียบร้อยเเล้ว</span>',
                                    showCancelButton: false,
                                    confirmButtonText:
                                        '<span data_th="ตกลง" data_en="OK"></span>',
                                    reverseButtons: true,
                                }).then((result) => {
                                    $("#EmployeeModal").modal("hide");
                                    location.reload();

                                });
                                changeDatalang();
                            } else {
                                $("#EmployeeModal").modal("show");
                            }
                        } else {
                            $("#EmployeeModal").modal("show");
                        }

                        var getstatus = sessionStorage.getItem("group_status");

                        if (
                            form_status == "NCOS2" ||
                            form_status == "NCOC2" ||
                            form_status == "NCOCR2" ||
                            form_status == "NCOCRD2"
                        ) {
                            $("#EditNameList").attr("disabled", false);


                            $(".div_resson").removeClass("d-none");
                        } else if (form_status == "WA") {
                            $("#EditNameList").attr("disabled", false);
                            $(".div_resson").addClass("d-none");
                        } else {
                            $("#EditNameList").attr("disabled", true);
                            $(".div_resson").addClass("d-none");
                        }

                        if (getstatus == "NCOS2") {


                        }

                        let reasonHtml = "";

                        if (dataObj.consider_data.alien_reasonlist != null) {
                            var consider_data_alien_reasonlist =
                                dataObj.consider_data.alien_reasonlist
                                    .checkdocalien_reason_list;
                            if (consider_data_alien_reasonlist.length > 0) {
                                for (
                                    var g = 0;
                                    g < consider_data_alien_reasonlist.length;
                                    g++
                                ) {
                                    reasonHtml +=
                                        "<li><span data_th='" +
                                        consider_data_alien_reasonlist[
                                            g
                                        ].reason_checkdoc_name_th?.toSafeJsonHtml() +
                                        "' data_en='" +
                                        consider_data_alien_reasonlist[g].reason_checkdoc_name_en +
                                        "'></span></li>";
                                }
                            }
                        }

                        var reasonHtmlMore = "";

                        if (dataObj.consider_data.alien_reasonlist != null) {
                            if (
                                dataObj.consider_data.alien_reasonlist.checkdocalien_reason !=
                                ""
                            ) {
                                $(".ressonmore").removeClass("d-none");
                                $("#divShowreasonMore").removeClass("d-none");
                                reasonHtml += ` <hr>`;
                                reasonHtmlMore += `<li>${dataObj.consider_data.alien_reasonlist.checkdocalien_reason}</li>`;
                            }
                        }

                        $("#divShowreason").html(reasonHtml);
                        $("#divShowreasonMore").html(reasonHtmlMore);

                        $("#prefixNameEnDropdown").val(selectedAlien.pf_id);
                        $("#name_encheckreg").val(selectedAlien.firstname_en);
                        $("#middle_name_encheckreg").val(selectedAlien.middlename_en);
                        $("#lastname_encheckreg").val(selectedAlien.lastname_en);

                        $("#name_thcheckreg").val(selectedAlien.firstname_th);
                        $("#middle_name_thcheckreg").val(selectedAlien.middlename_th);
                        $("#lastname_thcheckreg").val(selectedAlien.lastname_th);

                        $("#nationality_al").val(selectedAlien.nationality_id);
                        $("#permit_cate_id").val(dataObj.permit_cate_id);
                        $("#job_description").val(dataObj.job_01);
                        $("#birthdate_al").val(selectedAlien.birth_date);

                        $("#age_al").val(selectedAlien.alien_age);
                        $("#tel_al").val(selectedAlien.phone);
                        $("#email_al").val(selectedAlien.email);

                        $("#form_id").val(form_id);
                        $("#alien_id").val(selectedAlien.alien_id);
                        $(
                            "#" + selectedAlien.alien_stay_permis[0].stay_permis_type_id
                        ).prop("checked", true);
                        $("#placepassport_al").val(
                            selectedAlien.alien_stay_permis[0].stay_permis_no
                        );
                        $("#issued_al").val(
                            selectedAlien.alien_stay_permis[0].stay_permis_issue_at
                        );
                        $("#country_al").val(
                            selectedAlien.alien_stay_permis[0].stay_permis_country_id
                        );
                        var dateofexpiry =
                            selectedAlien.alien_stay_permis[0].stay_permis_issue_dt
                                .toString()
                                .split(" ");
                        $("#dateofexpiry").val(dateofexpiry[0]);
                        var enddateofexpiry =
                            selectedAlien.alien_stay_permis[0].stay_permis_expire_dt
                                .toString()
                                .split(" ");
                        $("#enddateofexpiry").val(enddateofexpiry[0]);

                        $("#bus_type_id").val("data_th", dataObj.bus_type_name_th);
                        $("#bus_type_id").val("data_en", dataObj.bus_type_name_en);

                        $("#bus_type_id_value").val(dataObj.bus_type_id);
                        $("#permit_cate_id").val(dataObj.permit_cate_id);
                        if (dataObj.is_accept_work == "N") {
                            $("#is_accept_work_N").prop("checked", true);
                            $("#GetDataAcceptWorkInfo").css("display", "block");
                            $(".work_come_in").hide();
                            $("#uploadfilealien").hide();
                            var acceptworklist = dataObj.acceptworklist;
                            for (var j = 0; j < acceptworklist.length; j++) {
                                $("#work_info_" + acceptworklist[j].accept_work_id).prop(
                                    "checked",
                                    true
                                );
                                if (acceptworklist[j].accept_work_id == "0") {
                                    $(".work_info_detail").show();
                                    $("#accept_work_other").val(
                                        acceptworklist[j].accept_work_other
                                    );
                                }
                            }
                        } else {
                            $("#is_accept_work_Y").prop("checked", true);
                            $(".work_come_in").show();
                            $(".work_info_detail").hide();
                            $("#GetDataAcceptWorkInfo").css("display", "none");
                            $("#uploadfilealien").show();

                            var doclist = dataObj.doclist.extra_n_doclist;

                            var setData = "";
                            setData += "<div class='col-12 col-md-12'>";
                            setData +=
                                "<h3 class='mb-4 mt-3 head-step lang_doc_evidence_1'>เอกสารเเละหลักฐาน ครั้งที่ 1</h3>";
                            setData += "</div>";
                            for (let c = 0; c < doclist.length; c++) {
                                const requestDocList = doclist[c].requestdoclist;
                                setData += "<div class='col-md-12'>";
                                if (requestDocList.length >= 2) {
                                    setData += `<p class='text-upload'>${c + 1
                                        }${"."}<span class='lang_select_at_least_one_file'></span>`;
                                    if (doclist[c].required == "Y") {
                                        setData += "<span class='asterisk'>*</span>";
                                    }
                                    setData +=
                                        "<span id='err_file_name_" +
                                        doclist[c].group_form_type_doc_id +
                                        "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                                }

                                for (let j = 0; j < requestDocList.length; j++) {
                                    if (requestDocList.length < 2) {
                                        setData += "<div>";
                                        setData += `<p class='text-upload'>${c + 1
                                            }${"."}<span data_th='${requestDocList[
                                                j
                                            ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                            }'></span>`;
                                        if (doclist[c].required == "Y") {
                                            setData += "<span class='asterisk'>*</span>";
                                        }
                                        setData +=
                                            "<span id='err_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                                    } else {
                                        setData += '<div style="margin-left:15px;">';

                                        setData += `<p class='text-upload'><span data_th='${requestDocList[
                                            j
                                        ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                            }'></span>`;
                                    }

                                    setData += "</p>";
                                    setData += "</p>";
                                    setData += "<div class='row'>";
                                    setData += "<div class='col-lg-12'>";

                                    if (requestDocList[j].document_type_id === "72") {
                                        if (requestDocList[j].doclist.length > 0) {
                                            if (requestDocList[j].doclist[0]?.doc_new_lastest == "Y") {
                                                setData += "<div class='health_report_" + requestDocList[j].document_type_id + "''>"
                                                setData += "<span class='text-upload-subtext'>อัปเดตล่าสุด :</span> <span class='mr-3'>" + convert_to_thai_datetime(requestDocList[j].doclist[0]?.emp_created_timestamp) + "</span> <span class='badge badge-pill statusAPSS2' style='font-weight:600;'>อัปเดตเเล้ว</span>";
                                                setData += "</div>"
                                                setData += '<input type="hidden" id="doc_new_lastest_file_name_' + requestDocList[j].document_type_id + '" value="' + requestDocList[j].doclist[0]?.doc_new_lastest + '">'
                                            }
                                        } else {
                                            setData += requestDocList[j].document_type_id == "42"
                                                ? `<p class="text-upload-subtext lang_support_image_file">รองรับไฟล์ .PNG .JPG เเละ .JPEG ขนาดไม่เกิน 4 MB</p>`
                                                : `<p class="text-upload-subtext lang_support_pdf">รองรับไฟล์ .PDF ขนาดไม่เกิน 4 MB</p>`;
                                        }
                                    }

                                    if (requestDocList[j].document_type_id === "72") {
                                        if (requestDocList[j].chkis_upload == "N" || requestDocList[j].chkis_upload == "H" && requestDocList[j].doclist[0]?.insurance_no) {
                                            if (requestDocList[j].doclist?.length > 0) {
                                                setData += "<div class='health-report health_report_" + requestDocList[j].document_type_id +"'>";
                                                setData += `<span style='font-weight:600;color:#0A090A;'>เลขที่กรมธรรม :</span> <span>${requestDocList[j].doclist[0]?.insurance_no || "-"}</span><br>`;
                                                setData += `<span style='font-weight:600;color:#0A090A;'>หน่วยงานที่ทําประกันสุขภาพ :</span> <span>${requestDocList[j].doclist[0]?.insurance_from || "-"}</span><br>`;
                                                setData += `<span style='font-weight:600;color:#0A090A;'>สถานะทําประกันสุขภาพ :</span> <span style="color:${requestDocList[j].doclist[0]?.insurance_status_color_code}">${requestDocList[j].doclist[0]?.insurance_status || "-"}</span> (วันที่ซื้อประกัน <span data_th='${convert_to_thai_date(requestDocList[j].doclist[0]?.insurance_date_buy) || "-"}' data_en='${convert_to_eng_date(requestDocList[j].doclist[0]?.insurance_date_buy) || "-"}'></span>)<br>`;
                                                setData += `<span style='font-weight:600;color:#0A090A;'>วันทีเริ่มต้นประกัน :</span> <span data_th='${convert_to_thai_date(requestDocList[j].doclist[0]?.insurance_date_start) || "-"}' data_en='${convert_to_eng_date(requestDocList[j].doclist[0]?.insurance_date_start) || "-"}'></span><br>`;
                                                setData += `<span style='font-weight:600;color:#0A090A;'>วันทีสิ้นสุดประกัน :</span> <span data_th='${convert_to_thai_date(requestDocList[j].doclist[0]?.insurance_date_end) || "-"}' data_en='${convert_to_eng_date(requestDocList[j].doclist[0]?.insurance_date_end) || "-"}'></span>`;
                                                setData += "</div>";
                                            } else {
                                                setData +=
                                                    '<input type="hidden" id="insurance_file_name" value="N">';

                                            }
                                            setData +=
                                                '<input type="hidden" id="doc_new_lastest" value="' + requestDocList[j].doclist[0]?.doc_new_lastest +'">';
                                        }
                                        setData +=
                                            '<input type="hidden" id="insurance_condition" value="' + requestDocList[j].doclist[0]?.alien_health_insurance_check + '">';
                                        setData +=
                                            '<input type="hidden" id="chkis_upload" value="' + requestDocList[j].chkis_upload + '">';

                                    }


                                    setData += "<div class='row mb-3' style='margin-left:0px;'>";
                                    if (requestDocList[j].doclist.length != 0) {
                                        if (requestDocList[j].doclist[0].doc_status_id == "C") {
                                            setData +=
                                                "<div class='upload d-none' style='padding-right:14px;'>";
                                        } else {
                                            if (requestDocList[j].chkis_upload == "N") {

                                                setData +=
                                                    "<div class='upload d-none' style='padding-right:14px;'>";
                                            } else {
                                                setData +=
                                                    "<div class='upload' style='padding-right:14px;'>";
                                            }

                                        }
                                    } else {
                                        if (requestDocList[j].chkis_upload == "N") {

                                            setData +=
                                                "<div class='upload d-none' style='padding-right:14px;'>";
                                        } else {
                                            setData +=
                                                "<div class='upload' style='padding-right:14px;'>";
                                        }

                                    }

                                    let hasDoc = requestDocList[j].doclist.length !== 0;
                                    let isGroup = requestDocList.length >= 2;
                                    let isImage = requestDocList[j].document_type_id === "42";

                                    let nameAttr = isGroup ? "file_name_group" : "file_name";
                                    let acceptAttr = isImage ? " accept='image/png, image/jpeg, image/jpeg'" : "";
                                    let docId = "";
                                    if (hasDoc) {
                                        let tempId = requestDocList[j].doclist[0].document_id || "";
                                        docId = tempId.startsWith("https") ? "" : `,&quot;${tempId}&quot;`;
                                    }

                                    setData +=
                                        `<input type='file'${acceptAttr} name='${nameAttr}' id='file_name_${requestDocList[j].document_type_id}' ` +
                                        `onchange='SelectFileEdit(event,&quot;${requestDocList[j].document_type_id}&quot;${docId})' hidden/>`;

                                    setData +=
                                        "<label class='lang_lbl_selectfile' for='file_name_" +
                                        requestDocList[j].document_type_id +
                                        "' style='cursor:pointer;'>เลือกไฟล์</label>";


                                    setData += "</div>";
                                    if (requestDocList[j].doclist.length != 0) {
                                        if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                            setData +=
                                                "<div class='block-doc justify-content-between border_file_reject' id='divshow_file_name_" +
                                                requestDocList[j].document_type_id +
                                                "'>";
                                        } else {
                                            setData +=
                                                "<div class='block-doc justify-content-between border_file_success' id='divshow_file_name_" +
                                                requestDocList[j].document_type_id +
                                                "'>";
                                        }
                                    } else {
                                        setData +=
                                            "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' style='display:none;'>";
                                    }
                                    setData += "<div class='d-flex'>";
                                    setData += "<div>";
                                    setData += "<img src='/assets/images/document.svg'>";
                                    setData += "</div>";
                                    setData += "<div class='filename'>";
                                    if (requestDocList[j].doclist.length != 0) {
                                        if (requestDocList[j].chkis_upload == "N" || requestDocList[j].chkis_upload == "H" && requestDocList[j].doclist[0]?.insurance_no) {
                                            if (requestDocList[j].doclist[0].document_id.startsWith("https")) {
                                                setData += "<a id='taga_" + requestDocList[j].document_type_id + "' href='" + requestDocList[j].doclist[0].document_id + "' target='_blank'><span class='name-file namedoc' id='show_file_name_" +
                                                    requestDocList[j].document_type_id + "'>" + requestDocList[j].doclist[0].insurance_from + "</span></a>";
                                            } else {
                                                setData +=
                                                    "<a id='taga_" +
                                                    requestDocList[j].document_type_id +
                                                    "' onclick='GetDocFilePathByDocumentID(&quot;" +
                                                    requestDocList[j].doclist[0].document_id +
                                                    "&quot;,&quot;" +
                                                    requestDocList[j].form_type_doc_id +
                                                    "&quot,&quot;" +
                                                    dataObj.form_id +
                                                    "&quot)'><span class='name-file namedoc' id='show_file_name_" +
                                                    requestDocList[j].document_type_id +
                                                    "'>" +
                                                    requestDocList[j].doclist[0].insurance_from +
                                                    "</span></a>";
                                            }

                                        } else {
                                            setData +=
                                                "<a id='taga_" +
                                                requestDocList[j].document_type_id +
                                                "' onclick='GetDocFilePathByDocumentID(&quot;" +
                                                requestDocList[j].doclist[0].document_id +
                                                "&quot;,&quot;" +
                                                requestDocList[j].form_type_doc_id +
                                                "&quot,&quot;" +
                                                dataObj.form_id +
                                                "&quot)'><span class='name-file namedoc' id='show_file_name_" +
                                                requestDocList[j].document_type_id +
                                                "'>" +
                                                requestDocList[j].doclist[0].doc_title +
                                                "</span></a>";
                                        }

                                        setData +=
                                            '<input type="hidden" id="DocIDValue_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="' +
                                            requestDocList[j].document_id +
                                            '">';
                                    }
                                    setData +=
                                        "<p class='name-file namedoc' id='show_file_name_" +
                                        requestDocList[j].document_type_id +
                                        "'></p>";
                                    setData += "</div>";
                                    setData += "<div>";
                                    if (requestDocList[j].doclist.length != 0) {
                                        if (requestDocList[j].doclist[0].doc_status_id == "C" || requestDocList[j].chkis_upload == "N") {

                                        } else {
                                            setData +=
                                                "<img class='close' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                                requestDocList[j].document_type_id +
                                                "&quot;,&quot;" +
                                                requestDocList[j].doclist[0].document_id +
                                                "&quot;)'>";
                                        }
                                    } else {
                                        setData +=
                                            "<img class='close' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;)'>";
                                    }

                                    setData += "</div>";
                                    setData += "</div>";
                                    setData += "</div>";
                                    setData += "</div>";
                                    setData += "</div>";
                                    setData += "</div>";
                                    if (requestDocList[j].doclist.length != 0) {
                                        if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                            setData +=
                                                "<div class='resson' id='resson_" +
                                                requestDocList[j].document_type_id +
                                                "'><p class='textresson lang_text_reason'>หมายเหตุ</p><p> " +
                                                requestDocList[j].doclist[0].doc_comment +
                                                " </p></div>";
                                            setData +=
                                                '<input type="hidden" id="hidden_check_file_name_' +
                                                requestDocList[j].document_type_id +
                                                '" value="' +
                                                requestDocList[j].doclist[0].doc_status_id +
                                                '">';
                                        }
                                    }

                                    setData += "</div>";


                                    setData +=
                                        '<input type="hidden" id="requiredvalue_file_name_' +
                                        requestDocList[j].document_type_id +
                                        '" value="' +
                                        doclist[c].required +
                                        '">';

                                    setData +=
                                        '<input type="hidden" id="hidden_form_type_file_name_' +
                                        requestDocList[j].document_type_id +
                                        '" value="' +
                                        requestDocList[j].form_type_doc_id +
                                        '">';
                                    setData +=
                                        '<input type="hidden" id="required_file_name_' +
                                        requestDocList[j].document_type_id +
                                        '" value="' +
                                        doclist[c].group_form_type_doc_id +
                                        '">';

                                    if (requestDocList[j].doclist.length != 0) {
                                        setData +=
                                            '<input type="hidden" id="have_file_current_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="' +
                                            requestDocList[j].doclist[0].document_id +
                                            '">';
                                        setData +=
                                            '<input type="hidden" id="hidden_select_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="">';
                                    } else {
                                        setData +=
                                            '<input type="hidden" id="have_file_current_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="N">';
                                    }
                                }

                                setData += "</div>";


                                setData += "<div class='col-md-12 mb-2 mt-2'>";
                                setData += "<hr>";
                                setData += "</div>";
                            }

                            jQuery("#uploadfilealien").html(setData);
                        }
                    }
                }
            },
            "json"
        );
    }

    await changeDatalang();
}
async function GetDataFormAppointmentMoreDocByGroupID() {
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
        };
    }
    var request = await jQuery.post(
        url_GetDataFormAppointmentMoreDocByGroupID,
        objRequest,
        function (data) {
            var objData = data.data;

            var setData = "";

            if (data.status == "success") {
                $("#alien_show_num").text(objData.length);
                for (var i = 0; i < objData.length; i++) {
                    setData += '<div class="col-12 mt-4">';
                    setData += '<div class="payment">';
                    setData += '<label class="cursor-pointer w-100">';
                    setData += '<div class="border rounded p-4 bg-white shadow-primary">';
                    setData += '<div class="row">';
                    setData += '<div class="col-lg-9 col-md-8 col-12 pr-md-4 pr-auto">';
                    setData +=
                        '<h4 class="mb-3" style="font-size: 18px; font-weight: bold;"><span class="lang_quantity"></span> ' +
                        objData[i].num_alien_appointment +
                        ' <sapn class="lang_person"></span></h4>';
                    if (objData[i].appointment_datetime == "CC") {
                        setData +=
                            '<h5 class="text-left">รายชื่อเเรงงานต่างด้าวที่ถูกปฏิเสธการจ้างงาน</h5>';
                    } else {
                        setData += '<h5 class="text-left">';
                        if (
                            objData[i].appointment_datetime != "" ||
                            objData[i].appointment_datetime != null
                        ) {
                            var appointment_datetime =
                                objData[i].appointment_datetime.split("-");
                            var sp_appointment_datetime =
                                appointment_datetime[2] +
                                "/" +
                                appointment_datetime[1] +
                                "/" +
                                appointment_datetime[0];
                        }

                        setData +=
                            '<span class="lang_appointment_date" style="color: #6F6F85;">วันที่นัดหมาย </span> <span>:</span><span data_th="' +
                            convert_to_thai_date(sp_appointment_datetime) +
                            '" data_en="' +
                            convert_to_eng_date(sp_appointment_datetime) +
                            '"></span>&nbsp;';
                        setData +=
                            '<span style="color: #EAEAED;">| &nbsp; </span> <span class="lang_location" style="color: #6F6F85;">สถานที่</span> <span>:</span> <span data_th="' +
                            objData[i].locatename_th?.toSafeJsonHtml() +
                            '" data_en="' +
                            objData[i].locatename_en +
                            '"></span>  &nbsp;';
                        setData +=
                            '<span style="color: #EAEAED;">| &nbsp; </span> <span class="lang_upload_doc" style="color: #6F6F85;">อัปโหลดเอกสาร</span> <span>:</span> <span style="color:#E24C86">' +
                            objData[i].num_alien_upload +
                            "</span>/" +
                            objData[i].num_alien_appointment +
                            ' <span class="lang_person"></span> &nbsp;';
                        setData +=
                            '<span style="color: #EAEAED;">| &nbsp; </span> <span class="lang_confirm_send_doc" style="color: #6F6F85;">ยืนยันการส่งเอกสาร </span> <span>:</span> <span style="color:#E24C86">' +
                            objData[i].num_alien_send +
                            "</span>/" +
                            objData[i].num_alien_appointment +
                            ' <span class="lang_person"></span> &nbsp;';
                        setData += "</h5>";
                    }

                    setData += "</div>";
                    setData +=
                        '<div class="col-lg-3 col-md-4 col-12 d-flex justify-content-lg-end justify-content-sm-start justify-content-center align-items-center pt-md-0 pt-3 column-gap-20">';

                    const locateTh = parseIfJson(objData[i].locatename_th);
                    jsonStr = JSON.stringify(locateTh);


                    setData +=
                        '<a href="javascript:void(0)" class="add-btn lang_watch_detail" value="" onclick="GetDataFormAppointmentMoreDocByAppointmentDate(&quot;' +
                        objData[i].appointment_datetime +
                        "&quot;,&quot;" +
                        sp_appointment_datetime +
                        '&quot)">ดูรายละเอียด</a>';
                    setData += "</div>";
                    setData += "</div>";
                    setData += "</div>";
                    setData += "</label>";
                    setData += "</div>";
                    setData += "</div>";
                }
                $("#alien_appointment_more").html(setData);

                $("#alien_appointment_more_2").html(setData);
            }
        },
        "json"
    );
    await changeDatalang();
}
async function GetDataFormAppointmentMoreDocByAppointmentDate(
    appointment_datetime,
    sp_appointment_datetime
) {
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            appointment_datetime: appointment_datetime,
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            appointment_datetime: appointment_datetime,
        };
    }
    var request = await jQuery.post(
        url_GetDataFormAppointmentMoreDocByAppointmentDate,
        objRequest,
        function (data) {
            if (data.status == "success") {

                var objData = data.data;
                var formlist = objData.formlist;

                $("#appointment_datetime").val(appointment_datetime);
                if (formlist.length != "0") {
                    $("#div_appointment1").css("display", "none");
                    $("#div_appointment2").css("display", "none");
                    $("#div_appointment").css("display", "none");
                    $("#div_appointment0").css("display", "none");
                    $("#div_appointment3").css("display", "none");


                    $("#div_Detail_more").css("display", "block");
                    $("#tabShowTableEmployee").css("display", "none");
                    $("#statusAPSS").css("display", "none");
                    var statusDefult = "WCD";
                    var found = false;

                    if (formlist.length > 0) {
                        for (var f = 0; f < formlist.length; f++) {
                            if (formlist[f].form_status_id === statusDefult) {
                                found = true;
                                break;
                            }
                        }
                    }

                    if (found) {
                        $("#btn_SendDoc").attr("disabled", false);
                    } else {
                        $("#btn_SendDoc").attr("disabled", true);
                    }
                    $("#alien_upload").text(objData.num_alien_upload);
                    $("#head_text_alien").text(objData.num_alien_all);
                    $("#text_num_al").text(objData.num_alien_all);
                    if (appointment_datetime == "CC") {
                        $(".div_no_CC").addClass("d-none");
                        $(".div_CC").removeClass("d-none");
                        $("#btn_SendDoc").hide();
                    } else {
                        $("#btn_SendDoc").show();
                        $(".div_CC").addClass("d-none");
                        $(".div_no_CC").removeClass("d-none");

                        $("#alien_upload_all").text(objData.num_alien_all);
                        $("#alien_send").text(objData.num_alien_send);
                        $("#all_user").text(objData.num_alien_all);
                        $("#location_ap").attr("data_th", jsonStr);
                        var dateresult = "";
                        if (appointment_datetime) {
                            var splitdate = appointment_datetime.split("-");
                            dateresult =
                                splitdate[2] + "/" + splitdate[1] + "/" + splitdate[0];
                        }
                        $("#sp_appointment_datetime").attr(
                            "data_th",
                            convert_to_thai_date(dateresult)
                        );
                        $("#sp_appointment_datetime").attr(
                            "data_en",
                            convert_to_eng_date(dateresult)
                        );
                        changeDatalang()
                    }

                    datatable_Appointment_DetailMore = $(
                        `#datatable_Appointment_DetailMore`
                    ).DataTable({
                        paging: true,
                        lengthChange: false,
                        searching: false,
                        ordering: true,
                        destroy: true,
                        columns: [
                            { data: "stay_permis_no", className: "text-left" },
                            { data: "alien_name_en", className: "text-left" },
                            { data: "alien_nationality_th", className: "text-left" },
                            { data: "alien_date_of_birth", className: "text-left" },
                            { data: "stay_permis_no", className: "text-left" },
                            { data: "stay_permis_no", className: "text-left" },
                            { data: "stay_permis_no", className: "text-left" },
                        ],
                        columnDefs: [
                            {
                                targets: 1,
                                render: function (data, type, row, meta) {
                                    return `<span>${row.alien_name_en}</span><br><span style="color:#808080;"><span class="lang_ref_code">หมายเลขอ้างอิง</span> : ${row.alien_id}</span>`;
                                },
                            },
                            {
                                targets: 2,
                                render: function (data, type, row, meta) {
                                    return ` <span data_th='${row.alien_nationality_th?.toSafeJsonHtml()}' data_en='${row.alien_nationality_en
                                        }'></span>`;
                                },
                            },
                            {
                                targets: 3,
                                render: function (data, type, row, meta) {
                                    return ` <span data_th='${convert_to_thai_date(
                                        row.alien_date_of_birth
                                    )}' data_en='${convert_to_eng_date(
                                        row.alien_date_of_birth
                                    )}'></span>`;
                                },
                            },

                            {
                                targets: 4,
                                render: function (data, type, row, meta) {
                                    return (
                                        '<span class="badgestatus" style="color:' +
                                        row.status_color_font +
                                        ";background-color:" +
                                        row.status_color +
                                        '"><span data_th="' +
                                        row.transection_status_th?.toSafeJsonHtml() +
                                        '" data_en="' +
                                        row.transection_status_en +
                                        '"></span></span>'
                                    );
                                },
                            },
                            {
                                targets: 5,
                                render: function (data, type, row, meta) {
                                    var getstatus = sessionStorage.getItem("group_status");
                                    var status_notshow = [
                                        "WCD",
                                        "WCOSNA3",
                                        "WCOSAS3",
                                        "SS",
                                        "WCPOS",
                                        "WPC",
                                        "CC",
                                        "WBI",
                                        "CCOS2",
                                    ];
                                    if (status_notshow.includes(row.form_status_id)) {
                                        return ``;
                                    } else {
                                        return `<a href="javascript:void(0)" class="lang_upload btn btn-outline-primary action-button offcanvas-toggle lang_upload" onclick="UploadFileMOU(&quot;${row.form_id}&quot;,&quot;${row.form_status_id}&quot;,&quot;${getstatus}&quot;)">อัปโหลด</a>`;
                                    }
                                },
                            },
                            {
                                targets: 6,
                                render: function (data, type, row, meta) {
                                    var getstatus = sessionStorage.getItem("group_status");
                                    var group_id = getParameterByName("group_id");
                                    var user_id = core_datas.default_profile.user_id;
                                    var word_detail = "detail";
                                    if (
                                        row.form_status_id == "APSS" ||
                                        row.form_status_id == "WCD" ||
                                        row.form_status_id == "AP" || row.form_status_id == "NCOS3" || row.form_status_id == "WCPOS" || row.form_status_id == "WBI" || row.form_status_id == "WPC"
                                    ) {
                                        return `<div class="dropdown show">
  <a class="btn btn-secondary dropdown-toggle button-dropdown-height" href="#" role="button" id="dropdownMenuLink" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
   <i class="fa-solid fa-ellipsis-vertical"></i>
  </a>  <div class="dropdown-menu" aria-labelledby="dropdownMenuLink">
    <a href="javascript:void(0)" class="dropdown-item offcanvas-edit-toggle lang_watch_detail" style="font-weight:700" onclick="DetailMOU(&quot;${row.form_id}&quot;,&quot;${row.form_status_id}&quot;,&quot;${getstatus}&quot;,&quot;${word_detail}&quot;)">ดูรายละเอียด</a>
                                                                  <a href="javascript:void(0)" class="dropdown-item lang_reject_work" id="cancelAssign"  style="color:#FF3B30;font-weight:700" onclick="UpdateRejectAcceptWork(&quot;${group_id}&quot;,&quot;${row.form_id}&quot;,&quot;${user_id}&quot;)">ปฏิเสธการจ้างงาน</a>
  </div>
</div>`;
                                    } else {
                                        return `<a href='javascript:void(0)' onclick="DetailMOU(&quot;${row.form_id}&quot;,&quot;${row.form_status_id}&quot;,&quot;${getstatus}&quot;,&quot;${word_detail}&quot;)"><svg width="32" height="32" viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M0.5 6C0.5 2.96243 2.96243 0.5 6 0.5H26C29.0376 0.5 31.5 2.96243 31.5 6V26C31.5 29.0376 29.0376 31.5 26 31.5H6C2.96243 31.5 0.5 29.0376 0.5 26V6Z" stroke="#E24C86"/>
<path d="M16.0007 12.6654C16.9173 12.6654 17.6673 11.9154 17.6673 10.9987C17.6673 10.082 16.9173 9.33203 16.0007 9.33203C15.084 9.33203 14.334 10.082 14.334 10.9987C14.334 11.9154 15.084 12.6654 16.0007 12.6654ZM16.0007 14.332C15.084 14.332 14.334 15.082 14.334 15.9987C14.334 16.9154 15.084 17.6654 16.0007 17.6654C16.9173 17.6654 17.6673 16.9154 17.6673 15.9987C17.6673 15.082 16.9173 14.332 16.0007 14.332ZM16.0007 19.332C15.084 19.332 14.334 20.082 14.334 20.9987C14.334 21.9154 15.084 22.6654 16.0007 22.6654C16.9173 22.6654 17.6673 21.9154 17.6673 20.9987C17.6673 20.082 16.9173 19.332 16.0007 19.332Z" fill="#E24C86"/>
</svg></a >`;
                                    }
                                },
                            },
                        ],

                        data: formlist,
                        drawCallback: function (settings) {

                            changeDatalang();
                        },
                    });
                    $("#loading").delay().fadeOut("");
                }
            }
        },
        "json"
    );
}
function UpdateRejectAcceptWork(group_id, form_id, user_id) {
    user_reject = user_id;
    form_reject = form_id;
    group_reject = group_id;
    Swal.fire({
        title: '<span class="lang_reject_work">ปฎิเสธการจ้างงาน</span>',
        html: "<span class='lang_Do_you_want_reject_employment'>คุณต้องการปฎิเสธการจ้างงานใช่หรือไม่</span>",
        imageUrl: important_data.warningIconUrl,
        showCancelButton: true,
        cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
        confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
    }).then((result) => {
        if (result.isConfirmed) {
            GetDataAcceptWorkInfo();
            checkselectacceptwork();
            $("#ModalUpdateRejectAcceptWork").modal("show");
        }
    });
    changeDatalang();
}
function DeleteDataNameList(group_id, form_id, user_id) {
    Swal.fire({
        title:
            "<span class='lang_Cancel_foreign_worker_record'>ยกเลิกรายชื่อแรงงานต่างด้าว</span>",
        html: "<span class='lang_Cancel_foreign_worker_record_Once_confirmed'>ท่านต้องการการยกเลิกรายชื่อแรงงานต่างด้าว เมื่อยืนยันแล้ว รายชื่อนี้จะไม่สามารถใช้งานได้อีก ท่านต้องการยกเลิกรายชื่อแรงงานต่างด้าวหรือไม่?</span>",
        imageUrl: important_data.warningIconUrl,
        showCancelButton: true,
        cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
        confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
    }).then(async (result) => {
        if (result.isConfirmed) {
            const isTokenValid = await validate_token_presubmit();
            if (!isTokenValid) {
                return false;
            }
            let objRequest = {
                user_id: user_id,
                group_id: group_id,
                form_id: form_id,
                accept_work_id: "0",
                accept_work_other: "ลบรายชื่อเเรงงานต่างด้าว",
                relate_account_id: core_datas.current_profile.profile_id,
            };
            var request = jQuery.post(
                url_UpdateRejectAcceptWork,
                objRequest,
                function (data) {
                    if (data.status == "success") {
                        Swal.fire({
                            title:
                                "<span class='lang_Cancel_foreign_worker_record'>ยกเลิกรายชื่อแรงงานต่างด้าว</span>",
                            html: "<span class='lang_Foreign_record_cancel_successful'>ยกเลิกรายชื่อเเรงงานต่างด้าวสำเร็จ</span>",
                            imageUrl: important_data.successIconUrl,
                        }).then((result) => {
                            location.reload();

                        });
                        changeDatalang();
                    } else {
                        Swal.fire({
                            title:
                                "<span class='lang_Cancel_foreign_worker_record'>ยกเลิกรายชื่อแรงงานต่างด้าว</span>",
                            html: "<span class='lang_Foreign_record_cancel_Failed'>ยกเลิกรายชื่อเเรงงานต่างด้าวไม่สำเร็จ</span>",
                            imageUrl: important_data.errorIconUrl,
                        }).then((result) => {
                            location.reload();

                        });
                        changeDatalang();
                    }
                },
                "json"
            );
        }
    });
    changeDatalang();
}
function UploadFileMOU(form_id, form_status_id, getstatus, detail) {
    $("#empuploadfilealien").empty();
    $("#listuploadfile_alien").empty();
    $("#listuploadfile_alien_1").empty();
    $("#empuploadfilealien_1").empty();
    $("#form_id").val(form_id);

    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            form_id: form_id,
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            form_id: form_id,
        };
    }

    var request = jQuery.post(
        url_GetDataRequestFormEtrackingMOU_ByFormID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var dataObj = data.data[0];
                var selectedAlien = dataObj.alien_data;

                var statusshow = ["NCOS2", "NCOC2", "NCOCR2", "NCOCRD2", "NCOS3"];
                if (statusshow.includes(form_status_id)) {
                    if (dataObj.is_update_upload == "N") {
                        Swal.fire({
                            imageUrl: errorIconUrl,
                            title:
                                '<span class="lang_Notification">เเจ้งเตือน</span>',
                            html: '<span class="lang_documents_have_already_been_added">เนื่องจากพบว่ามีการเพิ่มเอกสารเรียบร้อยเเล้ว</span>',
                            confirmButtonText: "<span data_th='ตกลง' data_en='OK'></span>",
                        }).then((result) => {
                            location.reload();
                        });
                        changeDatalang();
                        return false;
                    }
                }

                $("#UploadFileMOU").modal("show");
                $(".div_resson_3").addClass("d-none");
                $(".alien_name").text(
                    selectedAlien.firstname_en +
                    " " +
                    " " +
                    selectedAlien.middlename_en +
                    " " +
                    selectedAlien.lastname_en
                );
                $(".alien_nationality").attr("data_th", selectedAlien.nationality_th);
                $(".alien_nationality").attr("data_en", selectedAlien.nationality_en);
                $(".alien_birth").attr(
                    "data_th",
                    convert_to_thai_date(selectedAlien.birth_date) || "-"
                );
                $(".alien_birth").attr(
                    "data_en",
                    convert_to_eng_date(selectedAlien.birth_date) || "-"
                );
                $(".alien_address").text(selectedAlien.lastname_en);
                $(".alien_tel").text(selectedAlien.phone || "-");

                $(".alien_email").text(selectedAlien.email || "-");
                $("#age_al").text(selectedAlien.alien_age);
                $("#tel_al").text(selectedAlien.phone);
                $("#email_al").text(selectedAlien.email);

                $("#form_id").val(form_id);
                $("#alien_id").val(selectedAlien.alien_id);
                $(".data_job").text(dataObj.job_01);

                $(".stay_permis_name").attr("data_th", selectedAlien.alien_stay_permis[0].stay_permis_type_name_th);
                $(".stay_permis_name").attr("data_en", selectedAlien.alien_stay_permis[0].stay_permis_type_name_en);
                $(".stay_permis_no").text(
                    selectedAlien.alien_stay_permis[0].stay_permis_no
                );
                $(".stay_permis_issue_at").text(
                    selectedAlien.alien_stay_permis[0].stay_permis_issue_at
                );
                $(".stay_permis_issue_dt").attr(
                    "data_th",
                    convert_to_thai_date(
                        selectedAlien.alien_stay_permis[0].stay_permis_issue_dt
                    )
                );
                $(".stay_permis_issue_dt").attr(
                    "data_en",
                    convert_to_eng_date(
                        selectedAlien.alien_stay_permis[0].stay_permis_issue_dt
                    )
                );

                $(".stay_permis_expire_dt").attr(
                    "data_th",
                    convert_to_thai_date(
                        selectedAlien.alien_stay_permis[0].stay_permis_expire_dt
                    )
                );
                $(".stay_permis_expire_dt").attr(
                    "data_en",
                    convert_to_eng_date(
                        selectedAlien.alien_stay_permis[0].stay_permis_expire_dt
                    )
                );

                if (form_status_id == "NCOS3") {
                    $("#AddNewMore").css("display", "none");
                    $("#UpdateDocMore").removeClass("d-none");
                    $(".div_resson_3").removeClass("d-none");

                    let reasonHtml = "";

                    if (dataObj.consider_data.alien_reasonlist != null) {
                        var consider_data_alien_reasonlist =
                            dataObj.consider_data.alien_reasonlist.checkdocalien_reason_list;
                        if (consider_data_alien_reasonlist.length > 0) {
                            for (var g = 0; g < consider_data_alien_reasonlist.length; g++) {
                                reasonHtml +=
                                    "<li data_th='" +
                                    consider_data_alien_reasonlist[
                                        g
                                    ].reason_checkdoc_name_th?.toSafeJsonHtml() +
                                    "' data_en='" +
                                    consider_data_alien_reasonlist[g].reason_checkdoc_name_en +
                                    "'></li>";
                            }
                        }
                    }

                    var reasonHtmlMore = "";

                    if (dataObj.consider_data.alien_reasonlist != null) {
                        if (
                            dataObj.consider_data.alien_reasonlist.checkdocalien_reason != ""
                        ) {
                            $(".ressonmore_3").removeClass("d-none");
                            $("#divShowreasonMore_3").removeClass("d-none");
                            reasonHtml += ` <hr>`;
                            reasonHtmlMore += `<li>${dataObj.consider_data.alien_reasonlist.checkdocalien_reason}</li>`;
                        }
                    }

                    $("#divShowreason_3").html(reasonHtml);
                    $("#divShowreasonMore_3").html(reasonHtmlMore);
                } else if (
                    form_status_id == "WCOSAS3" ||
                    form_status_id == "WCOSNA3" ||
                    form_status_id == "SS" ||
                    form_status_id == "WCPOS"
                ) {
                    $("#can_button").css("display", "none");
                    $("#AddNewMore").css("display", "none");
                    $("#UpdateDocMore").css("display", "none");
                    $(".div_resson_3").addClass("d-none");
                } else if (form_status_id == "CC") {
                    $(".div_resson_3").removeClass("d-none");
                    let reasonHtml = "";

                    if (dataObj.acceptworklist.length != 0) {
                        for (var g = 0; g < dataObj.acceptworklist.length; g++) {
                            if (dataObj.acceptworklist[g].accept_work_id == "0") {
                                reasonHtml +=
                                    "<li><span data_th='" +
                                    dataObj.acceptworklist[
                                        g
                                    ].accept_work_name_th?.toSafeJsonHtml() +
                                    "' data_en='" +
                                    dataObj.acceptworklist[g].accept_work_name_en +
                                    "'></span><span> : " +
                                    dataObj.acceptworklist[g].accept_work_other +
                                    "</span></li>";
                            } else {
                                reasonHtml +=
                                    "<li data_th='" +
                                    dataObj.acceptworklist[
                                        g
                                    ].accept_work_name_th?.toSafeJsonHtml() +
                                    "' data_en='" +
                                    dataObj.acceptworklist[g].accept_work_name_en +
                                    "'></li>";
                            }
                        }
                    }

                    $("#divShowreason_3").html(reasonHtml);
                }

                var statuswcos = [
                    "WCOSNA2",
                    "WCOSAS2",
                    "CCEDO",
                    "NCOS2",
                    "WCOCAS2",
                    "WCOCNA2",
                    "WCOCAS2",
                    "NCOC2",
                    "WCOCOCNA2",
                    "WCOCOC2",
                    "WCOCOC2",
                    "NCOCR2",
                    "NCOCRD2",
                    "CCOS2",
                    "AP",
                    "CCAPO",
                    "APSS",
                    "CC",
                    "CCMAP",
                    "NCOS3",
                    "WCOSAS3",
                    "WCOSNA3",
                    "WCPOS",
                    "SS",
                    "CCPC",
                    "WPC",
                    "WP3",
                ];

                if (statuswcos.includes(form_status_id)) {
                    $("#div_shoefile_detail").css("display", "block");
                    $("#empuploadfilealien").css("display", "none");
                    $("#can_button").css("display", "none");
                    $("#AddNewMore").css("display", "none");
                    $("#UpdateDocMore").css("display", "none");
                    var doclist = dataObj.doclist.extra_n_doclist;

                    var setDatadocfile_list = "";
                    setDatadocfile_list +=
                        "<div><h3 class='mb-4 mt-3 head-step lang_doc_evidence_1'>เอกสารหลักฐานครั้งที่ 1</h3></div>";
                    for (let c = 0; c < doclist.length; c++) {
                        const requestDocList = doclist[c].requestdoclist;

                        setDatadocfile_list += "<div>";
                        if (requestDocList.length >= 2) {
                            setDatadocfile_list += `<p class='text-upload'>${c + 1
                                }${"."}<span class='lang_select_at_least_one_file'></span>`;
                        }

                        for (let j = 0; j < requestDocList.length; j++) {

                            setDatadocfile_list += '<div class="row">';
                            if (requestDocList.length < 2) {
                                setDatadocfile_list += "<div class='col-md-6'>";

                                setDatadocfile_list += `<p class='text-upload'>${c + 1
                                    }${"."}<span data_th='${requestDocList[
                                        j
                                    ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                    }'></span>`;
                            } else {
                                setDatadocfile_list += "<div class='col-md-6'>";
                                setDatadocfile_list += '<div style="margin-left:15px;">';
                                setDatadocfile_list += `<p class='text-upload'><span data_th='${requestDocList[
                                    j
                                ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                    }'></span>`;
                                setDatadocfile_list += "</p>";
                                setDatadocfile_list += "</div>";
                            }

                            setDatadocfile_list += "</p>";
                            setDatadocfile_list += "</div>";
                            setDatadocfile_list +=
                                "<div class='col-md-6' style='text-align:end;'>";

                            if (requestDocList[j].doclist.length > 0) {
                                setDatadocfile_list +=
                                    "<a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                    dataObj.form_id +
                                    " &quot;, &quot; " +
                                    requestDocList[j].doclist[0].document_id +
                                    " &quot;, &quot; " +
                                    requestDocList[j].form_type_doc_id +
                                    " &quot;)'></a>";

                            } else {
                                setDatadocfile_list += "-";
                            }

                            setDatadocfile_list += "</div>";
                            setDatadocfile_list += "</div>";
                            setDatadocfile_list += "</div>";


                        }
                        setDatadocfile_list += "<div class='col-md-12 mb-2 mt-2'>";
                        setDatadocfile_list += "<hr>";
                        setDatadocfile_list += "</div>";
                    }

                    $("#div_listuploadfile_alien_1").removeClass("d-none");

                    jQuery("#listuploadfile_alien_1").html(setDatadocfile_list);

                    if (getstatus == "SS" || form_status_id == "SS") {
                        $("#div_empuploadfilealien_1").removeClass("d-none");
                        $("#can_button").css("display", "none");
                        $("#AddNewMore").css("display", "none");
                        var doclist = dataObj.doclist.extra_y_doclist;
                        var setDatadocfile_list3 = "";
                        setDatadocfile_list3 +=
                            "<div><h3 class='mb-4 mt-3 head-step lang_doc_evidence_2'>เอกสารเเละหลักฐาน ครั้งที่ 2</h3></div>";
                        for (let c = 0; c < doclist.length; c++) {
                            const requestDocList = doclist[c].requestdoclist;

                            setDatadocfile_list3 += "<div>";
                            if (requestDocList.length >= 2) {
                                setDatadocfile_list3 += `<p class='text-upload'>${c + 1
                                    }${"."}<span class='lang_select_at_least_one_file'></span>`;
                            }

                            for (let j = 0; j < requestDocList.length; j++) {

                                setDatadocfile_list3 += '<div class="row">';
                                if (requestDocList.length < 2) {
                                    setDatadocfile_list3 += "<div class='col-md-6'>";

                                    setDatadocfile_list3 += `<p class='text-upload'>${c + 1
                                        }${"."}<span data_th='${requestDocList[
                                            j
                                        ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                        }'></span>`;
                                } else {
                                    setDatadocfile_list3 += "<div class='col-md-6'>";
                                    setDatadocfile_list3 += '<div style="margin-left:15px;">';
                                    setDatadocfile_list3 += `<p class='text-upload'><span data_th='${requestDocList[
                                        j
                                    ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                        }'></span>`;
                                    setDatadocfile_list3 += "</p>";
                                    setDatadocfile_list3 += "</div>";
                                }

                                setDatadocfile_list3 += "</p>";
                                setDatadocfile_list3 += "</div>";
                                setDatadocfile_list3 +=
                                    "<div class='col-md-6' style='text-align:end;'>";

                                if (requestDocList[j].doclist.length > 0) {
                                    setDatadocfile_list3 +=
                                        "<a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook'  onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                        dataObj.form_id +
                                        " &quot;, &quot; " +
                                        requestDocList[j].doclist[0].document_id +
                                        " &quot;, &quot; " +
                                        requestDocList[j].form_type_doc_id +
                                        " &quot;)'></a>";

                                } else {
                                    setDatadocfile_list3 += "-";
                                }

                                setDatadocfile_list3 += "</div>";
                                setDatadocfile_list3 += "</div>";
                                setDatadocfile_list3 += "</div>";


                            }
                            setDatadocfile_list3 += "<div class='col-md-12 mb-2 mt-2'>";
                            setDatadocfile_list3 += "<hr>";
                            setDatadocfile_list3 += "</div>";
                        }

                        jQuery("#empuploadfilealien_1").html(setDatadocfile_list3);
                    }
                }

                if (form_status_id == "APSS" || form_status_id == "NCOS3") {
                    if (
                        detail == "detail" &&
                        form_status_id == "NCOS3" &&
                        getstatus == "WCOSAS3"
                    ) {
                        $("#can_button").css("display", "none");
                        $("#AddNewMore").css("display", "none");
                        $("#UpdateDocMore").css("display", "none");
                    } else if (form_status_id == "WCOSNA3" && getstatus == "APSS") {
                        $("#can_button").css("display", "none");
                        $("#AddNewMore").css("display", "none");
                        $("#div_listuploadfile_alien_1").removeClass("d-none");
                    } else if (form_status_id == "NCOS3") {
                        $("#can_button").css("display", "block");
                        $("#UpdateDocMore").css("display", "block");
                        $("#AddNewMore").css("display", "none");
                    } else {
                        $("#can_button").css("display", "block");
                        $("#AddNewMore").css("display", "block");
                    }

                    $("#div_shoefile_detail").css("display", "none");
                    $("#empuploadfilealien").css("display", "contents");

                    var doclist = dataObj.doclist.extra_y_doclist;
                    var setData = "";
                    setData += "<div class='col-12 col-md-12'>";
                    setData +=
                        "<h3 class='mb-4 mt-3 head-step lang_doc_evidence_2'>เอกสารเเละหลักฐาน ครั้งที่ 2</h3>";
                    setData += "</div>";
                    for (let c = 0; c < doclist.length; c++) {
                        const requestDocList = doclist[c].requestdoclist;
                        setData += "<div class='col-md-12'>";
                        if (requestDocList.length >= 2) {
                            setData += `<p class='text-upload'>${c + 1
                                }${"."}<span class='lang_select_at_least_one_file'></span>`;
                            if (doclist[c].required == "Y") {
                                setData += "<span class='asterisk'>*</span>";
                            }
                            setData +=
                                "<span id='err_file_name_" +
                                doclist[c].group_form_type_doc_id +
                                "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                        }

                        for (let j = 0; j < requestDocList.length; j++) {
                            if (requestDocList.length < 2) {
                                setData += "<div>";
                                setData += `<p class='text-upload'>${c + 1
                                    }${"."}<span data_th='${requestDocList[
                                        j
                                    ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                    }'></span>`;
                                if (doclist[c].required == "Y") {
                                    setData += "<span class='asterisk'>*</span>";
                                }
                                setData +=
                                    "<span id='err_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                            } else {
                                setData += '<div style="margin-left:15px;">';

                                setData += `<p class='text-upload'><span data_th='${requestDocList[
                                    j
                                ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                    }'></span>`;
                            }

                            setData += "</p>";
                            setData += "</p>";
                            setData += "<div class='row'>";
                            setData += "<div class='col-lg-12'>";


                            if (requestDocList[j].document_type_id === "41") {
                                if (requestDocList[j].doclist.length > 0) {
                                    if (requestDocList[j].doclist[0]?.doc_new_lastest == "Y") {
                                        setData += "<div class='health_report_" + requestDocList[j].document_type_id + "''>"
                                        setData += "<span class='text-upload-subtext'>อัปเดตล่าสุด :</span> <span class='mr-3'>" + convert_to_thai_datetime(requestDocList[j].doclist[0]?.emp_created_timestamp) + "</span> <span class='badge badge-pill statusAPSS2' style='font-weight:600;'>อัปเดตเเล้ว</span>";
                                        setData += "</div>"
                                        setData += '<input type="hidden" id="doc_new_lastest_file_name_' + requestDocList[j].document_type_id + '" value="' + requestDocList[j].doclist[0]?.doc_new_lastest + '">'
                                    }
                                } else {
                                    setData += requestDocList[j].document_type_id == "42"
                                        ? `<p class="text-upload-subtext lang_support_image_file">รองรับไฟล์ .PNG .JPG เเละ .JPEG ขนาดไม่เกิน 4 MB</p>`
                                        : `<p class="text-upload-subtext lang_support_pdf">รองรับไฟล์ .PDF ขนาดไม่เกิน 4 MB</p>`;
                                }
                            }


                            if (requestDocList[j].document_type_id === "41") {
                                if (requestDocList[j].chkis_upload == "N" || requestDocList[j].chkis_upload == "H" && requestDocList[j].doclist[0]?.health_from) {
                                    if (requestDocList[j].doclist?.length > 0) {
                                        setData += "<div class='health-report health_report_" + requestDocList[j].document_type_id +"'>";
                                        setData += `<span style='font-weight:600;color:#0A090A;'>ผลการตรวจสุขภาพจาก :</span> <span>${requestDocList[j].doclist[0]?.health_from || "-"}</span><br>`;
                                        setData += `<span style='font-weight:600;color:#0A090A;'>วันที่ตรวจสุขภาพ :</span> <span data_th='${convert_to_thai_date(requestDocList[j].doclist[0]?.health_date) || "-"}' data_en='${convert_to_eng_date(requestDocList[j].doclist[0]?.health_date) || "-"}'></span><br>`;
                                        setData += `<span style='font-weight:600;color:#0A090A;'>ผลการตรวจสุขภาพ :</span> <span style="color:${requestDocList[j].doclist[0]?.health_result_color_code}">${requestDocList[j].doclist[0]?.health_result || "-"}</span>`;
                                        setData += "</div>";
                                        setData += '<input type="hidden" id="health_report" value="Y">';
                                    } else {
                                        setData +=
                                            '<input type="hidden" id="health_report" value="N">';
                                    }
                                    setData +=
                                        '<input type="hidden" id="doc_new_lastest" value="' + requestDocList[j].doclist[0]?.doc_new_lastest + '">';

                                    setData +=
                                        '<input type="hidden" id="health_condition" value="' + requestDocList[j].doclist[0]?.alien_health_insurance_check + '">';
                                }

                                setData +=
                                    '<input type="hidden" id="chkis_upload" value="' + requestDocList[j].chkis_upload + '">';
                            }

                            setData += "<div class='row mb-3' style='margin-left:0px;'>";

                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList[j].doclist[0].doc_status_id == "C") {
                                    setData +=
                                        "<div class='upload d-none' style='padding-right:14px;'>";
                                } else {
                                    if (requestDocList[j].chkis_upload == "N") {

                                        setData +=
                                            "<div class='upload d-none' style='padding-right:14px;'>";
                                    } else {
                                        setData +=
                                            "<div class='upload' style='padding-right:14px;'>";
                                    }

                                }
                            } else {
                                if (requestDocList[j].chkis_upload == "N") {

                                    setData +=
                                        "<div class='upload d-none' style='padding-right:14px;'>";
                                } else {
                                    setData +=
                                        "<div class='upload' style='padding-right:14px;'>";
                                }

                            }

                            let hasDoc = requestDocList[j].doclist.length !== 0;
                            let isGroup = requestDocList.length >= 2;
                            let isImage = requestDocList[j].document_type_id === "42";

                            let nameAttr = isGroup ? "file_name_emp_group" : "file_name_emp";
                            let acceptAttr = isImage ? " accept='image/png, image/jpeg, image/jpeg'" : "";
                            let docId = "";
                            if (hasDoc) {
                                let tempId = requestDocList[j].doclist[0].document_id || "";
                                docId = tempId.startsWith("https") ? "" : `,&quot;${tempId}&quot;`;
                            }

                            setData +=
                                `<input type='file'${acceptAttr} name='${nameAttr}' id='file_name_${requestDocList[j].document_type_id}' ` +
                                `onchange='SelectFileEdit(event,&quot;${requestDocList[j].document_type_id}&quot;${docId})' hidden/>`;


                            setData +=
                                "<label class='lang_lbl_selectfile' for='file_name_" +
                                requestDocList[j].document_type_id +
                                "' style='cursor:pointer;'>เลือกไฟล์</label>";
                            if (requestDocList[j].doclist.length != 0) {
                                setData +=
                                    "<span class='lang_nofile_chosen' id='divselect_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' style='display:none;'>No file chosen</span>";
                            } else {
                                setData +=
                                    "<span class='lang_nofile_chosen' id='divselect_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "'>No file chosen</span>";
                            }
                            setData += "</div>";
                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                    setData +=
                                        "<div class='block-doc justify-content-between border_file_reject' id='divshow_file_name_" +
                                        requestDocList[j].document_type_id +
                                        "'>";
                                } else {
                                    setData +=
                                        "<div class='block-doc justify-content-between border_file_success' id='divshow_file_name_" +
                                        requestDocList[j].document_type_id +
                                        "'>";
                                }
                            } else {
                                setData +=
                                    "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' style='display:none;'>";
                            }
                            setData += "<div class='d-flex'>";
                            setData += "<div>";
                            setData += "<img src='/assets/images/document.svg'>";
                            setData += "</div>";
                            setData += "<div class='filename'>";
                            if (requestDocList[j].doclist.length != 0) {

                                if (requestDocList[j].chkis_upload == "N" || requestDocList[j].chkis_upload == "H" && requestDocList[j].doclist[0]?.health_from) {

                                    if (requestDocList[j].doclist[0].document_id.startsWith("https")) {
                                        setData += "<a id='taga_" + requestDocList[j].document_type_id + "' href='" + requestDocList[j].doclist[0].document_id + "' target='_blank'><span class='name-file namedoc' id='show_file_name_" +
                                            requestDocList[j].document_type_id + "'>" + requestDocList[j].doclist[0].health_from + "</span></a>";
                                    } else {
                                        setData +=
                                            "<a id='taga_" +
                                            requestDocList[j].document_type_id +
                                            "' onclick='GetDocFilePathByDocumentID(&quot;" +
                                            requestDocList[j].doclist[0].document_id +
                                            "&quot;,&quot;" +
                                            requestDocList[j].form_type_doc_id +
                                            "&quot,&quot;" +
                                            dataObj.form_id +
                                            "&quot)'><span class='name-file namedoc' id='show_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "'>" +
                                            requestDocList[j].doclist[0].health_from +
                                            "</span></a>";
                                    }


                                } else {
                                    setData +=
                                        "<a id='taga_" +
                                        requestDocList[j].document_type_id +
                                        "' onclick='GetDocFilePathByDocumentID(&quot;" +
                                        requestDocList[j].doclist[0].document_id +
                                        "&quot;,&quot;" +
                                        requestDocList[j].form_type_doc_id +
                                        "&quot,&quot;" +
                                        dataObj.form_id +
                                        "&quot)'><span class='name-file namedoc' id='show_file_name_" +
                                        requestDocList[j].document_type_id +
                                        "'>" +
                                        requestDocList[j].doclist[0].doc_title +
                                        "</span></a>";
                                }

                            }
                            setData +=
                                "<p class='name-file namedoc' id='show_file_name_" +
                                requestDocList[j].document_type_id +
                                "'></p>";
                            setData += "</div>";
                            setData += "<div>";
                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList[j].doclist[0].doc_status_id == "C" || requestDocList[j].chkis_upload == "N") {


                                } else {
                                    setData +=
                                        "<img class='close close_file' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                        requestDocList[j].document_type_id +
                                        "&quot;,&quot;" +
                                        requestDocList[j].doclist[0].document_id +
                                        "&quot;)'>";
                                }
                            } else {
                                setData +=
                                    "<img class='close close_file' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)'>";
                            }

                            setData += "</div>";
                            setData += "</div>";
                            setData += "</div>";
                            setData += "</div>";
                            setData += "</div>";
                            setData += "</div>";
                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                    setData +=
                                        "<div class='resson' id='resson_" +
                                        requestDocList[j].document_type_id +
                                        "'><p class='textresson lang_text_reason'>หมายเหตุ</p><p> " +
                                        requestDocList[j].doclist[0].doc_comment +
                                        " </p></div>";
                                    setData +=
                                        '<input type="hidden" id="hidden_check_file_name_' +
                                        requestDocList[j].document_type_id +
                                        '" value="' +
                                        requestDocList[j].doclist[0].doc_status_id +
                                        '">';
                                }
                            }

                            setData += "</div>";
                            setData +=
                                '<input type="hidden" id="requiredvalue_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                doclist[c].required +
                                '">';

                            setData +=
                                '<input type="hidden" id="hidden_form_type_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].form_type_doc_id +
                                '">';
                            setData +=
                                '<input type="hidden" id="required_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                doclist[c].group_form_type_doc_id +
                                '">';

                            if (requestDocList[j].doclist.length !== 0) {
                                setData +=
                                    '<input type="hidden" id="have_file_current_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="' +
                                    requestDocList[j].doclist[0].document_id +
                                    '">';
                                setData +=
                                    '<input type="hidden" id="hidden_select_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="">';
                            } else {
                                setData +=
                                    '<input type="hidden" id="have_file_current_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="N">';
                            }

                            if (requestDocList[j].document_type_id === "41") {
                                if (requestDocList[j].chkis_upload == "N") {
                                    if (requestDocList[j].doclist.length == 0) {
                                        setData += `
                                  <div class="BorderEMP w-100 ml-4">
                                    <div class="d-flex align-items-center pl-2">
                                      <img src="/assets/images/warning.svg" class="pr-2" style="height:24px;">
                                      <span>
                                       เนื่องจากยังไม่มีผลตรวจสุขภาพ กรุณานำรหัสอ้างอิงคนต่างด้าว หรือเลขที่หนังสือเดินทาง ไปตรวจสุขภาพกับโรงพยาบาลที่ได้รับอนุญาตจากกรมการจัดหางานก่อน จึงจะสามารถดำเนินการในขั้นตอนถัดไปได้
                                      </span>
                                    </div>
                                  </div>
                                `;
                                    }
                                } else if (requestDocList[j].chkis_upload == "H") {
                                    if (requestDocList[j].doclist.length == 0) {
                                        setData += `
                                  <div class="BorderEMP w-100 ml-4">
                                    <div class="d-flex align-items-center pl-2">
                                      <img src="/assets/images/warning.svg" class="pr-2" style="height:24px;">
                                      <span>
                                        ไม่พบข้อมูลผลตรวจสุขภาพ กรุณาตรวจสุขภาพผ่านบริษัทที่ได้รับอนุญาตโดยใช้รหัสอ้างอิงคนต่างด้าวหรือเลขหนังสือเดินทาง หากตรวจสุขภาพเรียบร้อยแล้วโปรดแนบเอกสาร
                                      </span>
                                    </div>
                                  </div>
                                `;
                                    }
                                }
                            }
                        }

                        setData += "</div>";
                        setData += "<div class='col-md-12 mb-2 mt-2'>";
                        setData += "<hr>";
                        setData += "</div>";
                    }
                    jQuery("#empuploadfilealien").html(setData);
                }

                if (form_status_id == "WCOSNA3" && getstatus == "NCOS3") {
                    $(".upload").hide();
                    $(".close_file").hide();
                }


                if (
                    detail == undefined &&
                    form_status_id == "NCOS3" &&
                    getstatus == "WCOSAS3"
                ) {
                    $("#div_empuploadfilealien_1").css("display", "none");
                }
                changeDatalang();
            }
        },
        "json"
    );
}
function DetailMOU(form_id, form_status_id, getstatus, detail) {
    jQuery("#empuploadfilealien_detail_1").empty();
    $("#form_detail_alien")[0].reset();
    $("#DetailMOU").modal("show");

    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            form_id: form_id,
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            form_id: form_id,
        };
    }

    var request = jQuery.post(
        url_GetDataRequestFormEtrackingMOU_ByFormID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var dataObj = data.data[0];

                var selectedAlien = dataObj.alien_data;
                var reasonHtmlCC = "";
                $("#accept_reason ").empty();
                $(".div_resson_3").addClass("d-none");
                $("#alien_name").text(
                    selectedAlien.firstname_en +
                    " " +
                    selectedAlien.middlename_en +
                    " " +
                    selectedAlien.lastname_en
                );
                $("#alien_nationality").text(selectedAlien.nationality_th);
                $("#alien_birth").attr(
                    "data_th",
                    convert_to_thai_date(selectedAlien.birth_date) || "-"
                );
                $("#alien_birth").attr(
                    "data_en",
                    convert_to_eng_date(selectedAlien.birth_date) || "-"
                );
                $("#alien_address").text(selectedAlien.lastname_en);
                $("#alien_tel").text(selectedAlien.phone || "-");

                $("#alien_email").text(selectedAlien.email || "-");
                $("#age_al").text(selectedAlien.alien_age);
                $("#tel_al").text(selectedAlien.phone);
                $("#email_al").text(selectedAlien.email);

                $("#form_id").val(form_id);
                $("#alien_id").val(selectedAlien.alien_id);

                $(".stay_permis_name").attr("data_th", selectedAlien.alien_stay_permis[0].stay_permis_type_name_th);
                $(".stay_permis_name").attr("data_en", selectedAlien.alien_stay_permis[0].stay_permis_type_name_en);

                $("#stay_permis_no").text(
                    selectedAlien.alien_stay_permis[0].stay_permis_no
                );
                $("#stay_permis_issue_at").text(
                    selectedAlien.alien_stay_permis[0].stay_permis_issue_at
                );

                $("#stay_permis_issue_dt").attr(
                    "data_th",
                    convert_to_thai_date(
                        selectedAlien.alien_stay_permis[0].stay_permis_issue_dt
                    )
                );
                $("#stay_permis_issue_dt").attr(
                    "data_en",
                    convert_to_eng_date(
                        selectedAlien.alien_stay_permis[0].stay_permis_issue_dt
                    )
                );

                $("#stay_permis_expire_dt").attr(
                    "data_th",
                    convert_to_thai_date(
                        selectedAlien.alien_stay_permis[0].stay_permis_expire_dt
                    )
                );
                $("#stay_permis_expire_dt").attr(
                    "data_en",
                    convert_to_eng_date(
                        selectedAlien.alien_stay_permis[0].stay_permis_expire_dt
                    )
                );
                $("#data_job").text(dataObj.job_01);

                if (form_status_id == "WCOSNA3" || form_status_id == "WCD") {
                    $(".div_resson_3").hide();
                } else if (form_status_id == "CC") {
                    $("#accept_reason ").removeClass("d-none");
                    $(".div_resson_3").removeClass("d-none");
                    var acceptworklist = dataObj.acceptworklist;
                    if (acceptworklist.length != 0) {
                        for (var j = 0; j < acceptworklist.length; j++) {
                            if (acceptworklist[j].accept_work_id == "0") {
                                reasonHtmlCC +=
                                    "<li><span data_th='" +
                                    acceptworklist[j].accept_work_name_th?.toSafeJsonHtml() +
                                    "' data_en='" +
                                    acceptworklist[j].accept_work_name_en +
                                    "'></span><span> : " +
                                    acceptworklist[j].accept_work_other +
                                    "</span></li>";
                            } else {
                                reasonHtmlCC +=
                                    "<li data_th='" +
                                    acceptworklist[j].accept_work_name_th?.toSafeJsonHtml() +
                                    "' data_en='" +
                                    acceptworklist[j].accept_work_name_en +
                                    "'></li>";
                            }
                        }
                        $("#accept_reason").html(reasonHtmlCC);
                    }
                }

                $("#UpdateDocMore").hide();
                $("#empuploadfilealien_detail").css("display", "none");
                $("#div_listuploadfile_alien_detail_1").removeClass("d-none");
                var doclist = dataObj.doclist.extra_n_doclist;
                var setDatadocfile_list2 = "";
                setDatadocfile_list2 +=
                    "<div><h3 class='mb-4 mt-3 head-step lang_doc_evidence_1'>เอกสารเเละหลักฐาน ครั้งที่ 1</h3></div>";
                for (let c = 0; c < doclist.length; c++) {
                    const requestDocList = doclist[c].requestdoclist;

                    setDatadocfile_list2 += "<div>";
                    if (requestDocList.length >= 2) {
                        setDatadocfile_list2 += `<p class='text-upload'>${c + 1
                            }${"."}<span class='lang_select_at_least_one_file'></span>`;
                    }

                    for (let j = 0; j < requestDocList.length; j++) {

                        setDatadocfile_list2 += '<div class="row">';
                        if (requestDocList.length < 2) {
                            setDatadocfile_list2 += "<div class='col-md-6'>";

                            setDatadocfile_list2 += `<p class='text-upload'>${c + 1
                                }${"."}<span data_th='${requestDocList[
                                    j
                                ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                }'></span>`;
                        } else {
                            setDatadocfile_list2 += "<div class='col-md-6'>";
                            setDatadocfile_list2 += '<div style="margin-left:15px;">';
                            setDatadocfile_list2 += `<p class='text-upload'><span data_th='${requestDocList[
                                j
                            ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                }'></span>`;
                            setDatadocfile_list2 += "</p>";
                            setDatadocfile_list2 += "</div>";
                        }

                        setDatadocfile_list2 += "</p>";
                        setDatadocfile_list2 += "</div>";
                        setDatadocfile_list2 +=
                            "<div class='col-md-6' style='text-align:end;'>";

                        if (requestDocList[j].doclist.length > 0) {
                            setDatadocfile_list2 +=
                                "<a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook'  onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                dataObj.form_id +
                                " &quot;, &quot; " +
                                requestDocList[j].doclist[0].document_id +
                                " &quot;, &quot; " +
                                requestDocList[j].form_type_doc_id +
                                " &quot;)'></a>";

                        } else {
                            setDatadocfile_list2 += "-";
                        }

                        setDatadocfile_list2 += "</div>";
                        setDatadocfile_list2 += "</div>";
                        setDatadocfile_list2 += "</div>";


                    }
                    setDatadocfile_list2 += "<div class='col-md-12 mb-2 mt-2'>";
                    setDatadocfile_list2 += "<hr>";
                    setDatadocfile_list2 += "</div>";
                }
                jQuery("#listuploadfile_alien_detail_1").html(setDatadocfile_list2);

                if (form_status_id != "APSS") {
                    $("#div_empuploadfilealien_detail_1").removeClass("d-none");
                    $("#can_button").css("display", "none");
                    $("#AddNewMore").css("display", "none");
                    var doclist = dataObj.doclist.extra_y_doclist;
                    var setDatadocfile_list3 = "";
                    setDatadocfile_list3 +=
                        "<div><h3 class='mb-4 mt-3 head-step lang_doc_evidence_2'>เอกสารเเละหลักฐาน ครั้งที่ 2</h3></div>";
                    for (let c = 0; c < doclist.length; c++) {
                        const requestDocList = doclist[c].requestdoclist;

                        setDatadocfile_list3 += "<div>";
                        if (requestDocList.length >= 2) {
                            setDatadocfile_list3 += `<p class='text-upload'>${c + 1
                                }${"."}<span class='lang_select_at_least_one_file'></span>`;
                        }

                        for (let j = 0; j < requestDocList.length; j++) {

                            setDatadocfile_list3 += '<div class="row">';
                            if (requestDocList.length < 2) {
                                setDatadocfile_list3 += "<div class='col-md-6'>";

                                setDatadocfile_list3 += `<p class='text-upload'>${c + 1
                                    }${"."}<span data_th='${requestDocList[
                                        j
                                    ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                    }'></span>`;
                            } else {
                                setDatadocfile_list3 += "<div class='col-md-6'>";
                                setDatadocfile_list3 += '<div style="margin-left:15px;">';
                                setDatadocfile_list3 += `<p class='text-upload'><span data_th='${requestDocList[
                                    j
                                ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                    }'></span>`;
                                setDatadocfile_list3 += "</p>";
                                setDatadocfile_list3 += "</div>";
                            }

                            setDatadocfile_list3 += "</p>";
                            setDatadocfile_list3 += "</div>";
                            setDatadocfile_list3 +=
                                "<div class='col-md-6' style='text-align:end;'>";

                            if (requestDocList[j].doclist.length > 0) {
                                setDatadocfile_list3 +=
                                    "<a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook'  onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                    dataObj.form_id +
                                    " &quot;, &quot; " +
                                    requestDocList[j].doclist[0].document_id +
                                    " &quot;, &quot; " +
                                    requestDocList[j].form_type_doc_id +
                                    " &quot;)'></a>";

                            } else {
                                setDatadocfile_list3 += "-";
                            }

                            setDatadocfile_list3 += "</div>";
                            setDatadocfile_list3 += "</div>";
                            setDatadocfile_list3 += "</div>";


                        }
                        setDatadocfile_list3 += "<div class='col-md-12 mb-2 mt-2'>";
                        setDatadocfile_list3 += "<hr>";
                        setDatadocfile_list3 += "</div>";
                    }
                    jQuery("#empuploadfilealien_detail_1").html(setDatadocfile_list3);
                }

                changeDatalang();
            }
        },
        "json"
    );
}
async function AppointmentMoreDocForSendToCheck() {
    const isTokenValid = await validate_token_presubmit();
    if (!isTokenValid) {
        return false;
    }
    checkChecked();
    $("#confirm_send_btn").attr("disabled", true);
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            appointment_datetime: $("#appointment_datetime").val(),
            txt_search: $("#searchBtn").val(),
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            appointment_datetime: $("#appointment_datetime").val(),
            txt_search: $("#searchBtn").val(),
        };
    }

    var request = jQuery.post(
        url_GetDataFormAppointmentMoreDocForSendToCheck,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;

                $("#sendDocument").modal("show");

                AppointmentMoreDocForSend = $(`#AppointmentMoreDocForSend`).DataTable({
                    paging: true,
                    lengthChange: false,
                    searching: false,
                    ordering: true,
                    destroy: true,
                    columns: [
                        { data: "form_id", className: "text-left" },
                        { data: "stay_permis_no", className: "text-left" },
                        { data: "alien_name_en", className: "text-left" },
                        { data: "alien_nationality_th", className: "text-left" },
                        { data: "alien_date_of_birth", className: "text-left" },
                        { data: "stay_permis_no", className: "text-left" },
                    ],
                    columnDefs: [
                        {
                            targets: 0,
                            render: function (data, type, row, meta) {
                                return `<div class="checkbox mb-0"> <label class="mb-0"><input type="checkbox" name="form_sendDoc" id="${row.form_id}" onclick="checkChecked()"><span class="custom-checkbox"><i class="icon-check"></i></span></label> </div>`;
                            },
                        },
                        {
                            targets: 3,
                            render: function (data, type, row, meta) {
                                return `<span data_th='${row.alien_nationality_th?.toSafeJsonHtml()}' data_en='${row.alien_nationality_en
                                    }'> </span>`;
                            },
                        },
                        {
                            targets: 4,
                            render: function (data, type, row, meta) {
                                return `<span data_th='${convert_to_thai_date(
                                    row.alien_date_of_birth
                                )}' data_en='${convert_to_eng_date(
                                    row.alien_date_of_birth
                                )}'> </span>`;
                            },
                        },
                        {
                            targets: 5,
                            render: function (data, type, row, meta) {
                                if (row.form_status_id == "WCD") {
                                    return ``;
                                } else {
                                    return `<a href="javascript:void(0)" class="lang_upload btn btn-outline-primary action-button offcanvas-toggle" onclick="UploadFileMOU(&quot;${row.form_id}&quot;)">อัปโหลด</a>`;
                                }
                            },
                        },
                    ],

                    data: objData,
                    drawCallback: function (settings) {

                        changeDatalang();
                    },
                });
                $("#loading").delay().fadeOut("");
            }
        },
        "json"
    );
}
function checkChecked() {
    let atLeastOneChecked = false;
    $("#AppointmentMoreDocForSend")
        .DataTable()
        .rows()
        .every(function () {
            const row = this.node();
            const checkbox = $(row).find("input[name='form_sendDoc']");


            if (checkbox.prop("checked")) {
                atLeastOneChecked = true;
                return false;
            }
        });

    if (atLeastOneChecked) {
        $("#confirm_send_btn").attr("disabled", false);
    } else {
        $("#confirm_send_btn").attr("disabled", true);
    }

    $("#checkAll").prop("checked", false);
}

function OnchangeYear(data_year, data_month, data_day) {
    var days = $("#hidden_req_permit").val() || data_day;
    var years = Math.floor(days / 365);
    var months = Math.floor((days % 365) / 30);
    var remainingDays = days - years * 365 - months * 30;

    var selectedYear = parseInt(data_year);
    var selectedMonth = parseInt(data_month);
    var maxMonths = selectedYear === years ? months : 11;
    var maxDays =
        selectedYear === years && selectedMonth === months
            ? remainingDays
            : daysInMonth(selectedMonth, selectedYear);

    $("#month_work").find("option").remove().end();
    $("#day_work").find("option").remove().end();

    for (var j = 0; j <= maxMonths; j++) {
        $("#month_work").append(
            $("<option>", {
                value: j,
                text: j,
            })
        );
    }

    for (var k = 0; k <= maxDays; k++) {
        $("#day_work").append(
            $("<option>", {
                value: k,
                text: k,
            })
        );
    }
    function daysInMonth(month, year) {
        return new Date(year, month, 30).getDate();
    }
    $("#year_work").val(data_year);
    $("#month_work").val(data_month);
    $("#day_work").val(data_day);
}
async function GetDataNationalityInfoEdit(
    nationality_id_mou,
    recruitment_agency_id,
    fsc_id,
    immigration_checkpoint_id
) {
    let objRequest = {
        is_mou: "Y",
    };
    var request = await jQuery.post(
        url_GetDataNationalityInfo,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;

                var setData = "";
                if (nationality_id_mou == "") {
                    setData +=
                        "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                    for (var i = 0; i < objData.length; i++) {
                        setData +=
                            "<option value='" +
                            objData[i].nationality_id +
                            "' data_th='" +
                            objData[i].nationality_th?.toSafeJsonHtml() +
                            "' data_en='" +
                            objData[i].nationality_en +
                            "'>" +
                            objData[i].nationality_th +
                            "</option>";
                    }
                    jQuery("#nationality_id_mou").html(setData);
                } else {
                    if (nationality_id_mou != "") {
                        $("#nationality_id_mou").val(nationality_id_mou);
                        GetDataRecruitmentAgencyedit(
                            nationality_id_mou,
                            recruitment_agency_id
                        );
                        GetDataImmigrationCheckpointEdit(
                            nationality_id_mou,
                            fsc_id,
                            immigration_checkpoint_id
                        );
                    }
                }
            }
        },
        "json"
    );
    await changeDatalang();
}
async function GetDataRecruitmentAgencyedit(
    value_nationality,
    recruitment_agency_id
) {
    $("#recruitment_agency").removeAttr("disabled");
    let objRequest = {
        nationality_id: value_nationality,
    };

    var request = await jQuery.post(
        url_GetDataRecruitmentAgency,
        objRequest,
        function (data) {
            var StringData = JSON.stringify(data);

            if (data.status == "success") {
                var objData = data.data;

                var setData = "";
                if (value_nationality != "") {
                    setData +=
                        "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                    for (var i = 0; i < objData.length; i++) {
                        if (objData[i].active == "Y" || objData[i].active == "y") {
                            setData +=
                                "<option data-value='" +
                                objData[i].recruitment_agency_address +
                                "' value='" +
                                objData[i].recruitment_agency_id +
                                "' data_th='" +
                                objData[i].recruitment_agency_name_th +
                                "' data_en='" +
                                objData[i].recruitment_agency_name_en +
                                "'>" +
                                objData[i].recruitment_agency_name_th +
                                "</option>";
                        }
                    }

                    jQuery("#recruitment_agency").html(setData);
                }
                $("#recruitment_agency").val(recruitment_agency_id);
            }
        },
        "json"
    );
    await changeDatalang();
}
async function GetDataImmigrationCheckpointEdit(
    value_nationality,
    fsc_id,
    immigration_checkpoint_id
) {
    $("#DataImmigrationCheckpoint").removeAttr("disabled");
    $("#fsc_id").removeAttr("disabled");

    let objRequest = {
        nationality_id: value_nationality,
    };

    var request = await jQuery.post(
        url_GetDataImmigrationCheckpoint,
        objRequest,
        function (data) {
            var StringData = JSON.stringify(data);

            if (data.status == "success") {
                var objData = data.data;

                var setData = "";
                setData +=
                    "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    if (objData[i].active == "Y" || objData[i].active == "y") {
                        setData +=
                            "<option value='" +
                            objData[i].immigration_checkpoint_id +
                            "' data_th='" +
                            objData[i].immigration_checkpoint_name_th?.toSafeJsonHtml() +
                            "'  data_en='" +
                            objData[i].immigration_checkpoint_name_en +
                            "'>" +
                            objData[i].immigration_checkpoint_name_th +
                            "</option>";
                    }
                }

                jQuery("#DataImmigrationCheckpoint").html(setData);
                $("#DataImmigrationCheckpoint").val(immigration_checkpoint_id);
                GetDataFirstServiceCenterEdit("", value_nationality, fsc_id);
                $("#DataImmigrationCheckpoint").on("change", function () {
                    GetDataFirstServiceCenterEdit(this.value, value_nationality, fsc_id);
                });
            }
        },
        "json"
    );
    await changeDatalang();
}

async function GetDataFirstServiceCenterEdit(
    value_select,
    value_nationality,
    fsc_id
) {
    let objRequest = {
        nationality_id: value_nationality,
    };

    var request = await jQuery.post(
        url_GetDataFirstServiceCenter,
        objRequest,
        function (data) {
            var StringData = JSON.stringify(data);

            if (data.status == "success") {
                var objData = data.data;

                var setData = "";
                setData +=
                    "<option class='lang_opt_pleaseselect' value=''>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    if (objData[i].active == "Y" || objData[i].active == "y") {
                        setData +=
                            "<option value='" +
                            objData[i].fsc_id +
                            "' data_th='" +
                            objData[i].fsc_name_th +
                            "'  data_en='" +
                            objData[i].fsc_name_en +
                            "'>" +
                            objData[i].fsc_name_th +
                            "</option>";
                    }
                }

                jQuery("#fsc_id").html(setData);
                $("#fsc_id").val(fsc_id);
                if (value_select != "") {
                    $("#fsc_id").val(value_select);
                }
            }
        },
        "json"
    );
    await changeDatalang();
}

var selectFile46_59Upload = function (event, id) {
    var files = document.getElementById("file_name_" + id + "").files;

    if (files.length > 0) {
        var isValid = checkTypeNameSizeFile(id, files[0]);

        if (!isValid) {
            $("#file_name_" + id).val("");
            return false;
        }

        var reader = new FileReader();
        reader.readAsDataURL(files[0]);
        showTable();
        reader.onload = function () {
            var fileUpload = jQuery("#file_name_" + id);
            var file = fileUpload[0].files[0];
            if (file != undefined) {
                jQuery("#show_file_name_" + id).text(file.name);
                $("#divshow_file_name_" + id).css("display", "flex");
                $("#divselect_file_name_" + id).css("display", "none");
                arraycheck.push(id);

                var formData = new FormData();
                var dataarray = [];
                var arraytype = [];
                var tmpfiletitle = [];

                jQuery("[name='file_name_doc']").each(function (index) {
                    dataarray.push(id);

                    var fileUpload = jQuery("#" + this.id);

                    var file = fileUpload[0].files[0];
                    if (file != undefined) {
                        var columns = this.id;

                        var datatype = columns.split("_");
                        datatype = datatype[2];
                        var filenamere = file.name.replace("'", "_").replace(",", "");
                        var file_name_form_type_doc_id = $(
                            "#form_type_doc_id_" + this.id
                        ).val();

                        tmpfiletitle.push(filenamere);
                        arraytype.push(file_name_form_type_doc_id);
                        formData.append("docfile", file);
                    }
                });

                formData.append("demand_letter_id", localStorage.getItem("demand_id"));
                formData.append("user_id", core_datas.default_profile.user_id);

                formData.append("document_title", tmpfiletitle.toString());
                formData.append("form_type_doc_id", arraytype.toString());
                formData.append(
                    "relate_account_id",
                    core_datas.current_profile.profile_id
                );

                $.ajax({
                    method: "post",
                    processData: false,
                    contentType: false,
                    cache: false,
                    data: formData,
                    url: url_SubmitDocFileDemandLetterOrNameList,
                    success: function (response) {
                        var dataresponse = JSON.parse(response);
                        if (dataresponse.status == "success") {
                            jQuery("[name='file_name_doc']").val("");
                            jQuery("#show_file_name_" + id).text(file.name);
                            $("#divshow_file_name_" + id).css("display", "flex");
                            $("#divselect_file_name_" + id).css("display", "none");
                            $("#uploadfile_file_name_" + id).val(id);
                            location.reload();

                        } else {
                            $("#loading").delay().fadeOut("");
                            $("#success").removeClass("active");

                            if (dataresponse.status == "filefail") {

                                Swal.fire({
                                    imageUrl: important_data.warningIconUrl,
                                    title: `<span class="lang_invalid_file_type">ชนิดไฟล์เอกสารไม่ถูกต้อง</span>`,
                                    html: `<span class="lang_uploaded_invalid">เนื่องจากชนิดของไฟล์เอกสาร์ที่ท่านอัปโหลดไม่ถูกต้อง กรุณาแก้ไขเอกสารในหน้าการแนบเอกสาร และดำเนินการส่งคำขอใหม่อีกครั้ง</span>`,
                                    confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                                    didRender: function () {
                                        changeDatalang();
                                    },
                                })
                            } else {
                                Swal.fire({
                                    imageUrl: important_data.errorIconUrl,
                                    title:
                                        '<span data_th="เพิ่มไฟล์ไม่สำเร็จ" data_en="File addition failed"></span>',
                                    confirmButtonText: "<span data_th='ตกลง' data_en='OK'></span>",
                                    showCancelButton: false,
                                    reverseButtons: true,
                                }).then((result) => {

                                });
                            }


                            changeDatalang();
                        }
                    },
                    error: function (xhr) {
                        $("#loading").delay().fadeOut("");
                        $("#success").removeClass("active");
                        if (xhr.status === 403) {
                            AlertCF(() => location.reload());

                        } else {
                            Swal.fire({
                                imageUrl: important_data.errorIconUrl,
                                title: `<span class="lang_failed_request"></span>`,
                                html: `<span class="">ไม่สามารถดำเนินการต่อได้ในขณะนี้กรุณาลองดำเนินการใหม่อีกครั้ง</span>`,
                                confirmButtonText: "<span class='lang_lbl_ok'>ตกลง</span>",
                                didRender: () => {
                                    changeDatalang();
                                }
                            });
                        }
                    }
                });
            }
        };
    }
    changeDatalang();
};

function showTable() {
    if (arraycheck.length == lengthfile.length) {
        $("#head_step_2").show();
        $("#Alien_Table_None").css("display", "block");
    }
}


async function GetDataFormNameListNotAppointment() {
    $("#loading").css("display", "block");

    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
        };
    }
    var request = await jQuery.post(
        url_GetDataFormNameListNotAppointment,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;

                if (objData.length != "0") {
                    datatable_Appointment = $(`#datatable_Appointment`).DataTable({
                        paging: true,
                        lengthChange: false,
                        searching: false,
                        ordering: true,
                        destroy: true,
                        columns: [
                            { data: "form_id" },
                            { data: "alien_id_card", className: "text-left" },
                            { data: "alien_name_th", className: "text-left" },
                            { data: "alien_nationality_th", className: "text-left" },
                            { data: "alien_date_of_birth", className: "text-left" },
                            { data: "stay_permis_no", className: "text-left" },
                        ],
                        columnDefs: [
                            {
                                targets: 0,
                                render: function (data, type, row, meta) {
                                    return `${meta.row + meta.settings._iDisplayStart + 1}`;
                                },
                            },
                        ],

                        data: objData,
                    });
                    $("#loading").delay().fadeOut("");
                }
            }
        },
        "json"
    );
}
async function GetSelectRequestFormAlien() {
    var request = await jQuery.post(
        url_GetSelectRequestForm,
        null,
        function (data) {
            if (data.status == "success") {
                $("#loading").delay().fadeOut("");
                var objData = data.data;
                var setData = "";
                var objData = objData.alien_doclist;
                console.log("2")
                setData += "<div class='col-12 col-md-12'>";
                setData +=
                    "<h3 class='mb-4 mt-3 head-step lang_doc_evidence_1'>เอกสารเเละหลักฐาน ครั้งที่ 1</h3>";
                setData += "</div>";
                for (let c = 0; c < objData.length; c++) {
                    const requestDocList = objData[c].requestdoclist;
                    setData += "<div class='col-md-12'>";
                    if (requestDocList.length >= 2) {
                        setData += `<p class='text-upload'>${c + 1
                            }${"."}<span class='lang_select_at_least_one_file'></span>`;
                        if (objData[c].required == "Y") {
                            setData += "<span class='asterisk'>*</span>";
                            setData +=
                                "<span id='err_file_name_" +
                                objData[c].group_form_type_doc_id +
                                "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                        }
                    }

                    for (let j = 0; j < requestDocList.length; j++) {
                        if (requestDocList.length < 2) {
                            setData += "<div>";
                            setData += `<p class='text-upload'>${c + 1
                                }${"."}<span data_th='${requestDocList[
                                    j
                                ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                }'></span>`;
                            if (objData[c].required == "Y") {
                                setData += "<span class='asterisk'>*</span>";
                                setData +=
                                    "<span id='err_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                            }
                        } else {
                            setData += '<div style="margin-left:15px;">';

                            setData += `<p class='text-upload'><span data_th='${requestDocList[
                                j
                            ].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en
                                }'></span>`;
                        }

                        setData += "</p>";
                        setData += "</p>";
                        setData += "<div class='row'>";
                        setData += "<div class='col-lg-12'>";

                        if (requestDocList[j].document_type_id == "42") {
                            setData +=
                                "<p class='text-upload-subtext lang_support_image_file'>รองรับไฟล์ .PNG .JPG เเละ .JPEG ขนาดไม่เกิน 4 MB</p>";
                        } else {
                            setData +=
                                "<p class='text-upload-subtext lang_support_pdf'>รองรับไฟล์ .PDF ขนาดไม่เกิน 4 MB</p>";
                        }

                        if (requestDocList[j].document_type_id === "72") {
                            if (requestDocList[j].chkis_upload == "N") {
                                if (requestDocList[j].doclist.length > 0) {
                                    setData += "<div class='health-report'>";
                                    setData += `<span style='font-weight:600;color:#0A090A;'>เลขที่กรมธรรม :</span> <span>${requestDocList[j].doclist[0]?.insurance_no || "-"}</span><br>`;
                                    setData += `<span style='font-weight:600;color:#0A090A;'>หน่วยงานที่ทําประกันสุขภาพ :</span> <span>${requestDocList[j].doclist[0]?.insurance_from || "-"}</span><br>`;
                                    setData += `<span style='font-weight:600;color:#0A090A;'>สถานะทําประกันสุขภาพ :</span> <span style="color:${requestDocList[j].doclist[0]?.insurance_status_color_code}">${requestDocList[j].doclist[0]?.insurance_status || "-"}</span> (วันที่ซื้อประกัน <span data_th='${convert_to_thai_date(requestDocList[j].doclist[0]?.insurance_date_buy) || "-"}' data_en='${convert_to_eng_date(requestDocList[j].doclist[0]?.insurance_date_buy) || "-"}'></span>)<br>`;
                                    setData += `<span style='font-weight:600;color:#0A090A;'>วันทีเริ่มต้นประกัน :</span> <span data_th='${convert_to_thai_date(requestDocList[j].doclist[0]?.insurance_date_start) || "-"}' data_en='${convert_to_eng_date(requestDocList[j].doclist[0]?.insurance_date_start) || "-"}'></span><br>`;
                                    setData += `<span style='font-weight:600;color:#0A090A;'>วันทีสิ้นสุดประกัน :</span> <span data_th='${convert_to_thai_date(requestDocList[j].doclist[0]?.insurance_date_end) || "-"}' data_en='${convert_to_eng_date(requestDocList[j].doclist[0]?.insurance_date_end) || "-"}'></span>`;
                                    setData += "</div>";
                                } else {

                                }
                            }

                        }

                        if (requestDocList[j].chkis_upload !== "N") {
                            setData += "<div class='row mb-3' style='margin-left:0px;'>";
                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList[j].doclist[0].doc_status_id == "C") {
                                    setData +=
                                        "<div class='upload d-none' style='padding-right:14px;'>";
                                } else {
                                    setData += "<div class='upload' style='padding-right:14px;'>";
                                }
                            } else {
                                setData += "<div class='upload' style='padding-right:14px;'>";
                            }

                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList.length >= 2) {
                                    if (requestDocList[j].document_type_id == "42") {
                                        setData +=
                                            "<input type='file' accept='image/png, image/jpeg, image/jpeg' name='file_name_group' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='SelectFileEdit(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;,&quot;" +
                                            requestDocList[j].doclist[0].document_id +
                                            "&quot;)' hidden/>";
                                    } else {
                                        setData +=
                                            "<input type='file' name='file_name_group' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='SelectFileEdit(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;,&quot;" +
                                            requestDocList[j].doclist[0].document_id +
                                            "&quot;)' hidden/>";
                                    }
                                } else {
                                    if (requestDocList[j].document_type_id == "42") {
                                        setData +=
                                            "<input type='file' accept='image/png, image/jpeg, image/jpeg' name='file_name' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='SelectFileEdit(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;,&quot;" +
                                            requestDocList[j].doclist[0].document_id +
                                            "&quot;)' hidden/>";
                                    } else {
                                        setData +=
                                            "<input type='file' name='file_name' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='SelectFileEdit(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;,&quot;" +
                                            requestDocList[j].doclist[0].document_id +
                                            "&quot;)' hidden/>";
                                    }
                                }
                            } else {
                                if (requestDocList.length >= 2) {
                                    if (requestDocList[j].document_type_id == "42") {
                                        setData +=
                                            "<input type='file' accept='image/png, image/jpeg, image/jpeg' name='file_name_group' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='SelectFileEdit(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;)' hidden/>";
                                    } else {
                                        setData +=
                                            "<input type='file' name='file_name_group' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='SelectFileEdit(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;)' hidden/>";
                                    }
                                } else {
                                    if (requestDocList[j].document_type_id == "42") {
                                        setData +=
                                            "<input type='file' accept='image/png, image/jpeg, image/jpeg' name='file_name' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='SelectFileEdit(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;)' hidden/>";
                                    } else {
                                        setData +=
                                            "<input type='file' name='file_name' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='SelectFileEdit(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;)' hidden/>";
                                    }
                                }
                            }

                            setData +=
                                "<label class='lang_lbl_selectfile' for='file_name_" +
                                requestDocList[j].document_type_id +
                                "' style='cursor:pointer;'>เลือกไฟล์</label>";
                            if (requestDocList[j].doclist.length != 0) {
                                setData +=
                                    "<input type='hidden' id='hidden_select_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' value=''>";
                            }
                            setData += "</div>";
                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                    setData +=
                                        "<div class='block-doc justify-content-between border_file_reject' id='divshow_file_name_" +
                                        requestDocList[j].document_type_id +
                                        "'>";
                                } else {
                                    setData +=
                                        "<div class='block-doc justify-content-between border_file_success' id='divshow_file_name_" +
                                        requestDocList[j].document_type_id +
                                        "'>";
                                }
                            } else {
                                setData +=
                                    "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' style='display:none;'>";
                            }
                            setData += "<div class='d-flex'>";
                            setData += "<div>";
                            setData += "<img src='/assets/images/document.svg'>";
                            setData += "</div>";
                            setData += "<div class='filename'>";
                            if (requestDocList[j].doclist.length != 0) {
                                setData +=
                                    "<p class='name-file namedoc' id='show_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "'>" +
                                    requestDocList[j].doclist[0].doc_title +
                                    "</p>";
                                setData +=
                                    '<input type="hidden" id="DocIDValue_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="' +
                                    requestDocList[j].document_id +
                                    '">';
                            }
                            setData +=
                                "<p class='name-file namedoc' id='show_file_name_" +
                                requestDocList[j].document_type_id +
                                "'></p>";

                            setData += "</div>";
                            setData += "<div>";

                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList[j].doclist[0].doc_status_id == "C") {
                                    setData +=
                                        "<img class='close d-none' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                        requestDocList[j].document_type_id +
                                        "&quot;)'>";
                                } else {
                                    setData +=
                                        "<img class='close' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                        requestDocList[j].document_type_id +
                                        "&quot;,&quot;" +
                                        requestDocList[j].doclist[0].document_id +
                                        "&quot;)'>";
                                }
                            } else {
                                setData +=
                                    "<img class='close' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)'>";
                            }

                            setData += "</div>";
                            setData += "</div>";
                            setData += "</div>";
                            setData += "</div>";
                            setData += "</div>";
                            setData += "</div>";
                            if (requestDocList[j].doclist.length != 0) {
                                if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                    setData +=
                                        "<div class='resson' id='resson_" +
                                        requestDocList[j].document_type_id +
                                        "'><p class='textresson lang_text_reason'>หมายเหตุ</p><p> " +
                                        requestDocList[j].doclist[0].doc_comment +
                                        " </p></div>";
                                }
                            }

                            setData += "</div>";

                        }

                        setData +=
                            '<input type="hidden" id="requiredvalue_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            objData[c].required +
                            '">';

                        setData +=
                            '<input type="hidden" id="hidden_form_type_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            requestDocList[j].form_type_doc_id +
                            '">';
                        setData +=
                            '<input type="hidden" id="required_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            objData[c].group_form_type_doc_id +
                            '">';

                        if (requestDocList[j].doclist.length != 0) {
                            setData +=
                                '<input type="hidden" id="have_file_current_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].doclist[0].document_id +
                                '">';
                            setData +=
                                '<input type="hidden" id="hidden_select_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="">';
                        } else {
                            setData +=
                                '<input type="hidden" id="have_file_current_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="N">';
                        }

                        if (requestDocList[j].document_type_id === "72") {
                            if (requestDocList[j].chkis_upload == "N") {
                                if (requestDocList[j].doclist.length == 0) {
                                    setData += `
                                  <div class="BorderEMP w-100 ml-4">
                                    <div class="d-flex align-items-center pl-2">
                                      <img src="/assets/images/warning.svg" class="pr-2" style="height:24px;">
                                      <span>
                                        เนื่องจากยังไม่มีการทำประกันสุขภาพ กรุณานำรหัสอ้างอิงคนต่างด้าว หรือเลขที่หนังสือเดินทาง ไปทำประกันสุขภาพกับบริษัทที่ได้รับอนุญาตจากกรมการจัดหางานก่อน จึงจะสามารถดำเนินการในขั้นตอนถัดไปได้
                                      </span>
                                    </div>
                                  </div>
                                `;
                                }
                            } else if (requestDocList[j].chkis_upload == "H") {
                                if (requestDocList[j].doclist.length == 0) {
                                    setData += `
                                  <div class="BorderEMP w-100 ml-4">
                                    <div class="d-flex align-items-center pl-2">
                                      <img src="/assets/images/warning.svg" class="pr-2" style="height:24px;">
                                      <span>
                                        ไม่พบข้อมูลประกันสุขภาพ กรุณาทำประกันผ่านบริษัทที่ได้รับอนุญาตโดยใช้รหัสอ้างอิงคนต่างด้าวหรือเลขหนังสือเดินทาง หากซื้อแล้วโปรดแนบเอกสาร
                                      </span>
                                    </div>
                                  </div>
                                `;
                                }
                            }
                        }


                    }

                    setData += "</div>";


                    setData += "<div class='col-md-12 mb-2 mt-2'>";
                    setData += "<hr>";
                    setData += "</div>";
                }
                jQuery("#uploadfilealien").html(setData);
                changeDatalang();
            } else {
                alert(data.msg);
            }
        },
        "json"
    );
}
async function OpenModalAddNameList(form, form_status, demand_letter_id) {
    statuscheck = form_status;
    currentfile = [];
    $("#AddNewNameList").attr("disabled", false);
    resetFormValidation("form_add_alien");
    if (form_status == "add") {
        CheckDataNameListQuotaCompleteadd();
        GetDataDemandLetterJobPermitCateList();
    } else {
        $("#AddNewNameList").addClass("d-none");
        $("#EditNameList").removeClass("d-none");
        GetDataDemandLetterJobPermitCateList(form, demand_letter_id, form_status);

    }
    await Promise.all([GetDataPrefixInfo(), GetDataStayPermissionTypeInfo()]);
}
function resetFormValidation(formId) {
    const form = document.getElementById(formId);
    if (form) {

        form.querySelectorAll(".form-control").forEach((input) => {
            input.classList.remove("error");
            const elements = document.querySelectorAll(".lang_please_input_info");
            $("#age_al-error").css("display", "none");
            elements.forEach((el) => {
                el.remove();
            });
        });


        form.reset();
    }
}

function CheckDataNameListQuotaCompleteadd() {
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var group_id = getParameterByName("group_id");
        objRequest = {
            group_id: group_id,
        };
    } else {
        let objRequest = {
            group_id: getParameterByName("group_id"),
        };
    }

    var request = jQuery.post(
        url_CheckDataNameListQuota,
        objRequest,
        function (data) {
            if (data.status == "success") {
                GetSelectRequestFormAlien();
                $("#EmployeeModal").modal("show");
                $("#AddNewNameList").removeClass("d-none");
                $("#EditNameList").addClass("d-none");
                $("#EmployeeModal").on("hidden.bs.modal", function () {
                    resetFormValidation("form_add_alien");
                });

                jQuery("[name='file_name']").each(function (index) {
                    $(`#divshow_` + this.id).css("display", "none");
                    $(`.upload`).css("display", "block");
                    $(`#divselect_` + this.id).css("display", "inline");
                    $(`#` + this.id).val("");
                    $(`#filename` + this.id).val("");
                    $(`#show_` + this.id).text("");
                    $(`#upload` + this.id).val("");
                });
                jQuery("[name='file_name_group']").each(function (index) {
                    $(`#divshow_` + this.id).css("display", "none");
                    $(`.upload`).css("display", "block");
                    $(`#divselect_` + this.id).css("display", "inline");
                    $(`#` + this.id).val("");
                    $(`#filename` + this.id).val("");
                    $(`#show_` + this.id).text("");
                    $(`#upload` + this.id).val("");
                });
            } else {
                Swal.fire({
                    imageUrl: important_data.warningIconUrl,
                    title:
                        '<span data_th="ไม่สามารถส่งคำขอได้" data_en="Unable to Submit the Request"></span>',
                    html:
                        '<span data_th="' +
                        data.msg +
                        '" data_en="' +
                        data.msgen +
                        '"></span>',
                    showCancelButton: false,
                    confirmButtonText: '<span data_th="ตกลง" data_en="OK"></span>',
                    reverseButtons: true,
                }).then((result) => {
                    $("#EmployeeModal").modal("hide");
                    location.reload();

                });
                changeDatalang();
            }
        },
        "json"
    );
    changeDatalang();
}
function CheckDataNameListQuotaComplete() {
    let objRequest = {
        group_id: getParameterByName("group_id"),
    };
    var request = jQuery.post(
        url_CheckDataNameListQuotaComplete,
        objRequest,
        function (data) {
            if (data.status == "success") {
                $("#loading").delay().fadeOut("");
                $("#tabShowTableEmployee").css("display", "none");
                $("#tabShowPayment").css("display", "block");
                $("#tabShowPaymentSuccess").css("display", "none");
                $("#Alien_Table").hide();
                $("#divpayment").hide();
            } else {
                $("#loading").delay().fadeOut("");
                Swal.fire({
                    showCancelButton: false,
                    imageUrl: important_data.warningIconUrl,
                    title: '<span class="lang_Notification">แจ้งเตือน</span>',
                    html:
                        '<span data_th="' +
                        data.msg +
                        '" data_en="' +
                        data.msgen +
                        '"></span>',
                    confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                    reverseButtons: true,
                }).then((result) => {


                });
                changeDatalang();
            }
        },
        "json"
    );
    changeDatalang();
}
async function SubmitDataAddNewRequestFormNameList() {
    const isTokenValid = await validate_token_presubmit();
    if (!isTokenValid) {
        return false;
    }
    $("#loading").css("display", "block");
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");

        objRequest = {
            user_id: userid,
            group_id: group_id,
            payment_uuid: $("#uuid").val(),
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            payment_uuid: $("#uuid").val(),
        };
    }
    var request = await jQuery.post(
        url_SubmitDataAddNewRequestFormNameList,
        objRequest,
        function (data) {
            if (data.status == "success") {
                GetSubmitDataRequestFormNameListSuccess();
                GetDataPaymentInquiryTransactionMOU(getParameterByName("group_id"));
            } else {
                $("#loading").delay().fadeOut("");

                Swal.fire({
                    imageUrl: important_data.warningIconUrl,
                    title: '<span class="lang_Notification">แจ้งเตือน</span>',
                    html: `<span>${data.msg}</span>`,
                    confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                    showCancelButton: false,
                    reverseButtons: true,
                });
                changeDatalang();
            }
        },
        "json"
    );
    await changeDatalang();
}
async function GetSubmitDataRequestFormNameListSuccess() {
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");

        objRequest = {
            user_id: userid,
            group_id: group_id,
        };
    } else {
        var userid = core_datas.default_profile.user_id;
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
        };
    }
    var request = await jQuery.post(
        url_GetSubmitDataRequestFormNameListSuccess,
        objRequest,
        function (data) {
            if (data.status == "success") {
                $("#loading").delay().fadeOut("");
                var objData = data.data;
                var status = sessionStorage.getItem("group_status");
                if (
                    status == "NCOS2" ||
                    status == "NCOC2" ||
                    status == "NCOCR2" ||
                    status == "NCOCRD2"
                ) {
                    var urlParams = new URLSearchParams(window.location.search);
                    if (
                        (urlParams.has("status") && urlParams.get("status") === "NCOS2") ||
                        (urlParams.has("status") && urlParams.get("status") === "NCOC2")
                    ) {
                        urlParams.set("status", "WCOSNA2");
                        window.location.href =
                            window.location.pathname + "?" + urlParams.toString();
                    }
                } else {
                    $("#loading").delay().fadeOut("");
                    var status_data = statusGroup;
                    if (status_data == "WA") {
                        $("#tabShowPaymentSuccess").show();
                        $("#tabShowPayment").hide();
                        $("#tabShowTableEmployee").hide();
                        $("#Alien_Table").hide();
                        $("#createdtimestamp_success").attr(
                            "data_th",
                            convert_to_thai_datetime(objData.created_timestamp)
                        );
                        $("#createdtimestamp_success").attr(
                            "data_en",
                            convert_to_eng_datetime(objData.created_timestamp)
                        );

                        $("#form_id_success").text(objData.group_id);
                    }
                }
            }
        },
        "json"
    );
    await changeDatalang();
}


function OpenModalPaymentStepOne(obj) {
    var objData = JSON.parse(obj.getAttribute("_payment-obj"));
    var objpayment = new Object();

    objpayment.request_unique_no = objData.payment_uuid;
    objpayment.payment_request_no = objData.payment_request_no;
    objpayment.payment_amount = objData.payment_price;
    objpayment.payment_channel = objData.ref3;
    reqChanel = objData.ref_id;

    transaction_id = objData.ref_transaction_id;
    $("#image_qrcode").empty();
    $("#Modal_selectPay").modal("show");
    $("#ref1_detail").text(objData.ref1);
    $("#ref2_detail").text(objData.ref2);
    $("#payment_type_name_th_detail").attr(
        "data_th",
        objData.payment_type_name_th
    );
    $("#payment_type_name_th_detail").attr(
        "data_en",
        objData.payment_type_name_en
    );
    $("#payment_type_name_th_st1").attr("data_th", objData.payment_type_name_th);
    $("#payment_type_name_th_st1").attr("data_en", objData.payment_type_name_en);
    $("#ref3_detail").text(objData.ref3);
    $("#payment_price_detail").text(objData.payment_price);

    $("#transaction_amount_st1").text(objData.payment_price);

    $("#ref3_st2").text(objData.ref3);
    $("#ref1_st2").text(objData.ref1);
    $("#ref2_st2").text(objData.ref2);
    $("#payment_price_st2").text(objData.payment_price);

    $("#payment_expired_time_st2").text(
        convert_to_thai_datetime(objData.payment_expired_time)
    );

    var setdata = "";
    var setdataOut = "";
    if (objData.payment_list.length != 0) {
        for (var i = 0; i < objData.payment_list.length; i++) {
            var payment_list = objData.payment_list[i];

            if (objData.payment_list[i].payment_type_id == "PAYMENTFEE") {
                if (objData.is_payment_fee == "IN") {
                    setdata +=
                        '<div class="col-6 text-left" data_th="' +
                        payment_list.payment_type_name_th?.toSafeJsonHtml() +
                        '" data_en="' +
                        payment_list.payment_type_name_en +
                        '"></div>';
                    setdata +=
                        '<div class="col-6 text-right">' +
                        payment_list.payment_price +
                        ' <span class="lang_lbl_currency"></span></div>';
                } else if (objData.is_payment_fee == "OUT") {
                    setdataOut =
                        '<div class="col-12 text-left" data_th="' +
                        payment_list.payment_type_name_th?.toSafeJsonHtml() +
                        '" data_en="' +
                        payment_list.payment_type_name_en +
                        '"></div>';
                }
            } else {
                setdata +=
                    '<div class="col-6 text-left" data_th="' +
                    payment_list.payment_type_name_th?.toSafeJsonHtml() +
                    '" data_en="' +
                    payment_list.payment_type_name_en +
                    '"></div>';
                setdata +=
                    '<div class="col-6 text-right">' +
                    payment_list.payment_price +
                    ' <span class="lang_lbl_currency"></span></div>';
            }

            if (objData.payment_list[i].payment_type_id == "PAYMENTFEE") {
                if (objData.is_payment_fee == "OUT") {
                    $("#payment_fee_text_detail").attr(
                        "data_th",
                        objData.payment_list[i].payment_type_name_th
                    );
                    $("#payment_fee_text_detail").attr(
                        "data_en",
                        objData.payment_list[i].payment_type_name_en
                    );
                }
            }
        }
        setdata += '<div class="col-12"><hr></div>';
        setdata += '<div class="col-6 text-left lang_lbl_total"> </div>';
        setdata +=
            '<div class="col-6 text-right">' +
            objData.payment_price +
            ' <span class="lang_lbl_currency"></span></div>';

        $("#show_payment_etracking").html(setdata);
        $("#payment_text_fee_out_detail").html(setdataOut);
    }

    const base64String = "" + objData.qrcode_image + "";

    function base64ToBlob(base64, mime) {
        const byteChars = atob(base64);
        const byteNumbers = new Array(byteChars.length);
        for (let i = 0; i < byteChars.length; i++) {
            byteNumbers[i] = byteChars.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        return new Blob([byteArray], { type: mime });
    }


    const mimeType = "image/png";
    const blob = base64ToBlob(base64String, mimeType);
    const url = URL.createObjectURL(blob);
    const img = document.createElement("img");
    img.setAttribute("width", "30%");
    img.src = url;
    document.getElementById("image_qrcode").appendChild(img);

    changeDatalang();
}
function Openproofpayment(transaction_id_value) {
    if (transaction_id_value == "" || transaction_id_value == undefined) {
        transaction_id_value = transaction_id;
    }
    let objRequest = {
        group_id: getParameterByName("group_id"),
        user_id: core_datas.default_profile.user_id,
        transaction_id: transaction_id_value,
        relate_account_id: core_datas.current_profile.profile_id,
    };
    var request = jQuery.post(
        url_GetDocFilePaymentReceipt,
        objRequest,
        function (data) {
            if (data.status == "success") {
                $("#WP_Payment").modal("hide");
                var objData = data.data.payment_document;

                if (important_data.is_mobile == "Y") {
                    mobileScriptChannel.postMessage("openFile:" + objData.document_url);
                } else {
                    window.open(objData.document_url, "_blank");
                }
            } else {
                Swal.fire({
                    imageUrl: errorIconUrl,
                    title: `<span class="lang_Notification">เกิดข้อผิดพลาด</span>`,
                    html: "<span>" + data.msg + "</span>",
                    confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                    reverseButtons: true,
                });
                changeDatalang();
            }
        },
        "json"
    );
    changeDatalang();
}


function backtoPaymentModal() {
    $("#DetailPayment").modal("hide");
    $("#Modal_selectPay").modal("show");
}
function OpentWP_Payment() {
    $("#Modal_selectPay").modal("hide");
    $("#DetailPayment").modal("show");
    sendPaymentStatusMOU(reqChanel);
}


async function UpdateDataPaymentSuccessRequestForm() {
    const isTokenValid = await validate_token_presubmit();
    if (!isTokenValid) {
        return false;
    }
    let objRequest = {
        group_id: getParameterByName("group_id"),
        user_id: core_datas.default_profile.user_id,
        transaction_id: transaction_id,
        relate_account_id: core_datas.current_profile.profile_id,
    };
    var request = jQuery.post(
        url_GetDocFilePaymentInvoice,
        objRequest,
        function (data) {
            if (data.status == "success") {
                $("#DetailPayment").modal("hide");
                $("#Modal_selectPay").modal("hide");
                $(".modal-backdrop").addClass("d-none");

                var objData = data.data.payment_document;

                if (important_data.is_mobile == "Y") {
                    sendmessage(objData.document_url);
                } else {
                    window.open(objData.document_url, "_blank");
                }
            } else {
                Swal.fire({
                    imageUrl: errorIconUrl,
                    title: `<span class="lang_Notification">เกิดข้อผิดพลาด</span>`,
                    html: "<span>" + data.msg + "</span>",
                    confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                });
                changeDatalang();
            }
        },
        "json"
    );
    changeDatalang();
}
var connection;
var chkPay = false;
function sendPaymentStatusMOU(reqChanel) {
    if (reqChanel != null) {
        $("#img_payment").empty();
        startConnectionMOU(reqChanel);
    }
}

var _mouSignalRRetry = {};
var MOU_SIGNALR_MAX_RETRY = 5;
function startConnectionMOU(reqChanel) {
    if (!reqChanel.startsWith(sendMsg)) {
        reqChanel = sendMsg + reqChanel;
    }

    connection = new signalR.HubConnectionBuilder()
        .withUrl(important_data.hubUrl, {
            transport: signalR.HttpTransportType.ServerSentEvents,
        })
        .withAutomaticReconnect()
        .build();

    connection
        .start()
        .then(() => {
            console.log("SignalR connected");
            _mouSignalRRetry[reqChanel] = 0; // connect สำเร็จ → reset ตัวนับ
            if (reqChanel) {
                joinGroupMOU(reqChanel);
            }
        })
        .catch((err) => {
            console.log("SignalR connection error:", err);
            var n = (_mouSignalRRetry[reqChanel] || 0) + 1;
            _mouSignalRRetry[reqChanel] = n;
            if (n > MOU_SIGNALR_MAX_RETRY) {
                console.warn("[MOU] หยุด retry SignalR หลังพยายาม " + MOU_SIGNALR_MAX_RETRY + " ครั้ง");
                return;
            }
          
            setTimeout(function () { startConnectionMOU(reqChanel); }, 250 * n); // Retry (backoff)
        });

    connection.onclose(() => {
        console.log("onclose");
    });

    connection.onreconnecting(() => {
        console.log("reconncect");
    });

    connection.onreconnected(() => {
        if (reqChanel) {
            joinGroupMOU(reqChanel);
        }
    });

    connection.on("ReceiveMessage", function (message) {
        var split_status = message.split("|");
        if (split_status[0] == "success") {
            $("#DetailPayment").modal("hide");
            $("#text_status_payment").addClass("lang_lbl_payment_complete");
            const img = document.createElement("img");
            img.src = "/assets/images/success.jpg";
            img.className = "successimg";

            $("#img_payment").html(img);
            $("#success_payment").modal("show");
            leaveGroupMOU(reqChanel);
        } else if (split_status[0] == "fail") {
            $("#WP_Payment").modal("hide");
            $("#text_status_payment").html(
                "<span class='lang_lbl_payment_fail_retry'>ชำระเงินไม่สำเร็จ กรุณาลองใหม่อีกครั้ง</span>"
            );

            const img = document.createElement("img");
            img.src = "/assets/images/picmobile/invalid.png";
            img.className = "successimg";
            $("#img_payment").html(img);
            $("#success_payment").modal("show");
            $("#btn_openpayment").hide();
            $(".proofpayment").addClass("d-none");
        } else if (split_status[0] == "expired") {
            $("#WP_Payment").modal("hide");
            $("#text_status_payment").html(
                "<span class='lang_lbl_payment_fail'>ชำระเงินไม่สำเร็จ</span>"
            );
            const img = document.createElement("img");
            img.src = "/assets/images/picmobile/invalid.png";
            img.className = "successimg";
            $("#img_payment").html(img);
            $("#success_payment").modal("show");
            $("#btn_openpayment").hide();

            $(".expDiv").removeClass("d-none");
            $(".proofpayment").addClass("d-none");

            $("#exp_label").addClass("lang_exp_description");
        }
        changeDatalang();
    });
}
function reloadQueueMOU() {
    var group_req = getParameterByName("group_id");
    if (group_req != null) {
        startConnectionQueueMOU(group_req);
    }
}
function startConnectionQueueMOU(queueChanel) {
    if (!queueChanel.startsWith(sendMsgQ)) {
        queueChanel = sendMsgQ + queueChanel;
    }
    connection = new signalR.HubConnectionBuilder()
        .withUrl(important_data.hubUrlqueue, {
            transport: signalR.HttpTransportType.ServerSentEvents,
        })
        .withAutomaticReconnect()
        .build();

    connection
        .start()
        .then(() => {
            console.log("SignalR connected");
            _mouSignalRRetry[queueChanel] = 0; // connect สำเร็จ → reset ตัวนับ
            if (queueChanel) {
                joinGroupMOU(queueChanel);
            }
        })
        .catch((err) => {
            console.log("SignalR connection error:", err);
            var n = (_mouSignalRRetry[queueChanel] || 0) + 1;
            _mouSignalRRetry[queueChanel] = n;
            if (n > MOU_SIGNALR_MAX_RETRY) {
                console.warn("[MOU] หยุด retry SignalR (queue) หลังพยายาม " + MOU_SIGNALR_MAX_RETRY + " ครั้ง");
                return;
            }
           
            setTimeout(function () { startConnectionQueueMOU(queueChanel); }, 250 * n); // Retry (backoff)
        });

    connection.onclose(() => {
        console.log("onclose");
    });

    connection.onreconnecting(() => {
        console.log("reconncect");
    });

    connection.onreconnected(() => {
        if (queueChanel) {
            joinGroupMOU(queueChanel);
        }
    });

    connection.on("QueueMessage", function (message) {
        console.log(message);
        window.location.reload();
    });
}

function closeConnection() {
    connection
        .stop()
        .then(() => {
            console.log("Connection to paymentHub closed.");
        })
        .catch((err) => {
            console.error("Error closing connection: ", err);
        });
}
function joinGroupMOU(reqChanel) {

    if (connection) {
        connection
            .invoke("JoinGroup", reqChanel)
            .catch((err) => console.error("Error joining group:", err));
    }
}
function leaveGroupMOU(reqChanel) {
    if (connection) {
        connection
            .invoke("LeaveGroup", reqChanel)
            .catch((err) => console.log("Error leaving group:", err));
    }
}
async function AppealList(group_id) {
    $("#DetailDocumentApprealList").empty();
    if (group_id != null) {
        $.ajax({
            type: "POST",
            url: url_AppealList,
            data: {
                group_id: group_id,
            },
            dataType: "json",
            success: function (data) {
                $("#divNoneDocAppreal").hide();
                $("#divHasDocAppreal").show();

                var solutionappreal = "";

                solutionappreal += `

                    <div col-12>
                         <h5 class="card-title lang_appeal_disapproval">การอุทธรณ์คำสั่งไม่อนุญาตให้คนต่างด้าวทำงาน</h5>
                    </div>
                    <div col="12">
                                    <p class="card-text lang_appeal_royal_decree" style="text-indent: 2em;">
                                        พระราชกำหนดการบริหารจัดการการทำงานของคนต่างด้าว พ.ศ. 2560 และที่แก้ไขเพิ่มเติม
                                    </p>
                                    <p class="card-text lang_appeal_description" style="text-indent: 2em;">
                                        มาตรา 95 ในกรณีที่นายทะเบียนไม่อนุญาตตามมาตรา 59 มาตรา 60 มาตรา 63 มาตรา 63/1 มาตรา 63/2 หรือมาตรา 64 หรือไม่ต่ออายุใบอนุญาตทำงานตามมาตรา 67 หรือเพิกถอนใบอนุญาตทำงานตามมาตรา 90 คนต่างด้าวซึ่งขออนุญาตทำงาน ผู้ซึ่งขออนุญาตแทนคนต่างด้าว หรือผู้รับอนุญาตให้ทำงาน แล้วแต่กรณี มีสิทธิอุทธรณ์ต่อรัฐมนตรีได้ภายในสามสิบวันนับแต่วันที่ได้รับแจ้งคำสั่งดังกล่าว โดยคำวินิจฉัยของรัฐมนตรีให้เป็นที่สุด
                                    </p>
                                </div>
                `;

                $(".solutionappreal").append(solutionappreal);
                if (data.isSuccess && data.data.length > 0) {
                    let nameDocAppealList = data.data;
                    let Appeal = "Appeal";
                    let num = 1;
                    DetailDocumentAppealList = "";
                    for (var i = 0; i < nameDocAppealList.length > 0; i++) {
                        if (i == 0) {
                            DetailDocumentAppealList += '<div class="row col-12">';

                            DetailDocumentAppealList +=
                                '<div class="col-1 text-center">' + num + "</div>";

                            DetailDocumentAppealList +=
                                '<div class="col-8 col-lg-10 mb-3"><span data_th="' +
                                nameDocAppealList[i].full_Name_Th +
                                '" data_en="' +
                                nameDocAppealList[i].full_Name_En +
                                '"></span>';

                            DetailDocumentAppealList += '<p class="text-size"></p>';


                            DetailDocumentAppealList +=
                                '<div class="col-12 collapse show" id="collapse' + i + '">';
                            DetailDocumentAppealList +=
                                '<span class="lang_request_not_allowed_remark"><b>หมายเหตุที่ไม่อนุญาตคำขอ</b></span>';
                            DetailDocumentAppealList +=
                                '<p><span class="lang_for_registrar"> (หมายเหตุจากนายทะเบียน)</span></p>';
                            DetailDocumentAppealList +=
                                ' <p class="card-text">' +
                                nameDocAppealList[i].approve_Reason +
                                " </p>";

                            DetailDocumentAppealList += "</div>";


                            DetailDocumentAppealList += " </div>";

                            DetailDocumentAppealList += '<div class="col-1">';

                            DetailDocumentAppealList +=
                                '<button type="button" id="bbtn' +
                                i +
                                '" class="btn dark btn-appreal"  data-toggle="collapse' +
                                i +
                                '"  data-target="#collapse' +
                                i +
                                '" aria-expanded="false" aria-controls="collapse' +
                                i +
                                '" onclick="toggleButtonTextAppreal(this)">';

                            DetailDocumentAppealList +=
                                '<span class="lang_lbl_hide" style="color: var(--pink-700);">ซ่อน</span>';


                            DetailDocumentAppealList += "</button>";

                            DetailDocumentAppealList += "</div>";

                            DetailDocumentAppealList += " </div>";

                            DetailDocumentAppealList +=
                                '<div class="col-1 text-center"></div>';


                            DetailDocumentAppealList += '<div class="row-6">';
                            DetailDocumentAppealList += '<div class="col-12">';
                            DetailDocumentAppealList +=
                                '<button type="button" class="btn btn-dark m-0 previous action-button text-left" id="btn_DisApprove" onclick="GetDocumentConfirm(\'' +
                                nameDocAppealList[i].form_Id +
                                "', '" +
                                nameDocAppealList[i].document_Id +
                                "', '" +
                                "appreal" +
                                '\');" ><span class="lang_print_appeal_document">พิมพ์เอกสารยื่นอุทธรณ์</span></button>';

                            DetailDocumentAppealList += `
                                ${renderActionButtonMOU(
                                    nameDocAppealList[i].form_Id,
                                    nameDocAppealList[i].document_Id,
                                    "/assets/images/refresh.png",
                                    nameDocAppealList[i].updateTimeStamp,
                                    "",
                                    true,
                                    group_id,
                                    nameDocAppealList[i].rowId,
                                    Appeal)
                                }

                                `;


                            DetailDocumentAppealList += "</div>";
                            DetailDocumentAppealList += "</div>";

                        } else {
                            DetailDocumentAppealList += '<div class="row col-12">';

                            DetailDocumentAppealList =
                                '<div class="col-1 text-center"></div>';

                            DetailDocumentAppealList += '<div class="col-8 col-lg-10 mb-3">';

                            DetailDocumentAppealList +=
                                '<span data_th="' +
                                nameDocAppealList[i].full_Name_Th +
                                '" data_en="' +
                                nameDocAppealList[i].full_Name_En +
                                '"></span>';

                            DetailDocumentAppealList += '<p class="text-size"></p>';


                            DetailDocumentAppealList +=
                                '<div class="col-11 collapse show" id="collapse' + i + '">';
                            DetailDocumentAppealList += `<div class='col-1'></div>`;
                            DetailDocumentAppealList +=
                                '<span class="lang_request_not_allowed_remark"><b>หมายเหตุที่ไม่อนุญาตคำขอ</b></span>';
                            DetailDocumentAppealList +=
                                '<p><span class="lang_for_registrar"> (หมายเหตุจากนายทะเบียน)</span></p>';
                            DetailDocumentAppealList +=
                                ' <p class="card-text">' +
                                nameDocAppealList[i].approve_Reason +
                                " </p>";

                            DetailDocumentAppealList += "</div>";


                            DetailDocumentAppealList += " </div>";

                            DetailDocumentAppealList += "</div>";

                            DetailDocumentAppealList += '<div class="col-1">';

                            DetailDocumentAppealList +=
                                '<button type="button" id="bbtn' +
                                i +
                                '" class="btn dark btn-appreal"  data-toggle="collapse' +
                                i +
                                '"  data-target="#collapse' +
                                i +
                                '" aria-expanded="false" aria-controls="collapse' +
                                i +
                                '" onclick="toggleButtonTextAppreal(this)">';


                            DetailDocumentAppealList +=
                                '<span class="lang_lbl_hide" style="color: var(--pink-700);">ซ่อน</span>';

                            DetailDocumentAppealList += "</button>";

                            DetailDocumentAppealList += "</div>";

                            DetailDocumentAppealList += "</div>";

                            DetailDocumentAppealList +=
                                '<div class="col-1 text-center"></div>';


                            DetailDocumentAppealList += '<div class="row-6">';
                            DetailDocumentAppealList += '<div class="col-12">';
                            DetailDocumentAppealList +=
                                '<button type="button" class="btn btn-dark m-0 previous action-button text-left" id="btn_DisApprove" onclick="GetDocumentConfirm(\'' +
                                nameDocAppealList[i].form_Id +
                                "', '" +
                                nameDocAppealList[i].document_Id +
                                '\');" ><span class="lang_print_appeal_document">พิมพ์เอกสารยื่นอุทธรณ์</span></button>';

                            DetailDocumentAppealList += `

                                ${
                                renderActionButtonMOU(
                                    nameDocAppealList[i].form_Id,
                                    nameDocAppealList[i].document_Id,
                                    "/assets/images/refresh.png",
                                    nameDocAppealList[i].updateTimeStamp,
                                    "",
                                    true,
                                    group_id,
                                    nameDocAppealList[i].rowId,
                                    Appeal)
                            }

                                `;


                            DetailDocumentAppealList += "</div>";
                            DetailDocumentAppealList += "</div>";


                            if (i > 0) {
                                let textSelectedFile = `<div class='col-1'></div>
                                          <div class='col-11 mb-4'></div>
                                          <div class='col-1 text-center'>${num}</div`;
                                DetailDocumentAppealList =
                                    DetailDocumentAppealList.slice(0, 0) +
                                    textSelectedFile +
                                    DetailDocumentAppealList.slice(0);
                            }
                        }
                        if (DetailDocumentAppealList != "") {
                            $("#DetailDocumentApprealList").append(DetailDocumentAppealList);
                        }

                        num++;
                        if (DetailDocumentAppealList != "") {
                            $("#DetailDocumentApprealList").append("<hr class='w-100' />");
                        }
                    }
                } else {
                    $("#divNoneDocAppreal").show();
                    $("#divHasDocAppreal").hide();
                }
            },
            error: function (error) {
                console.log(error);
            },
        });
    }
    changeDatalang();
}

function openActiveMOU(activeIndex, type) {
    const allTabs = [
        "tab_default_1",
        "tab_default_2",
        "tab_default_3",
        "tab_default_4",
        "tab_default_5",
        "tab_default_6",
        "tab_default_7",
    ];
    const allDivs = [
        "div_tab1",
        "div_tab2",
        "div_tab3",
        "div_tab4",
        "div_tab5",
        "div_tab6",
        "div_tab7",
    ];

    const deleteIndex = activeIndex - 1;

    if (type == "active") {
        allTabs.forEach((tab) => $("#" + tab).removeClass("active"));
        allDivs.forEach((div) => $("#" + div).removeClass("active"));

        if (activeIndex != null) {
            $("#" + allTabs[deleteIndex]).addClass("active");
            $("#" + allDivs[deleteIndex]).addClass("active");
        }
    } else {

        $("#" + allDivs[deleteIndex]).show();
    }
}


async function GetDetailDocumentListMOU(group_id) {
    $("#DetailDocumentList").empty();
    $("#loading").show();

    if (!group_id) return;

    try {
        const res = await $.ajax({
            type: "POST",
            url: "/ConfirmationDocument/GetDetailSectionMouList",
            data: { group_Id: group_id },
            dataType: "json"
        });

        if (!res.isSuccess || !res.data || res.data.length === 0) {
            $("#divNoneDoc").show();
            $("#divHasDoc").hide();
            return;
        }

        $("#divNoneDoc").hide();
        $("#divHasDoc").show();

        let num = 1;

        res.data.forEach(doc => {
            $("#DetailDocumentList").append(renderMainDoc(doc, num, group_id));
            $("#DetailDocumentList").append("<hr class='w-100' />");

            num++;
        });
        const noticeHtml = `
                      <div style="font-size:15px">
                            <div>**เมื่อกดสร้างเอกสารแล้ว ระบบจะกำหนดให้สามารถกดสร้างเอกสารใหม่ได้อีกครั้งหลังจากครบ 24 ชั่วโมง</div>
                            <div style="color:#FF3B30;">**หลังจากรีเฟรชเอกสารเเล้ว กรุณารอสักครู่ เอกสารจะถูกอัปเดตภายใน 1-2 นาที</div>
                        </div>
                    `;

        $("#DetailDocumentList").append(noticeHtml);
        changeDatalang();
        $("#loading").css("display", "none");
    } catch (err) {
        $("#loading").css("display", "none");
        console.error(err);
    }
}


function renderMainDoc(doc, num, group_id) {
    let html = `
        <div class="row align-items-center col-12">
            <div class="col-1 text-center">${num}</div>
            <div class="col-8 col-lg-8">
                <span data_th="${doc.nameFile_Th?.toSafeJsonHtml()}" data_en=""></span>
                <span>${doc.employerName ?? ""}</span>
                <p class="text-size"></p>
            </div>
    `;

    if (doc.nameLists && doc.nameLists.length > 0) {
        html += renderSubDocs(doc, num, group_id);
    } else {
        html += renderActionArea(doc, group_id);
    }

    html += "</div>";

    return html;
}


function renderSubDocs(doc, num, group_id) {
    let html = "";
    let subNum = 1;

    doc.nameLists.forEach(item => {
        html += `
            <hr class="w-100"/>
            <div class="col-1"></div>
            <div class="col-1 text-center">${num}.${subNum}</div>
            <div class="col-7">
                <span data_th="${item.nameListFile_Th?.toSafeJsonHtml()}" data_en=""></span>
                ${renderStatusBadge(item.nameList_Status)}
                <p class="text-size"></p>
            </div>
            ${renderActionArea(item, group_id)}


        `;
        subNum++;
    });

    return html;
}

function renderStatusBadge(status) {
    if (status === "CCOS" || status === "CCOS2") {
        openActiveMOU("7", null);
        return `<span class="badge badge-pill statusCC lang_not_allow">ไม่อนุญาต</span>`;
    }

    if (status === "CC") {
        return `<span class="badge badge-pill statusCC lang_cancel_by_emp">ยกเลิกจากนายจ้าง</span>`;
    }

    return "";
}

function renderActionArea(item, group_id) {
    let html = "";
    const nextTime = addOneHourAndFormatTH(item.updateTimeStamp);


    const canShowRefresh = !!item.updateTimeStamp;


    html += renderActionButtonMOU(
        item.form_Id,
        item.document_Id,
        "/assets/images/document.svg",

    );


    if (canShowRefresh) {
        html += renderActionButtonMOU(
            item.form_Id,
            item.document_Id,
            "/assets/images/refresh.png",
            item.updateTimeStamp,
            "",
            true,
            group_id,
            item.rowId
        );

        html += `${nextTime.canShow ? `<span style="font-size:13px; display:block;"> สามารถกดได้อีกครั้ง ${nextTime.text}</span>` : ""}`;


    }


    return `
        <div class="row col-3 justify-content-end text-end">
            ${html}
        </div>

    `;
}

function renderActionButtonMOU(
    formId,
    docId,
    imgSrc,
    updateTimeStamp = "",
    extraStyle = "",
    reload = false,
    group_id,
    rowId,
    Appeal
) {
    const isMobile = important_data.is_mobile?.toLowerCase() === "y";
    const linkAttr = isMobile
        ? `target="_blank"`
        : `href="javascript:void(0);"`;


    if (reload) {
        if (docId) {
            return `
            <a href="javascript:void(0);" type="button" class="btn action-button" title="สร้างเอกสารใหม่"
               onclick="handleUpdateClickMOU('${docId}', '${updateTimeStamp}','${group_id}','','${Appeal}')">
                <img src="${imgSrc}"
                     style="background:white;border-radius:6px;
                            padding:6px;border:1px solid #F6C8D9;
                            width:40px;${extraStyle}">
            </a>
        `;
        } else {
            return `
            <a href="javascript:void(0);" type="button" class="btn action-button" title="สร้างเอกสารใหม่"
               onclick="handleUpdateClickMOU('${rowId}', '${updateTimeStamp}','${group_id}','GenerateDocument','${Appeal}')">
                <img src="${imgSrc}"
                     style="background:white;border-radius:6px;
                            padding:6px;border:1px solid #F6C8D9;
                            width:40px;${extraStyle}">
            </a>
        `;
        }

    }


    return `
        <a target="_blank" class="btn action-button"
           title="เปิดเอกสาร"
           onclick="GetDocumentConfirm('${formId}','${docId}')">
            <img src="${imgSrc}"
                 style="background:white;border-radius:6px;padding:6px;border:1px solid #F6C8D9;">
        </a>
    `;
}
function handleUpdateClickMOU(docId, updateTimeStamp, group_id, GenerateDocument, Appeal) {

    if (!canClickAgain(updateTimeStamp)) {
        const nextTimeText = addOneHourAndFormatTH(updateTimeStamp);

        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            html: `<span><b>เเจ้งเตือน</b></span><br><br><span>สามารถกดได้อีกครั้ง<br><b>${nextTimeText.text}</b></span>`,
            confirmButtonText: "ปิด"
        });

        return false;
    }

    if (GenerateDocument) {
        GenerateDocumentByUserMOU(docId, group_id, Appeal);
    } else {

        UpdateDocumentByUserMOU(docId, group_id, Appeal);
    }

}
async function UpdateDocumentByUserMOU(document_Id, group_id, Appeal) {
    const isTokenValid = await validate_token_presubmit();
    if (!isTokenValid) {
        return false;
    }
    $("#loading").css("display", "block");
    await $.ajax({
        type: "POST",
        url: "/ConfirmationDocument/UpdateDocumentByUser",
        data: {
            docId: document_Id,
        },
        dataType: "json",
        success: function (data) {
            console.log(data);
            if (data.isSuccess) {
                $("#loading").css("display", "none");
                if (Appeal) {
                    window.location.reload();
                    AppealList(group_id)
                } else {
                    window.location.reload();
                    GetDetailDocumentListMOU(group_id)
                }

            } else {
                $("#loading").css("display", "none");
                Swal.fire({
                    imageUrl: important_data.warningIconUrl,
                    title: '<span class="">เเจ้งเตือน</span>',
                    html: '<span class="">"' + data.msg + '"</span> ',
                    confirmButtonText: '<span class="">ปิด</span>',
                    reverseButtons: true,
                    allowOutsideClick: false,
                    allowEscapeKey: false
                }).then((result) => {


                })
                changeDatalang();
            }
        },
    });


}
async function GenerateDocumentByUserMOU(rowId, group_id, Appeal) {
    const isTokenValid = await validate_token_presubmit();
    if (!isTokenValid) {
        return false;
    }
    $("#loading").css("display", "block");
    await $.ajax({
        type: "POST",
        url: "/ConfirmationDocument/GenerateDocumentByUser",
        data: {
            rowId: rowId,
        },
        dataType: "json",
        success: function (data) {
            console.log(data);
            if (data.isSuccess) {
                $("#loading").css("display", "none");
                if (Appeal) {
                    window.location.reload();
                    AppealList(group_id)
                } else {
                    window.location.reload();
                    GetDetailDocumentListMOU(group_id)
                }

            } else {
                $("#loading").css("display", "none");
                Swal.fire({
                    imageUrl: important_data.warningIconUrl,
                    title: '<span class="">เเจ้งเตือน</span>',
                    html: '<span class="">"' + data.msg + '"</span> ',
                    confirmButtonText: '<span class="">ปิด</span>',
                    reverseButtons: true,
                    allowOutsideClick: false,
                    allowEscapeKey: false
                }).then((result) => {


                })
                changeDatalang();
            }
        },
    });


}


function addOneHourAndFormatTH(isoDate) {
    const lastTime = new Date(isoDate);
    const nextTime = new Date(lastTime);
    nextTime.setHours(nextTime.getHours() + 1);


    const now = new Date();

    return {
        canShow: now < nextTime,
        text: nextTime.toLocaleString("th-TH", {
            day: "numeric",
            month: "short",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit"
        })
    };
}


async function GetDocumentConfirm(id, object_key) {
    show_loading();

    await $.ajax({
        type: "POST",
        url: "/ConfirmationDocument/GetDocument",
        data: {
            id: id,
            object_key: object_key,
        },
        dataType: "json",
        success: function (data) {
            if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
                sendmessage(data.data);
                $("#loading").hide();
            } else {
                window.open(data.data, "_blank");
                $("#loading").hide();
            }
        },
    });
}

function toggleButtonTextAppreal(button) {
    const targetId = $(button).data("target");
    const span = $(button).find("span");

    $(targetId).on("shown.bs.collapse", function (e) {
        e.preventDefault();
        span.removeClass("lang_lbl_hide").addClass("lang_lbl_expand");
    });

    $(targetId).on("hidden.bs.collapse", function (e) {
        e.preventDefault();

        span.removeClass("lang_lbl_expand").addClass("lang_lbl_hide");
    });
    $(targetId).collapse("toggle");
    changeDatalang();
}
async function AppealTopdfReuse(group_Id, form_Id, form_Seq) {
    if (group_Id != null) {
        show_loading();
        $.ajax({
            type: "POST",
            url: "ConfirmationDocument/GenerateTemplateAppeal",
            data: {
                group_Id: group_Id,
                form_Id: form_Id,
                form_Seq: form_Seq,
            },

            success: function (data, textStatus, jqXHR) {
                var response = JSON.parse(data);

                renderFileContentToPdf(response.data);


                $("#loading").hide();
            },
            error: function (error) {
                console.log(error);
                $("#loading").hide();
            },
        });
    }
}
var checkempfile;
var checkempfile_sale;
function checkValidateFileEmp() {
    if (checkempfile) {
        $("#nextstepcheckfileEditMOU").trigger("click");
    } else {
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: '<span class="lang_Employer_documents_incomplete">เอกสารนายจ้างไม่ครบ</span>',
            html: '<span class="lang_employer_attach_documents">กรุณาติดต่อนายจ้างเพื่อแนบเอกสารเพื่อใช้ในการยื่น</span>',
            showCancelButton: true,
            cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            didRender: function () {
                changeDatalang();
            },
        }).then((result) => {
            if (result.isConfirmed) {

            }
        });
        changeDatalang();
    }
    changeDatalang();
}
function checkValidateFileEmpRequest() {
    if (checkempfile & checkempfile_sale) {
        $("#nextstepcheckfileMOU").trigger("click");
    } else {
        if (checkempfile == false) {
            Swal.fire({
                imageUrl: important_data.warningIconUrl,
                title: '<span class="lang_Employer_documents_incomplete">เอกสารนายจ้างไม่ครบ</span>',
                html: '<span class="lang_employer_attach_documents">กรุณาติดต่อนายจ้างเพื่อแนบเอกสารเพื่อใช้ในการยื่น</span>',
                cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
                confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                reverseButtons: true,
            }).then((result) => {
                if (result.isConfirmed) {

                }
            });
            changeDatalang();
        } else if (checkempfile_sale == false) {
            Swal.fire({
                imageUrl: important_data.warningIconUrl,
                title: '<span class="lang_Employer_documents_incomplete">เอกสารนายจ้างไม่ครบ</span>',
                html: '<span class="lang_employer_attach_documents">กรุณาติดต่อนายจ้างเพื่อแนบเอกสารเพื่อใช้ในการยื่น</span>',
                showCancelButton: true,
                cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
                confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                reverseButtons: true,
            }).then((result) => {
                if (result.isConfirmed) {

                }
            });
            changeDatalang();
        } else {
            Swal.fire({
                imageUrl: important_data.warningIconUrl,
                title: '<span class="lang_Employer_documents_incomplete">เอกสารนายจ้างไม่ครบ</span>',
                html: '<span class="lang_employer_attach_documents">กรุณาติดต่อนายจ้างเพื่อแนบเอกสารเพื่อใช้ในการยื่น</span>',
                showCancelButton: true,
                cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
                confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                reverseButtons: true,
            }).then((result) => {
                if (result.isConfirmed) {

                }
            });
            changeDatalang();
        }
    }
    changeDatalang();
}
$("#nextstepcheckfileMOU").on("click", function () {
    var app = getParameterByName("App");
    var chkData = true;
    var EmployerID = $("#employer_id").val();


    jQuery("[name='file_name']").each(function (index) {

        var fileUpload = jQuery("#" + this.id);
        var ckrequired = jQuery("#requiredvalue_" + this.id).val();
        var ckhave_file_current = jQuery("#have_file_current_" + this.id).val();
        var doc_new_lastest = $("#doc_new_lastest_" + this.id).val();
        var document_type_id = $("#document_type_id_" + this.id).val();
        var fileUp = fileUpload[0].files[0];
        var chIDFile = fileUpload[0].id;
        var aTag = document.createElement("a");
        aTag.id = "file_link_" + chIDFile;
        aTag.innerHTML = "<img src='/assets/images/document.svg' class='iconbook'>";


        if (ckrequired == "N") {

            if (ckhave_file_current != "N") {
                if (fileUp != undefined) {
                    $("#err_" + chIDFile + "").css("display", "none");
                    jQuery("#file_link_" + chIDFile).text("-");
                    $(".link-table a").css("text-decoration", "none");
                } else {
                    if (ckhave_file_current == undefined) {
                        $("#err_" + chIDFile + "").css("display", "none");
                        jQuery("#file_link_" + chIDFile).text("-");
                        $(".link-table a").css("text-decoration", "none");
                    } else {
                        $("#err_" + chIDFile + "").css("display", "none");
                        setOnclickFileMOU(
                            core_datas.current_profile.profile_id,
                            ckhave_file_current,
                            document_type_id,
                            chIDFile
                        );
                    }
                }
            } else {
                if (fileUp != undefined) {
                    $("#err_" + chIDFile + "").css("display", "none");
                } else {
                    $("#err_" + chIDFile + "").css("display", "none");
                    jQuery("#file_link_" + chIDFile).text("-");
                    $(".link-table a").css("text-decoration", "none");
                }
            }
        } else {
            if (
                ckhave_file_current == "N" ||
                ckhave_file_current == "" ||
                ckhave_file_current == undefined
            ) {
                if (fileUp != undefined) {
                    $("#err_" + chIDFile + "").css("display", "none");
                } else {
                    if (doc_new_lastest == "true") {
                        $("#err_" + chIDFile + "").css("display", "none");
                        setOnclickFileMOU(
                            core_datas.current_profile.profile_id,
                            ckhave_file_current,
                            document_type_id,
                            chIDFile
                        );
                    } else {
                        chkData = false;
                        $("#err_" + chIDFile + "").css("display", "inline");
                        jQuery("#file_link_" + chIDFile).text("-");
                        $(".link-table a").css("text-decoration", "none");
                    }
                }
            } else {

                if (fileUp != undefined) {
                    $("#err_" + chIDFile + "").css("display", "none");
                } else {
                    setOnclickFileMOU(
                        core_datas.current_profile.profile_id,
                        ckhave_file_current,
                        document_type_id,
                        chIDFile
                    );

                }
            }
        }
    });

    var arrayFileRequest = [];
    var group_true_N_MOU = [];
    var group_true_MOU = [];
    var arrayfileUploadGroup_MOU = [];
    var group_false_MOU = [];
    var chkDataGroup_MOU = true;
    var chkDataGroup_MOU_Y = true;
    jQuery("[name='file_name_group']").each(function (index) {
        var app = getParameterByName("App");
        var fileUpload = jQuery("#" + this.id);
        var ckrequired = jQuery("#requiredvalue_" + this.id).val();
        var ckhave_file_current = jQuery("#have_file_current_" + this.id).val();
        var group_required = $("#required_" + this.id).val();
        var hidden_form_type_file_name = $("#hidden_form_type_" + this.id).val();
        var hidden_form_id = $("#hidden_form_id_" + this.id).val();
        var hidden_check_file_name = $("#hidden_check_" + this.id).val();
        var doc_new_lastest = $("#doc_new_lastest_" + this.id).val();
        var hidden_select = $("#hidden_select_" + this.id).val();
        var document_type_id = $("#document_type_id_" + this.id).val();
        var fileUp = fileUpload[0].files[0];
        var arrayFileNew = [];
        var chIDFile = fileUpload[0].id;


        if (!filevalidate.includes(group_required)) {
            filevalidate.push(group_required);
        }


        if (ckrequired == "N") {

            if (ckhave_file_current != "N") {
                KeepFile(group_required);
                setOnclickFileMOU(
                    core_datas.current_profile.profile_id,
                    ckhave_file_current,
                    document_type_id,
                    chIDFile
                );
            } else {
                if (fileUp != undefined) {
                    KeepFile(group_required);
                } else {
                    if (doc_new_lastest == "true") {
                        setOnclickFileMOU(
                            core_datas.current_profile.profile_id,
                            ckhave_file_current,
                            document_type_id,
                            chIDFile
                        );
                    } else {
                        if (
                            ckrequired == "N" &&
                            hidden_check_file_name != "UC" &&
                            hidden_check_file_name !== undefined
                        ) {
                            if (!arrayfileUploadGroup_MOU.includes(group_required)) {
                                jQuery("#file_link_" + chIDFile).text("-");
                                $(".link-table a").css("text-decoration", "none");
                                Removefiledupilcate(group_required);
                            }
                        } else if (
                            ckrequired == "N" &&
                            hidden_check_file_name == undefined
                        ) {
                            jQuery("#file_link_" + chIDFile).text("-");
                            $(".link-table a").css("text-decoration", "none");
                            chkDataGroup_MOU = true;
                            group_false_MOU = group_false_MOU.filter(function (value) {
                                return !group_true_N_MOU.includes(value);
                            });
                            KeepFile(group_required);
                            if (!arrayfileUploadGroup_MOU.includes(group_required)) {
                                arrayfileUploadGroup_MOU.push(group_required);
                            }
                        } else {
                            jQuery("#file_link_" + chIDFile).text("-");
                            $(".link-table a").css("text-decoration", "none");
                            Removefiledupilcate(group_required);
                        }
                    }
                }
            }
        } else {
            if (
                ckhave_file_current == "N" ||
                ckhave_file_current == "" ||
                ckhave_file_current == undefined
            ) {
                if (fileUp != undefined) {
                    if (!group_true_MOU.includes(group_required)) {
                        group_true_MOU.push(group_required);
                    }
                    if (!arrayfileUploadGroup_MOU.includes(group_required)) {
                        arrayfileUploadGroup_MOU.push(group_required);
                    }
                    chkDataGroup_MOU_Y = true;
                    KeepFile(group_required);
                } else {
                    if (group_true_MOU.length != 0) {
                        chkDataGroup_MOU_Y = true;
                        if (!arrayfileUploadGroup_MOU.includes(group_required)) {
                            arrayfileUploadGroup_MOU.push(group_required);
                        }
                        group_false_MOU = group_false_MOU.filter(function (value) {
                            return !group_true_MOU.includes(value);
                        });
                    } else {
                        chkDataGroup_MOU_Y = false;
                        group_false_MOU.push(group_required);
                    }
                    jQuery("#file_link_" + chIDFile).text("-");
                    $(".link-table a").css("text-decoration", "none");
                }
            } else {
                KeepFile(group_required);
                arrayFileRequest.push(ckrequired);


                if (fileUp == undefined) {
                    if (!arrayfileUploadGroup_MOU.includes(group_required)) {
                        arrayfileUploadGroup_MOU.push(group_required);

                        setOnclickFileMOU(
                            core_datas.current_profile.profile_id,
                            ckhave_file_current,
                            document_type_id,
                            chIDFile
                        );
                    }
                } else {

                    if (!arrayfileUploadGroup_MOU.includes(group_required)) {
                        arrayfileUploadGroup_MOU.push(group_required);
                    }
                }

                if (!group_true_MOU.includes(group_required)) {
                    group_true_MOU.push(group_required);
                }


            }
        }
    });
    var allFieldsValid = true;
    if (chkDataGroup_MOU == false || chkDataGroup_MOU_Y == false) {
        filevalidate.forEach(function (fileName) {
            var errorElement = $("#err_file_name_" + fileName);
            if (group_false_MOU.includes(fileName)) {
                allFieldsValid = false;
                errorElement.css("display", "inline");
            } else {
                errorElement.css("display", "none");
            }
        });
    } else {
        filevalidate.forEach(function (fileName) {
            var errorElement = $("#err_file_name_" + fileName);
            if (arrayfileUploadGroup_MOU.includes(fileName)) {
                errorElement.css("display", "none");
            } else {

                allFieldsValid = false;
                errorElement.css("display", "inline");
            }
        });
    }

    var num = 1;
    var chkList = jQuery("#ShowtxtBox_File .checklength").length;
    if (chkList == 0) {
        $(".head_or").css("display", "none");
        $("#div_showFileOrther63").css("display", "none");
    } else {
        jQuery("#ShowtxtBox_File > .checklength").each(function (index) {
            $("#listFile_or63").empty();
            var uploadid = jQuery(this)[0].id.replace("getElementDoc", "");
            var fileUp = $("#hidden_file_up_" + uploadid)[0].files;
            var hidden_form_type = $("#hidden_form_type_" + uploadid).val();
            var DocIDValue_file_name = $("#DocIDValue_file_name_" + uploadid).val();
            var show_file_title = $("#show_file_title_" + uploadid).text();

            var hidden_form_id = $("#hidden_form_id").val();

            var arrayfile = [];
            var tmpfiletitle_name = [];

            var setData = "";
            if (fileUp.length > 0) {
                var objectUrl = URL.createObjectURL(fileUp[0]);

                jQuery("#file_other_name_" + uploadid).text(fileUp[0].name);
                arrayfile.push(objectUrl);

                tmpfiletitle_name.push(jQuery("#text_file_name_" + uploadid).val());
                var reader = new FileReader();

                reader.onload = function (event) {
                    let base64 = event.target.result;


                    if (tmpfiletitle_name.length != 0) {
                        $("#div_showFileOrther63").css("display", "block");
                        $(".head_or").css("display", "block");
                        $(".head_or").removeClass("d-none");
                        $("#div_showFileOrther63").removeClass("d-none");
                        $("#DataNull59").css("display", "none");
                        setData += "<tr>";
                        for (var c = 0; c < tmpfiletitle_name.length; c++) {
                            setData += "<td >" + num++ + "</td>";
                            setData +=
                                "<td class='text-left'>" + tmpfiletitle_name[c] + "</td>";

                            if (app == "Y" || app == "y") {
                                setData +=
                                    "<td  class='link-table'><a target='_blank'><img src='/assets/images/document.svg' class='iconbook' onclick='sendmessage(&quot;" +
                                    base64 +
                                    "&quot;)'></a ></td>";
                            } else {
                                if (show_file_title) {
                                    setData +=
                                        "<td  class='link-table'><a href='javascript:void(0)' id='FileDetailOpen_" +
                                        uploadid +
                                        "' target='_blank'><img src='/assets/images/document.svg' class='iconbook'></a ></td>";
                                    OpenFileCurrent(
                                        hidden_form_id,
                                        DocIDValue_file_name,
                                        hidden_form_type,
                                        uploadid
                                    );
                                } else {

                                    setData +=
                                        "<td data-label='เอกสาร' class='link-table'><a href='" +
                                        arrayfile[c] +
                                        "' target='_blank'><img src='/assets/images/document.svg' class='iconbook'></a ></td>";
                                }
                            }
                        }
                        setData += "</tr>";
                    } else {
                        $("#div_showFileOrther63").hide();
                        $(".head-fileorder").hide();
                    }

                    $("#listFile_or63").append(setData);
                };

                reader.readAsDataURL(fileUp[0]);
            } else {
                if (show_file_title) {
                    tmpfiletitle_name.push(show_file_title);
                }
            }
        });
    }

    if (chkData && allFieldsValid) {
        $("#summaries").addClass("active");
        current_fs = $(this).parent();
        next_fs = $(this).parent().next();
        next_fs.show();

        current_fs.animate(
            { opacity: 0 },
            {
                step: function (now) {

                    opacity = 1 - now;

                    current_fs.css({
                        display: "none",
                        position: "relative",
                    });
                    next_fs.css({ opacity: opacity });
                },
                duration: 500,
            }
        );
        setProgressBar(++current);
    }
});

function openActiveSection(activeIndex, type) {
    const allTabs = [
        "tab_default_1",
        "tab_default_2",
        "tab_default_3",
        "tab_default_4",
        "tab_default_5",
        "tab_default_6",
        "tab_default_7",
    ];
    const allDivs = [
        "div_tab1",
        "div_tab2",
        "div_tab3",
        "div_tab4",
        "div_tab5",
        "div_tab6",
        "div_tab7",
    ];

    const deleteIndex = activeIndex - 1;

    if (type == "active") {
        allTabs.forEach((tab) => $("#" + tab).removeClass("active"));
        allDivs.forEach((div) => $("#" + div).removeClass("active"));

        if (activeIndex != null) {
            $("#" + allTabs[deleteIndex]).addClass("active");
            $("#" + allDivs[deleteIndex]).addClass("active");
        }
    } else {
        $("#" + allDivs[deleteIndex]).show();
    }
}
async function handleTabClick(tabId) {
    let activeTabId = $(".active").attr("id");

    if (tabId == "" || tabId == null || tabId == undefined) {
        tabId = activeTabId;
    }
    if (tabId == "div_tab5") {


        $("#link_appointment").attr("src", "about:blank");
        GetDataFormAppointmentMoreDocByGroupID(), GetPostiFrameQueue();
    }
}
function clickradio() {
    $("#radiopayment").trigger("click");
    $("#btnpayment").prop("disabled", false);
    $("#paymentUpDate").prop("disabled", false);
}

async function GetDataPaymentInquiryTransactionMOU(group_id, check) {
    let _userId = null;

    if (core_datas.default_profile.user_type_id == "5") {
        _userId = core_datas.default_profile.user_id;
    } else {
        _userId = core_datas.default_profile.user_id;
    }
    let objRequest = {
        group_id: group_id,
        user_id: _userId,
        relate_account_id: core_datas.current_profile.profile_id,
    };

    var request = await jQuery.post(
        url_GetDataPaymentInquiryTransaction,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                var setPayment = "";
                if (objData.length != 0) {
                    $("#div_no_payment").addClass("d-none");
                    setPayment += "<div class='col-12'>";
                    setPayment +=
                        "<h3 class='mb-4 mt-3 head-step lang_payment'>การชำระเงิน</h3>";
                    setPayment += "</div>";
                    if (check) {
                        setPayment += "<div class='col-6'>";
                        setPayment +=
                            '<div class="row mb-3 badgestatus show_for_proxy" style="background-color:#FFFBEB;color:#DCB745;border-radius:5px;display:inline-flex;">';
                        setPayment += '<div class="">';
                        setPayment +=
                            '<img src="/assets/images/picmobile/warning.png" style="width: 25px;">';
                        setPayment += "</div>";
                        setPayment +=
                            '<p class="ml-2 lang_make_payments_appointments" style="margin-bottom:0 !important">ผู้ที่ทำการยื่นคำขอเท่านั้นที่สามารถทำการ ชำระเงินและนัดหมายได้</p>';
                        setPayment += "</div>";
                        setPayment += "</div>";
                    }

                    for (var i = 0; i < objData.length; i++) {
                        transaction_id = objData[i].ref_transaction_id;
                        paymentobj = objData[i];
                        setPayment +=
                            "<div class='form-card text-left' id='div_showselectpay_new'>";
                        setPayment += "<div class='card'>";
                        setPayment += "<div class='card-body'>";
                        setPayment += "<div class='row'>";
                        setPayment += "<div class='col-12 col-md-3'>";
                        setPayment += "<p class='text-secondary lang_lbl_list'>รายการ</p>";
                        setPayment +=
                            "<p class='b font-weight-bold' data_th='" +
                            objData[i].payment_type_name_th?.toSafeJsonHtml() +
                            "' data_en='" +
                            objData[i].payment_type_name_en +
                            "'></p>";
                        setPayment += "</div>";
                        setPayment += "<div class='col-12 col-md-3'>";
                        setPayment +=
                            "<p class='text-secondary lang_payment_schedule'>กำหนดการชำระเงิน</p>";
                        setPayment +=
                            "<p class='b font-weight-bold'><span data_th='" +
                            convert_to_thai_datetime(objData[i].payment_expired_time) +
                            "' data_en='" +
                            convert_to_eng_datetime(objData[i].payment_expired_time) +
                            "'></span></p>";
                        setPayment += "</div>";
                        setPayment += "<div class='col-12 col-md-3'>";
                        setPayment += "<p class='text-secondary lang_lbl_status'>สถานะ</p>";

                        if (objData[i].payment_status_id == "W") {
                            setPayment +=
                                "<p id='form_activity_detail' class='badge badge-pill status11-appoint'  data_th='" +
                                objData[i].payment_status_name_th?.toSafeJsonHtml() +
                                "' data_en='" +
                                objData[i].payment_status_name_en +
                                "'></p>";
                        } else if (objData[i].payment_status_id == "PS") {
                            setPayment +=
                                "<p id='form_activity_detail' class='badge badge-pill status13' data_th='" +
                                objData[i].payment_status_name_th?.toSafeJsonHtml() +
                                "' data_en='" +
                                objData[i].payment_status_name_en +
                                "'></p>";
                        } else if (objData[i].payment_status_id == "PF") {
                            setPayment +=
                                "<p id='form_activity_detail' class='badge badge-pill status12' data_th='" +
                                objData[i].payment_status_name_th?.toSafeJsonHtml() +
                                "' data_en='" +
                                objData[i].payment_status_name_en +
                                "'></p>";
                        }
                        setPayment += "</div>";
                        setPayment += "<div class='col-12 col-md-3 pt-3'>";
                        if (objData[i].payment_type_id == "COLLATERAL") {
                            if (objData[i].is_mou_enabled == "N" && getParameterByName("status") != "CC") {
                                setPayment +=
                                    '<button type="button" class="btn lang_make_payment btn-primary next action-button float-right w-100-bt" _payment-obj="' +
                                    JSON.stringify(objData[i]).replace(/"/g, "&quot;") +
                                    '" onclick="OpenModalPaymentStepOne(this)" disabled><span class="lang_make_payment">ชำระเงิน</span></button>';
                            } else {
                                if (objData[i].payment_status_id == "W" && getParameterByName("status") != "CC") {
                                    setPayment +=
                                        '<button type="button" class="btn lang_make_payment btn-primary next action-button float-right w-100-bt" _payment-obj="' +
                                        JSON.stringify(objData[i]).replace(/"/g, "&quot;") +
                                        '" onclick="OpenModalPaymentStepOne(this)"><span class="lang_make_payment">ชำระเงิน</span></button>';
                                } else if (objData[i].payment_status_id == "PS") {
                                    setPayment +=
                                        "<button type='button' class='btn btn-dark m-0 next action-button float-right w-100-bt lang_proof_payment' onclick='Openproofpayment(" +
                                        JSON.stringify(objData[i].ref_transaction_id) +
                                        ")'>หลักฐานการชำระเงิน</button>";
                                }
                            }
                        } else {
                            if (objData[i].payment_status_id == "W" && getParameterByName("status") != "CC") {
                                setPayment +=
                                    '<button type="button" class="btn lang_make_payment btn-primary next action-button float-right w-100-bt" _payment-obj="' +
                                    JSON.stringify(objData[i]).replace(/"/g, "&quot;") +
                                    '" onclick="OpenModalPaymentStepOne(this)"><span class="lang_make_payment">ชำระเงิน</span></button>';
                            } else if (objData[i].payment_status_id == "PS") {
                                setPayment +=
                                    "<button type='button' class='btn btn-dark m-0 next action-button float-right w-100-bt lang_proof_payment' onclick='Openproofpayment(" +
                                    JSON.stringify(objData[i].ref_transaction_id) +
                                    ")'>หลักฐานการชำระเงิน</button>";
                            }
                        }

                        setPayment += "</div>";
                        setPayment += "</div>";
                        setPayment += "</div>";
                        setPayment += "</div>";
                        setPayment += "</div>";
                    }
                    setPayment +=
                        "<p class='lang_fee_and_service_first'>*กรุณาชำระค่าคำขอเเละค่าธรรมเนียมก่อน จึงจะสามารถชำระค่าวางเงินหลักประกันได้</p>";
                    jQuery("#paymentlist46_59").html(setPayment);
                    jQuery("#paymentlist41").html(setPayment);
                } else {
                    if (check) {
                        $("#div_no_payment").addClass("d-none");
                        $("#div_no_payment_no_proxy").removeClass("d-none");
                    } else {
                        $("#div_no_payment").removeClass("d-none");
                    }
                }
            } else {
                $("#div_no_payment").removeClass("d-none");
            }
        },
        "json"
    );

    await changeDatalang();
}
function onloadfirst() {
    $("#birthdate_al,#dateofexpiry,#enddateofexpiry")
        .datepicker({
            format: "DD/MM/YYYY",
            startView: "day",
            minViewMode: 0,
            maxViewMode: 4,
            autoclose: true,
        })
        .on("change", function (ev) { });
    $("#dateofexpiry").on("change", function () {
        $("#enddateofexpiry").val("");
        if ($("#dateofexpiry").val()) {
            $("#dateofexpiry").removeClass("error");
            var sp_date = this.value.split("/");
            var specificDateFromString =
                sp_date[2] + "-" + sp_date[1] + "-" + sp_date[0];
            $("#enddateofexpiry").datepicker({
                minDate: new Date(specificDateFromString),
            });
        }


    });


}


function validateFiles() {
    let isValid = true;
    let groupErrors = {};
    $('[id^="file_name_"]').each(function () {
        const fileId = $(this).attr("id");
        const fileNumber = fileId.split("_").pop();
        const fileName = $(this).attr("name");
        const requiredValue = $("#requiredvalue_file_name_" + fileNumber).val();
        const group_file = $("#required_file_name_" + fileNumber).val();
        const errorElement = $("#err_file_name_" + fileNumber);
        const fileInput = $("#file_name_" + fileNumber);
        var have_file_current = $(
            "#have_file_current_file_name_" + fileNumber
        ).val();

        if (requiredValue === "Y" && fileName === "file_name_group" && group_file) {
            let groupValid = false;

            $('[id^="file_name_"]').each(function () {
                const currentFileInput = $(this);
                const currentGroupFile = $(
                    "#required_file_name_" + currentFileInput.attr("id").split("_").pop()
                ).val();
                if (
                    currentGroupFile === group_file &&
                    currentFileInput[0].files.length > 0
                ) {
                    groupValid = true;
                    return false;
                }
            });

            if (groupValid || groupErrors[group_file] === false) {
                errorElement.hide();
                groupErrors[group_file] = false;
            } else {
                groupErrors[group_file] = true;
                errorElement.show();
            }
        } else {
            if (have_file_current != "N" || fileInput[0].files.length > 0) {
                errorElement.hide();
            } else {
                errorElement.show();
                isValid = false;
            }
        }
    });

    for (const group_file in groupErrors) {
        if (groupErrors[group_file]) {
            $("#err_file_name_" + group_file).show();
        } else {
            $("#err_file_name_" + group_file).hide();
        }
    }
    return isValid;
}
function validateFilesUC() {
    let isValid = true;
    let groupErrors = {};
    $('[id^="file_name_"]').each(function () {
        const fileId = $(this).attr("id");
        const fileNumber = fileId.split("_").pop();
        const fileName = $(this).attr("name");
        const requiredValue = $("#requiredvalue_file_name_" + fileNumber).val();
        const checkValue = $("#hidden_check_file_name_" + fileNumber).val();
        const group_file = $("#required_file_name_" + fileNumber).val();
        const errorElement = $("#err_file_name_" + fileNumber);

        if (fileNumber === "72" && $("#doc_new_lastest").val() === "Y") {
            errorElement.hide();
            return true;
        }

        const fileInput = $("#file_name_" + fileNumber);
        var hidden_select = $("#hidden_select_file_name_" + fileNumber).val();
        var have_file_current = $(
            "#have_file_current_file_name_" + fileNumber
        ).val();

        if (checkValue === "UC") {
            if (
                checkValue === "UC" ||
                (requiredValue === "Y" &&
                    (fileName == "file_name_group" || fileName == "file_name"))
            ) {
                if (fileName === "file_name_group" && group_file) {

                    let groupValid = false;

                    $('[id^="file_name_"]').each(function () {
                        const currentFileInput = $(this);
                        const currentGroupFile = $(
                            "#required_file_name_" +
                            currentFileInput.attr("id").split("_").pop()
                        ).val();
                        if (
                            currentGroupFile === group_file &&
                            currentFileInput[0].files.length > 0
                        ) {
                            groupValid = true;
                            return false;
                        }
                    });


                    if (groupValid) {
                        if (checkValue === "UC") {
                            if (hidden_select && hidden_select != "Y") {
                                groupErrors[group_file] = false;
                            } else {
                                groupErrors[group_file] = true;
                                isValid = false;
                            }
                        } else {
                            if (have_file_current != "" && have_file_current != "N") {
                                groupErrors[group_file] = false;
                                errorElement.hide();
                            } else {
                                groupErrors[group_file] = true;
                            }
                        }
                    } else {
                        if (checkValue == "UC") {
                            if (hidden_select && hidden_select != "Y") {
                                groupErrors[group_file] = false;
                            } else {
                                groupErrors[group_file] = true;
                                isValid = false;
                            }
                        } else {
                            if (have_file_current != "" && have_file_current != "N") {
                                groupErrors[group_file] = false;
                            } else {
                                groupErrors[group_file] = true;
                                errorElement.show();
                                isValid = false;
                            }
                        }
                    }
                } else if (fileInput[0].files.length > 0) {

                    errorElement.hide();
                } else {
                    if (hidden_select && hidden_select != "Y") {
                        errorElement.hide();
                    } else {
                        errorElement.show();
                        isValid = false;
                    }

                }
            } else {
                errorElement.hide();
            }
        }
    });

    for (const group_file in groupErrors) {
        if (groupErrors[group_file]) {
            $("#err_file_name_" + group_file).show();
        } else {
            $("#err_file_name_" + group_file).hide();
        }
    }
    return isValid;
}

let checkmt;
async function handleDraft() {
    try {
        const draft = await loadDraftFromIndexedDB();
        if (draft) {
            if (
                "draft_" + draft.user_id + "" == "draft_" + core_datas.user_id + "" &&
                draft.form_type_id == formtypeID
            ) {
                Swal.fire({
                    imageUrl: important_data.warningIconUrl,
                    title: '<span class="lang_save_draft">บันทึกฉบับร่าง</span>',
                    html: `<span class="lang_You_have_draft">คุณมีฉบับร่างต้องการใช้ฉบับร่างหรือไม่?</span>`,
                    showCancelButton: true,
                    cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
                    confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                    reverseButtons: true,
                    allowOutsideClick: false,
                    allowEscapeKey: false
                }).then((result) => {
                    if (result.isConfirmed) {
                        checkmt = draft.form_type_id;
                        console.log(draft);
                        $("#workhour_per_day").val(draft.content.workhour_per_day);
                        $("#traditional_holiday").val(draft.content.traditional_holiday);
                        $("#wage_rate_per_day").val(draft.content.wage_rate_per_day);
                        $("#holiday_per_year").val(draft.content.holiday_per_year);
                        $("#dayoff_per_week").val(draft.content.dayoff_per_week);
                        $("#dayoff_per_week").val(draft.content.dayoff_per_week);
                        $("#start_age").val(draft.content.start_age);
                        $("#start_height").val(draft.content.start_height);
                        $("#start_weight").val(draft.content.start_weight);
                        $("#end_age").val(draft.content.end_age);
                        $("#end_height").val(draft.content.end_height);
                        $("#end_weight").val(draft.content.end_weight);
                        $("#nationality_id_mou").val(draft.content.nationality_id_mou);
                        DataImmigrationCheckpoint = draft.content.DataImmigrationCheckpoint;
                        GetDataImmigrationCheckpoint();
                        recruitment_agency = draft.content.recruitment_agency;
                        fsc_id = draft.content.fsc_id;


                        name_address_work = draft.content.name_address_work;
                        bus_type_id = draft.content.bus_type_id;

                        if (draft.content.permit_cate_id) {
                            $("#checkboxjob_4").prop("checked", true);
                            $("#numbermen_checkboxjob_4").val(
                                draft.content.number_of_alien_male
                            );
                            $("#numbermen_checkboxjob_4").attr("disabled", false);
                            $("#numberwomen_checkboxjob_4").val(
                                draft.content.number_of_alien_female
                            );
                            $("#numberwomen_checkboxjob_4").attr("disabled", false);
                            $("#permit_catework_Dropdown").val(draft.content.permit_cate_id);
                            $("#permit_catework_Dropdown").attr("disabled", false);
                        } else {
                            $("#checkboxjob_" + draft.content.permit_type).prop(
                                "checked",
                                true
                            );
                            $("#numbermen_checkboxjob_" + draft.content.permit_type).val(
                                draft.content.number_of_alien_male
                            );
                            $("#numbermen_checkboxjob_" + draft.content.permit_type).attr(
                                "disabled",
                                false
                            );
                            $("#numberwomen_checkboxjob_" + draft.content.permit_type).val(
                                draft.content.number_of_alien_female
                            );
                            $("#numberwomen_checkboxjob_" + draft.content.permit_type).attr(
                                "disabled",
                                false
                            );
                        }
                        $("#number_of_alien_all").val(draft.content.number_of_alien_all);
                        $("#number_of_alien_male").val(draft.content.number_of_alien_male);
                        $("#number_of_alien_female").val(
                            draft.content.number_of_alien_female
                        );

                        $("#address_recruitment").val(draft.content.address_recruitment);
                        $("#recruitment_agency_assign_at").val(
                            draft.content.recruitment_agency_assign_at
                        );
                        draftfile = draft.files;
                        draftfileorther = draft.fileorther;

                        if (
                            draft.content.employerID &&
                            draft.content.slt_employer_type !== ""
                        ) {
                            $("#slt_employer_type").val(draft.content.slt_employer_type);
                            $("#employerID_63").val(draft.content.employerID);
                            searchEmployer63();
                        } else {
                            onloadfiledraft();
                            onloadfileorther();
                        }

                        $("#name_address_work").val(draft.content.name_address_work);

                        if ($("#name_address_work :selected").val()) {
                            var slt_bus = $("#name_address_work :selected").attr(
                                "data-value-add"
                            );
                            var obj_bus = JSON.parse(slt_bus);
                            var setdatabus = "";
                            setdatabus +=
                                '<option class="lang_opt_pleaseselect" selected="" value="">กรุณาเลือก</option>';
                            for (var j = 0; j < obj_bus.length; j++) {
                                setdatabus +=
                                    "<option value='" +
                                    obj_bus[j].bus_type_id +
                                    "' data-sale='" +
                                    obj_bus[j].is_saleperson +
                                    "' data_th='" +
                                    obj_bus[j].bus_type_name_th +
                                    "'  data_en='" +
                                    obj_bus[j].bus_type_name_en +
                                    "'>" +
                                    obj_bus[j].bus_type_name_th +
                                    "</option>";
                            }
                            $("#bus_type_id").html(setdatabus);


                            $("#bus_type_id").val(draft.content.bus_type_id);
                        }


                    }
                });
                changeDatalang();
            }
        } else {
            console.log("No draft available or draft has expired.");
        }
    } catch (error) {
        console.error("Error loading draft:", error);
    }
}
async function saveDataDraft() {
    const content = {
        nationality_id_mou: $("#nationality_id_mou :selected").val(),
        recruitment_agency: $("#recruitment_agency :selected").val(),
        address_recruitment: $("#address_recruitment").val(),
        recruitment_agency_assign_at: $("#recruitment_agency_assign_at").val(),
        name_address_work: $("#name_address_work").val(),
        bus_type_id: $("#bus_type_id :selected").val(),

        start_age: $("#start_age :selected").val(),
        end_age: $("#end_age :selected").val(),
        start_height: $("#start_height").val(),
        end_height: $("#end_height").val(),
        start_weight: $("#start_weight").val(),
        end_weight: $("#end_weight").val(),

        workhour_per_day: $("#workhour_per_day").val(),
        dayoff_per_week: $("#dayoff_per_week").val(),
        holiday_per_year: $("#holiday_per_year").val(),
        wage_rate_per_day: $("#wage_rate_per_day").val(),
        traditional_holiday: $("#traditional_holiday").val(),

        DataImmigrationCheckpoint: $("#DataImmigrationCheckpoint").val(),
        fsc_id: $("#fsc_id").val(),
        number_of_alien_all: $("#number_of_alien_all").val(),
        number_of_alien_male: $("#number_of_alien_male").val(),
        number_of_alien_female: $("#number_of_alien_female").val(),
        is_saleperson: $("#issale").val(),
        permit_cate_id: $("#permit_catework_Dropdown :selected").val(),
        permit_type: $("input[type='radio'][name='permit_type']:checked").val(),
        slt_employer_type: $("#slt_employer_type :selected").val(),
        employerID: $("#employerID_63").val(),
    };

    let files = [];
    let fileorther = [];
    let filesnoemp = [];
    try {
        const file_name = jQuery("[name='file_name'], [name='file_name_group']")
            .map(async function () {
                var fileUpload = jQuery("#" + this.id);
                var file_up = fileUpload[0].files[0];

                if (file_up !== undefined) {
                    var file_name_form_type_doc_id = $(
                        "#form_type_doc_id_" + this.id
                    ).val();
                    var doc_type_id = this.id.split("_");

                    files.push({
                        blob: file_up,
                        fileTypeId: file_name_form_type_doc_id,
                        document_type_id: doc_type_id[2],
                        title: "",
                    });
                }
            })
            .get();

        const file_orther = jQuery("#ShowtxtBox_File > .checklength")
            .map(async function () {
                var uploadid = jQuery(this)[0].id.replace("getElementDoc", "");
                var fileUp = $("#hidden_file_up_" + uploadid)[0].files[0];

                if (fileUp.length != 0) {
                    var hidden_form_type = $("#hidden_form_type_up").val();
                    var filename_orther = $("#text_file_name_" + uploadid).val();
                    var doc_type_id = this.id.split("_");

                    fileorther.push({
                        blob: fileUp,
                        fileTypeId: hidden_form_type,
                        document_type_id: doc_type_id[2],
                        title: filename_orther,
                    });
                }
            })
            .get();


        await Promise.all(file_name, file_orther);


        const draftObj = {
            content: content,
            files: files,
            fileorther: fileorther,
            filenoemp: filesnoemp,
            form_type_id: formtypeID,
            user_id: core_datas.user_id,
            form_type_nameTH: form_type_nameTH,
            form_type_nameEN: form_type_nameEN,
        };

        console.log(draftObj);

        const draft = await loadDraftFromIndexedDB();
        if (draft) {
            if ("draft_" + draft.user_id + "" == "draft_" + core_datas.user_id + "") {
                if (draft.form_type_id == formtypeID) {
                    Swal.fire({
                        imageUrl: important_data.warningIconUrl,
                        title: '<span class="lang_save_draft">บันทึกฉบับร่าง</span>',
                        html: '<span class="lang_have_existing_draft">คุณมีฉบับร่างเดิมอยู่เเล้วต้องการบันทึกทับฉบับร่างเดิมหรือไม่?</span> ',
                        showCancelButton: true,
                        cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
                        confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                        reverseButtons: true,
                        allowOutsideClick: false,
                        allowEscapeKey: false
                    }).then((result) => {
                        if (result.isConfirmed) {
                            saveDraftToIndexedDB(draftObj);
                        }
                    });
                    changeDatalang();
                } else {
                    if (draft.form_type_id) {
                        var splitformtypeID = draft.form_type_id.split("_");
                    }

                    const draftFormTypeNameTh = parseIfJson(draft.form_type_nameTH);
                    let _draftFormTypeNameTh = draftFormTypeNameTh;
                    let _draftFormTypeNameEn = draft.form_type_nameEN;
                    if (typeof draftFormTypeNameTh == "object") {
                        _draftFormTypeNameTh =
                            draftFormTypeNameTh["TH"] ||
                            draftFormTypeNameTh["Th"] ||
                            draftFormTypeNameTh["th"];
                        _draftFormTypeNameEn =
                            draftFormTypeNameTh["TH"] ||
                            draftFormTypeNameTh["En"] ||
                            draftFormTypeNameTh["en"];
                    }

                    const formTypeNameTh = parseIfJson(form_type_nameTH);
                    let _formTypeNameTh = formTypeNameTh;
                    let _formTypeNameEn = form_type_nameEN;
                    if (typeof formTypeNameTh == "object") {
                        _formTypeNameTh =
                            formTypeNameTh["TH"] ||
                            formTypeNameTh["Th"] ||
                            formTypeNameTh["th"];
                        _formTypeNameEn =
                            formTypeNameTh["TH"] ||
                            formTypeNameTh["En"] ||
                            formTypeNameTh["en"];
                    }

                    Swal.fire({
                        imageUrl: important_data.warningIconUrl,
                        title: '<span class="lang_save_draft">บันทึกฉบับร่าง</span>',
                        html: `<span class="lang_you_have_draft_of">คุณมีฉบับร่างเดิมของ</span> <span data_th="${_draftFormTypeNameTh}" data_en="${_draftFormTypeNameEn}"></span> <span class="lang_draft_already_exists">อยู่เเล้วต้องการบันทึกทับฉบับร่างเดิมหรือไม่?</span>`,

                        showCancelButton: true,
                        cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
                        confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                        reverseButtons: true,
                        allowOutsideClick: false,
                        allowEscapeKey: false
                    }).then((result) => {
                        if (result.isConfirmed) {
                            saveDraftToIndexedDB(draftObj);
                        }
                    });
                    changeDatalang();
                }
            }
        } else {
            Swal.fire({
                imageUrl: important_data.warningIconUrl,
                title: '<span class="lang_save_draft">บันทึกฉบับร่าง</span>',
                html: '<span class="lang_Do_you_want_save_draft">คุณต้องการบันทึกแบบร่างหรือไม่?</span>',
                showCancelButton: true,
                cancelButtonText: '<span class="lang_txt_cancel">ยกเลิก</span>',
                confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                reverseButtons: true,
                allowOutsideClick: false,
                allowEscapeKey: false
            }).then((result) => {
                if (result.isConfirmed) {
                    saveDraftToIndexedDB(draftObj);
                }
            });
            changeDatalang();
        }
    } catch (error) {
        console.error("Error saving draft:", error);
    }
}
async function onloadfiledraft() {
    if (draftfile) {
        if (draftfile.length != 0) {
            for (var i = 0; i < draftfile.length; i++) {
                $("#divselect_file_name_" + draftfile[i].document_type_id).css(
                    "display",
                    "none"
                );
                $("#divshow_file_name_" + draftfile[i].document_type_id).css(
                    "display",
                    "flex"
                );
                $("#show_file_name_" + draftfile[i].document_type_id).text(
                    draftfile[i].blob.name
                );
                var fileURL = URL.createObjectURL(draftfile[i].blob);
                var span = $("#show_file_name_" + draftfile[i].document_type_id);
                if (span.parent("a").length === 0) {
                    span.wrap('<a href="' + fileURL + '" target="_blank"></a>');
                }

                $("#file_link_file_name_" + draftfile[i].document_type_id).html(
                    '<a href="' +
                    fileURL +
                    '" target="_blank"><img src="/assets/images/document.svg" class="iconbook"></a>'
                );

                var inputId = "file_name_" + draftfile[i].document_type_id;
                var blob = draftfile[i].blob;
                var fileName = draftfile[i].blob.name;
                setFileInputFromBlob(inputId, blob, fileName);

                $("[name='file_name_group']").each(function (index) {
                    let fileName = $(this).val();
                    if (fileName != "") {
                        let required = $("#required_" + this.id).val();
                        if (!filevalidate.includes(required)) {
                            filevalidate.push(required);
                        }
                    }
                });
            }
        }
    }
}
async function onloadfileorther() {
    var objother_doclist = draftfileorther;

    var setData_OR = "";
    setData_OR += '<div class="row col-12">';
    setData_OR +=
        '<h3 class="head-step lang_other_doc" style="width:48%;">เอกสารอื่นๆ เพิ่มเติม</h3>';
    setData_OR +=
        '<div class="buttonaddfile" style="width:50%;"><button type="button" class="btn btn-add-file action-button float-right" onclick="OncOpenModalAddFile()"><img src="/assets/images/plus.svg"><span class="lang_add_document"></span></button></div>';
    setData_OR += "</div>";
    setData_OR +=
        '<span class="head"><p class="text-upload-subtext lang_subtext_upload">เอกสารที่ใช้ประกอบการพิจารณาทั้งลูกจ้างและนายจ้าง สูงสุด 5 ไฟล์</p></span> ';


    if (objother_doclist) {
        if (objother_doclist.length == 0) {
        } else {
            $(".empty_file_orther").addClass("d-none");
            setData_OR +=
                '<div class="col-md-12" id="ShowtxtBox_File" style="padding-left:0px;">';

            for (var p = 0; p < objother_doclist.length; p++) {
                var fileURL = URL.createObjectURL(objother_doclist[p].blob);
                createHiddenFileInput(
                    p,
                    objother_doclist[p].blob,
                    objother_doclist[p].title
                );
                setData_OR +=
                    '<input type="hidden" id="hidden_check_' + p + '" value="Y">';

                setData_OR +=
                    '<div id="getElementDoc' +
                    p +
                    '" class="checklength gutters getElementDoc">';
                setData_OR += '<div class="col-md-12" style="padding-left:0px;">';
                setData_OR += `<p class='text-upload'>${objother_doclist[p].title}`;
                setData_OR += "</p>";
                setData_OR += "<div class='row'>";
                setData_OR += "<div class='col-lg-12'>";
                setData_OR +=
                    "<p class='text-upload-subtext lang_support_pdf'>รองรับไฟล์ .PDF ขนาดไม่เกิน 25 mb</p>";
                setData_OR += "<div class='row' style='margin-left:0px;'>";
                setData_OR +=
                    "<div class='upload' id='upload_" +
                    p +
                    "' style='padding-right:14px;'>";

                setData_OR +=
                    "<input type='file' name='file_name_orther' id='file_other_" +
                    p +
                    "' onchange='SelectFileOrther(event,&quot;" +
                    p +
                    "&quot;)' hidden/>";

                setData_OR +=
                    "<label for='file_other_" +
                    p +
                    "' style='cursor:pointer;'class='lang_lbl_selectfile'>เลือกไฟล์</label>";
                setData_OR += "</div>";
                setData_OR +=
                    "<div class='block-doc justify-content-between border_file_success' id='divshow_file_name_" +
                    p +
                    "' style='display:flex;'>";

                setData_OR += "<div class='d-flex'>";
                setData_OR += "<div><img src='/assets/images/document.svg'></div>";

                setData_OR += "<div class='filename'>";

                setData_OR +=
                    "<a id='tagaorther_" +
                    p +
                    "' href='" +
                    fileURL +
                    "' target='_blank'><span class='name-file namedoc' id='show_orther_file_name_" +
                    p +
                    "' value='" +
                    objother_doclist[p].title +
                    "'>" +
                    objother_doclist[p].blob.name +
                    "</span></a>";
                setData_OR += "</div>";
                setData_OR += "</div>";

                setData_OR +=
                    "<div><img class='close' src='/assets/images/Mark.svg' onclick='deleteboxfile63(&quot;" +
                    p +
                    "&quot;,&quot;" +
                    objother_doclist[p].title +
                    "&quot;)'></div>";

                setData_OR += "</div>";
                setData_OR += "</div>";
                setData_OR += "</div>";
                setData_OR += "</div>";

                setData_OR +=
                    '<input type="hidden" id="hidden_form_type_up" value="' +
                    objother_doclist[p].fileTypeId +
                    '" />';
                setData_OR +=
                    '<input type="hidden" id="text_file_name_' +
                    p +
                    '" value="' +
                    objother_doclist[p].title +
                    '" />';
                setData_OR +=
                    "<div class='col-md-12 mb-2 mt-2' style='padding-left:0px;' id='line_" +
                    p +
                    "'><hr></div>";
                setData_OR += "</div>";
                setData_OR += "</div>";
            }
            setData_OR += "</div>";
            jQuery("#uploadfile_orther").html(setData_OR);
        }
    }
    await changeDatalang();
}


;
$("#nextstepcheckfileEditMOU").on("click", function () {
    var app = getParameterByName("App");
    var chkData = true;
    jQuery("[name='file_name']").each(function (index) {
        var app = getParameterByName("App");
        var fileUpload = jQuery("#" + this.id);
        var ckrequired = jQuery("#requiredvalue_" + this.id).val();
        var ckhave_file_current = jQuery("#have_file_current_" + this.id).val();
        var hidden_form_type_file_name = $("#hidden_form_type_" + this.id).val();
        var hidden_form_id = $("#hidden_form_id_" + this.id).val();
        var hidden_check_file_name = $("#hidden_check_" + this.id).val();
        var doc_new_lastest = $("#doc_new_lastest_" + this.id).val();
        var hidden_select = $("#hidden_select_" + this.id).val();
        var fileUp = fileUpload[0].files[0];
        var arrayFileNew = [];
        var chIDFile = fileUpload[0].id;


        if (ckrequired == "N") {

            if (ckhave_file_current != "N") {
                if (fileUp != undefined) {
                    $("#err_" + chIDFile + "").css("display", "none");
                } else {
                    if (hidden_check_file_name == "UC") {
                        if (fileUp != undefined) {
                            $("#err_" + chIDFile + "").css("display", "none");
                        } else {
                            if (doc_new_lastest == "true") {
                                $("#err_" + chIDFile + "").css("display", "none");
                                setOnclickFile(
                                    ckhave_file_current,
                                    hidden_form_type_file_name,
                                    hidden_form_id,
                                    chIDFile
                                );
                            } else {
                                if (hidden_select == "0") {
                                    chkData = false;
                                    $("#err_" + chIDFile + "").css("display", "none");
                                } else {
                                    chkData = false;
                                    $("#err_" + chIDFile + "").css("display", "inline");
                                }
                            }
                        }
                    } else {
                        $("#err_" + chIDFile + "").css("display", "none");
                        setOnclickFile(
                            ckhave_file_current,
                            hidden_form_type_file_name,
                            hidden_form_id,
                            chIDFile
                        );
                    }
                }
            } else {
                if (fileUp != undefined) {
                    $("#err_" + chIDFile + "").css("display", "none");
                } else {
                    if (hidden_check_file_name == "UC") {
                        if (fileUp != undefined) {
                            $("#err_" + chIDFile + "").css("display", "none");
                        } else {
                            if (doc_new_lastest == "true") {
                                $("#err_" + chIDFile + "").css("display", "none");
                                setOnclickFile(
                                    ckhave_file_current,
                                    hidden_form_type_file_name,
                                    hidden_form_id,
                                    chIDFile
                                );
                            } else {
                                if (hidden_select == "Y") {
                                    chkData = true;
                                    $("#err_" + chIDFile + "").css("display", "none");
                                } else {
                                    chkData = false;
                                    $("#err_" + chIDFile + "").css("display", "inline");
                                }
                            }
                        }
                    } else {
                        $("#err_" + chIDFile + "").css("display", "none");
                        jQuery("#file_link_" + chIDFile).text("-");
                        $(".link-table a").css("text-decoration", "none");
                    }
                }
            }
        } else {
            if (
                ckhave_file_current == "N" ||
                ckhave_file_current == "" ||
                ckhave_file_current == undefined
            ) {
                if (fileUp != undefined) {
                    $("#err_" + chIDFile + "").css("display", "none");
                } else {
                    if (doc_new_lastest == "true") {
                        $("#err_" + chIDFile + "").css("display", "none");
                        setOnclickFile(
                            ckhave_file_current,
                            hidden_form_type_file_name,
                            hidden_form_id,
                            chIDFile
                        );
                    } else {
                        chkData = false;
                        $("#err_" + chIDFile + "").css("display", "inline");
                        jQuery("#file_link_" + chIDFile).text("-");
                        $(".link-table a").css("text-decoration", "none");
                    }
                }
            } else {

                if (fileUp != undefined) {
                    $("#err_" + chIDFile + "").css("display", "none");
                } else {
                    if (hidden_check_file_name == "UC") {
                        if (fileUp != undefined) {
                            $("#err_" + chIDFile + "").css("display", "none");
                        } else {
                            if (doc_new_lastest == "true") {
                                $("#err_" + chIDFile + "").css("display", "none");
                                setOnclickFile(
                                    ckhave_file_current,
                                    hidden_form_type_file_name,
                                    hidden_form_id,
                                    chIDFile
                                );
                            } else {
                                chkData = false;
                                $("#err_" + chIDFile + "").css("display", "inline");
                            }
                        }
                    } else {
                        setOnclickFile(
                            ckhave_file_current,
                            hidden_form_type_file_name,
                            hidden_form_id,
                            chIDFile
                        );
                    }
                }
            }
        }
    });

    let groupStatus = {}
    jQuery("[name='file_name_group']").each(function () {
        const group_required = $("#required_" + this.id).val();
        const hidden_check_file_name = $("#hidden_check_" + this.id).val();
        const fileUpload = jQuery("#" + this.id)[0];
        const fileUp = fileUpload.files[0];
        const hasOldFile = $("#have_file_current_" + this.id).val() !== "N";

        const hidden_form_type_file_name = $("#hidden_form_type_" + this.id).val();
        const hidden_form_id = $("#hidden_form_id_" + this.id).val();
        const chIDFile = fileUpload.id;

        if (!groupStatus[group_required]) {
            groupStatus[group_required] = {
                hasAnyFile: false,
                hasUCModified: true
            };
        }


        if (fileUp) {
            groupStatus[group_required].hasAnyFile = true;

            if (hidden_check_file_name === "UC") {
                groupStatus[group_required].hasUCModified = true;
            }


        }


        if (!fileUp && hasOldFile) {
            if (hidden_check_file_name === "UC") {

                groupStatus[group_required].hasUCModified = false;
            } else {

                groupStatus[group_required].hasAnyFile = true;

                setOnclickFile(
                    $("#have_file_current_" + this.id).val(),
                    hidden_form_type_file_name,
                    hidden_form_id,
                    chIDFile
                );
            }
        }
    });


    let allFieldsValid = true;
    Object.keys(groupStatus).forEach(group => {
        const status = groupStatus[group];
        const errorElement = $("#err_file_name_" + group);

        if (!status.hasAnyFile || !status.hasUCModified) {
            allFieldsValid = false;
            errorElement.css("display", "inline");
        } else {
            errorElement.css("display", "none");
        }
    });

    var checkorther = true;
    var setData = "";
    var num = 1;
    var app = getParameterByName("App");
    var chkList = jQuery("#ShowtxtBox_File .checklength").length;
    if (chkList == 0) {
        $(".head_or").css("display", "none");
        $("#div_showFileOrther63").css("display", "none");
    } else {
        jQuery("#ShowtxtBox_File > .checklength").each(function (index) {
            var uploadid = jQuery(this)[0].id.replace("getElementDoc", "");

            var fileUp2 = $("#hidden_file_up_" + uploadid)[0].files;
            var ckhave_file_current = jQuery("#have_file_current_" + this.id).val();
            var hidden_form_type = $("#hidden_form_type_" + uploadid).val();
            var DocIDValue_file_name = $("#DocIDValue_file_name_" + uploadid).val();
            var show_file_title = $("#show_file_title_" + uploadid).text();
            var hidden_form_id = $("#hidden_form_id").val();
            var hidden_check_file_name = $("#hidden_check_" + uploadid).val();

            var hidden_current_file_new = "";
            var arrayfile = [];
            var tmpfiletitle_name = [];

            var objectUrl = "";
            var reader = new FileReader();

            if (fileUp2.length > 0) {
                reader.readAsDataURL(fileUp2[0]);

                reader.onload = function (event) {
                    if (app == "Y" || app == "y") {
                        objectUrl = event.target.result;
                    } else {
                        objectUrl = URL.createObjectURL(fileUp2[0]);
                    }


                    jQuery("#file_other_name_" + uploadid).text(fileUp2[0].name);
                    arrayfile.push(objectUrl);
                    jQuery("#txt_name_" + uploadid).val();

                    if (hidden_check_file_name == "UC") {
                        tmpfiletitle_name.push(
                            jQuery("#hidden_current_file_name_" + uploadid).val()
                        );
                    } else {
                        tmpfiletitle_name.push(jQuery("#text_file_name_" + uploadid).val());
                    }

                    hidden_current_file_new = "N";

                    if (tmpfiletitle_name.length != 0) {
                        $("#div_showFileOrther63").css("display", "block");
                        $(".head_or").css("display", "block");
                        $("#div_showFileOrther63").removeClass("d-none");
                        $(".head_or").removeClass("d-none");
                        $("#DataNull59").css("display", "none");

                        setData += "<tr>";
                        setData += "<td >" + num++ + "</td>";
                        setData +=
                            "<td class='text-left'>" + tmpfiletitle_name[0] + "</td>";
                        if (app == "Y" || app == "y") {
                            if (hidden_current_file_new == "Y") {
                                setData +=
                                    "<td  class='link-table'><a id='FileDetailOpen_" +
                                    uploadid +
                                    "'><img src='/assets/images/document.svg' class='iconbook'></a ></td>";
                            } else {
                                setData +=
                                    "<td  class='link-table'><a><img src='/assets/images/document.svg' class='iconbook' onclick='sendmessage(&quot;" +
                                    arrayfile[0] +
                                    "&quot;)'></a ></td>";
                            }
                        } else {
                            if (hidden_current_file_new == "Y") {
                                setData +=
                                    "<td  class='link-table'><a onclick='GetDocFilePathByDocumentID(&quot;" +
                                    DocIDValue_file_name +
                                    "&quot;,&quot;" +
                                    hidden_form_type +
                                    "&quot;,&quot;" +
                                    hidden_form_id +
                                    "&quot;)' target='_blank'><img src='/assets/images/document.svg' class='iconbook'></a ></td>";
                            } else {
                                setData +=
                                    "<td  class='link-table'><a href='" +
                                    arrayfile[0] +
                                    "' target='_blank'><img src='/assets/images/document.svg' class='iconbook'></a ></td>";
                            }
                        }

                        setData += "</tr>";
                        $("#listFile_or63").html(setData);
                    } else {
                        $(".head_or").css("display", "none");
                        $("#div_showFileOrther63").css("display", "none");
                    }
                };
            } else {
                if (hidden_check_file_name == "UC") {
                    $("#err_" + uploadid).css("display", "inline");
                    checkorther = false;
                } else {
                    hidden_current_file_new = $(
                        "#hidden_current_file_new" + uploadid
                    ).val();
                    if (hidden_current_file_new == "Y") {
                        tmpfiletitle_name.push(show_file_title);
                    }
                }

                if (tmpfiletitle_name.length != 0) {
                    $("#div_showFileOrther63").css("display", "block");
                    $(".head_or").css("display", "block");
                    $("#div_showFileOrther63").removeClass("d-none");
                    $(".head_or").removeClass("d-none");
                    $("#DataNull59").css("display", "none");
                    setData += "<tr>";
                    setData += "<td >" + num++ + "</td>";
                    setData += "<td class='text-left'>" + tmpfiletitle_name[0] + "</td>";
                    if (app == "Y" || app == "y") {
                        if (hidden_current_file_new == "Y") {
                            setData +=
                                "<td  class='link-table'><a onclick='GetDocFilePathByDocumentID(&quot;" +
                                DocIDValue_file_name +
                                "&quot;,&quot;" +
                                hidden_form_type +
                                "&quot;,&quot;" +
                                hidden_form_id +
                                "&quot;)'><img src='/assets/images/document.svg' class='iconbook'></a ></td>";
                        } else {
                            setData +=
                                "<td  class='link-table'><a><img src='/assets/images/document.svg' class='iconbook' onclick='sendmessage(&quot;" +
                                arrayfile[0] +
                                "&quot;)'></a ></td>";
                        }
                    } else {
                        if (hidden_current_file_new == "Y") {
                            setData +=
                                "<td  class='link-table'><a onclick='GetDocFilePathByDocumentID(&quot;" +
                                DocIDValue_file_name +
                                "&quot;,&quot;" +
                                hidden_form_type +
                                "&quot;,&quot;" +
                                hidden_form_id +
                                "&quot;)' target='_blank'><img src='/assets/images/document.svg' class='iconbook'></a ></td>";
                        } else {
                            setData +=
                                "<td  class='link-table'><a href='" +
                                arrayfile[0] +
                                "' target='_blank'><img src='/assets/images/document.svg' class='iconbook'></a ></td>";
                        }
                    }

                    setData += "</tr>";

                    $("#listFile_or63").html(setData);
                } else {
                    $(".head_or").css("display", "none");
                    $("#div_showFileOrther63").css("display", "none");
                }
            }
        });
    }


    if (typeof requireOtherDocForNCOCRD === "function"
        && !requireOtherDocForNCOCRD(typeof statusGroup !== "undefined" ? statusGroup : "")) {
        return false;
    }

    if (chkData && allFieldsValid && checkorther) {
        current_fs = $(this).parent();
        next_fs = $(this).parent().next();
        next_fs.show();

        current_fs.animate(
            { opacity: 0 },
            {
                step: function (now) {

                    opacity = 1 - now;

                    current_fs.css({
                        display: "none",
                        position: "relative",
                    });
                    next_fs.css({ opacity: opacity });
                },
                duration: 500,
            }
        );
        setProgressBar(++current);
    }
});
