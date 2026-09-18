var arraycheck = [];
var lengthfile = [];
var successIconUrl = $("#successIcon").val();
var errorIconUrl = $("#errorIcon").val();
var statusGroup;
var transaction_id;
var currentfile = [];
var checkempfile = true;
var objString;
var roundedNumber;
var list_alien_id = [];
var agency_id_show = "";
var objData_setPayment;


async function GetDataNationalityInfo41() {
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
                    "<option value='' class='lang_opt_pleaseselect'>กรุณาเลือก</option>";
                for (var i = 0; i < objData.length; i++) {
                    setData +=
                        "<option value='" +
                        objData[i].nationality_id +
                        "' data_th='" +
                        objData[i].nationality_th +
                        "' data_en='" +
                        objData[i].nationality_en +
                        "'></option>";
                }
                jQuery("#nationality_id_mou").html(setData);
            }
        },
        "json"
    );
    await Promise.all([
        GetDataRequestFormEtrackingMOU41_4_59_ByGroupIDEdit(),
        changeDatalang(),
    ]);
}
async function GetDataDemandLetter() {
    var request = await jQuery.post(
        url_GetDataDemandLetter,
        null,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                populateDropdowns(objData[0].req_permit_max_of_days);
            }
        },
        "json"
    );
    await GetDataEmployerTypeInfo();
}
async function GetDataEmployerTypeInfo() {
    var request = await jQuery.post(
        url_GetDataEmployerTypeInfo,
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
                        objData[i].u_map_id +
                        "'>" +
                        objData[i].emp_type_name_th +
                        " : " +
                        objData[i].user_type_name_th +
                        "</option>";
                }
                jQuery("#employer_type").html(setData);
            }
        },
        "json"
    );
    await Promise.all([changeDatalang()]);
}
async function GetDataRequestFormEtrackingMOU41_4_59_ByGroupIDEdit() {
    $("#loading").css("display", "block");
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
        url_GetDataRequestFormEtrackingMOU41_4_59_ByGroupID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var dataObj = data.data[0];
                var objData = data.data[0];
                var demand_letter_data = dataObj.demand_letter_data;
                var objDataEmp = dataObj.demand_letter_data.employer_data;
                var agency_data = dataObj.demand_letter_data.agency_data;
                sessionStorage.setItem("agency_id", demand_letter_data.agency_id);

                $("#number_formid").html(
                    "<span class='lang_lbl_req_number'></span> : " + objData.group_id
                );
                $("#title_form").html(
                    "<span class='lang_lbl_list'></span> : <span style='color:#11111F;'><span data_th='" +
                    objData.form_type_name_th +
                    "' data_en='" +
                    objData.form_type_name_en +
                    "'></span></span>"
                );
                $("#text_home").html("<span class='lang_menu_home'></span> /")
                $("#text_index").html(
                    '<a class="text-secondary font-weight-normal" href="/Permit/Tracking"><span class="lang_track_request_status">ติดตามสถานะคำขอ</span></a> /'
                );
                $("#text_header").html("<span class='lang_lbl_req_number'></span> :" + " " + objData.group_id)
                $("#date_re").html(
                    "<span class='lang_lbl_submit_date'></span> : <span style='color:#11111F;'><span data_th='" +
                    convert_to_thai_datetime(objData.created_timestamp) +
                    "' data_en='" +
                    convert_to_eng_datetime(objData.created_timestamp) +
                    "'></span></span>"
                );
                $("#form_activity_detail").attr(
                    "data_th",
                    objData.transection_status_th
                );
                $("#form_activity_detail").attr(
                    "data_en",
                    objData.transection_status_en
                );

                var arraystatus_cc = [
                    "CC",
                    "CCOS",
                    "CCOS1",
                    "CCOS2",
                    "NCOC",
                    "NCOC1",
                    "NCOC2",
                    "NCOS",
                    "NCOS1",
                    "NCOS2",
                    "NCOCR1",
                    "NCOCRD1",
                    "NCOCR2",
                    "NCOCRD2",
                ];
                var arraystatus_ss = ["SS"];
                var arraystatus_ap = ["APSS"];

                $("#form_activity_detail").addClass("badgestatus");
                $("#form_activity_detail").css({
                    color: objData.status_color_font,
                    background: objData.status_color,
                });
                if (arraystatus_cc.includes(objData.group_status_id)) {
                    $("#divforreasonshowdetail").show();
                } else if (arraystatus_ss.includes(objData.group_status_id)) {
                    $("#line_status").hide();
                } else if (arraystatus_ap.includes(objData.group_status_id)) {
                    $("#line_status").hide();
                } else {
                    $("#line_status").hide();
                    $("#divforreasonshowdetail").hide();
                }

                var setData = "";
                setData +=
                    '<option class="lang_opt_pleaseselect" selected="" value="">กรุณาเลือก</option>';
                for (var i = 0; i < objDataEmp.addrWrkList.length; i++) {
                    var StringData = JSON.stringify(
                        objDataEmp.addrWrkList[i].employer_addr_id
                    );
                    var str_addr_bus = JSON.stringify(
                        objDataEmp.addrWrkList[i].emp_addr_bus_type_list
                    );
                    var str_addr = JSON.stringify(objDataEmp.addrWrkList[i]);

                    setData +=
                        "<option value='" +
                        objDataEmp.addrWrkList[i].employer_addr_id +
                        "' data-value-add='" +
                        str_addr_bus +
                        "' data-address='" +
                        str_addr +
                        "'  data_th='" +
                        objDataEmp.addrWrkList[i].employer_full_address_th +
                        "' data_en='" +
                        objDataEmp.addrWrkList[i].employer_full_address_en +
                        "'></option>";
                }
                $("#name_address_work").html(setData);

                if (objDataEmp.addrWrkList.length == 1) {
                    $("#name_address_work").prop("disabled", true);
                    $("#name_address_work").prop("selectedIndex", 1);
                } else {
                    $("#name_address_work").prop("disabled", false);
                }
                if (
                    demand_letter_data.workplace_addr_id != "" ||
                    demand_letter_data.workplace_addr_id != null ||
                    demand_letter_data.workplace_addr_id != undefined
                ) {
                    $("#name_address_work").val(demand_letter_data.workplace_addr_id);
                }

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
                            "' data_en='" +
                            obj_bus[j].bus_type_name_en +
                            "'></option>";
                    }
                    $("#bus_type_id").html(setdatabus);


                }

                $("#bus_type_id").val(demand_letter_data.bus_type_id);

                $("#employer_id").val(objDataEmp.employer_id);
                $("#emp_type_id").val(objDataEmp.emp_type_id);
                $("#emp_type_name_th").val(objDataEmp.emp_type_name_th);
                $("#user_id").val(objDataEmp.user_id);
                $("#group_id").val(dataObj.group_id);


                $(".juristic_permit_no_edit,.juristic_permit_no_edit_st3").text(
                    agency_data.juristic_permit_no || "-"
                );
                if (agency_data.juristic_permit_issue_dt != "") {
                    $(
                        ".juristic_permit_issue_dt_edit,.juristic_permit_issue_dt_edit_st3"
                    ).attr(
                        "data_th",
                        convert_to_thai_date(agency_data.juristic_permit_issue_dt)
                    );
                    $(
                        ".juristic_permit_issue_dt_edit,.juristic_permit_issue_dt_edit_st3"
                    ).attr(
                        "data_en",
                        convert_to_eng_date(agency_data.juristic_permit_issue_dt)
                    );
                } else {
                    $(
                        ".juristic_permit_issue_dt_edit,.juristic_permit_issue_dt_edit_st3"
                    ).addClass("lang_not_specified");
                }

                if (agency_data.juristic_permit_until_dt != "") {
                    $(
                        ".juristic_permit_until_dt_edit,.juristic_permit_until_dt_edit_st3"
                    ).attr(
                        "data_th",
                        convert_to_thai_date(agency_data.juristic_permit_until_dt)
                    );
                    $(
                        ".juristic_permit_until_dt_edit,.juristic_permit_until_dt_edit_st3"
                    ).attr(
                        "data_en",
                        convert_to_eng_date(agency_data.juristic_permit_until_dt)
                    );
                } else {
                    $(
                        ".juristic_permit_until_dt_edit,.juristic_permit_until_dt_edit_st3"
                    ).addClass("lang_not_specified");
                }

                $("#name_address_work").val(demand_letter_data.workplace_addr_id);
                $("#hidden_bus_type_text").val(demand_letter_data.bus_type_name_th);
                $("#start_age").val(demand_letter_data.start_age);
                $("#end_age").val(demand_letter_data.end_age);
                $("#start_height").val(demand_letter_data.start_height);
                $("#end_height").val(demand_letter_data.end_height);
                $("#start_weight").val(demand_letter_data.start_weight);
                $("#end_weight").val(demand_letter_data.end_weight);


                $("#workhour_per_day").val(demand_letter_data.workhour_per_day);
                $("#dayoff_per_week").val(demand_letter_data.dayoff_per_week);
                $("#holiday_per_year").val(demand_letter_data.holiday_per_year);
                $("#wage_rate_per_day").val(demand_letter_data.wage_rate_per_day);
                OnchangeYear(
                    demand_letter_data.work_duration_year,
                    demand_letter_data.work_duration_month,
                    demand_letter_data.work_duration_day
                );
                $("#number_of_alien_all").val(demand_letter_data.number_of_alien_all);
                $("#number_of_alien_male").val(demand_letter_data.number_of_alien_male);
                $("#number_of_alien_female").val(
                    demand_letter_data.number_of_alien_female
                );
                $("#traditional_holiday").val(demand_letter_data.traditional_holiday);
                $("#issale").val(objData.demand_letter_data.is_saleperson);
                localStorage.setItem("demand_id", demand_letter_data.demand_letter_id);
                var demand_letter_job_list =
                    demand_letter_data.demand_letter_job_list[0];
                $("#checkboxjob_" + demand_letter_job_list.permit_cate_id).prop(
                    "checked",
                    true
                );
                var checkboxSelector =
                    "#checkboxjob_" + demand_letter_job_list.permit_cate_id;

                if ($(checkboxSelector).length === 0) {
                    $("#checkboxjob_4").prop("checked", true);
                    $("#numbermen_checkboxjob_4").val(
                        demand_letter_job_list.number_of_alien_male
                    );
                    $("#numberwomen_checkboxjob_4").val(
                        demand_letter_job_list.number_of_alien_female
                    );

                    $("#numbermen_checkboxjob_4").removeAttr("disabled");
                    $("#numberwomen_checkboxjob_4").removeAttr("disabled");
                    $("#div_ortherjob").css("display", "block");
                    $("#permit_catework_Dropdown").val(
                        demand_letter_job_list.permit_cate_id
                    );
                    $("#permit_catework_Dropdown").removeAttr("disabled");
                } else {
                    $("#checkboxjob_" + demand_letter_job_list.permit_cate_id).prop(
                        "checked",
                        true
                    );
                    $(
                        "#numbermen_checkboxjob_" + demand_letter_job_list.permit_cate_id
                    ).val(demand_letter_job_list.number_of_alien_male);
                    $(
                        "#numberwomen_checkboxjob_" + demand_letter_job_list.permit_cate_id
                    ).val(demand_letter_job_list.number_of_alien_female);
                    $(
                        "#numbermen_checkboxjob_" + demand_letter_job_list.permit_cate_id
                    ).removeAttr("disabled");
                    $(
                        "#numberwomen_checkboxjob_" + demand_letter_job_list.permit_cate_id
                    ).removeAttr("disabled");
                }

                var objdataformlist = dataObj.demand_letter_data;
                var checkreason = true;
                var checkreason_emp = true;
                var checkreasonMore = true;
                let reasonHtml = "";
                let reasonHtml_emp = "";
                let reasonHtml_proxy = "";


                if (objdataformlist.consider_data.employer_reasonlist == null) {
                    $(".reason_alien").addClass("d-none");
                    checkreason = true;
                }
                if (objdataformlist.consider_data.agency_reasonlist == null) {
                    $(".reason_emp").addClass("d-none");
                    checkreason_emp = true;
                }
                if (objdataformlist.consider_data.proxy_reasonlist == null) {
                    $(".reason_proxy").addClass("d-none");
                    checkreason_emp = true;
                }

                if (objdataformlist.consider_data.employer_reasonlist != null) {
                    var consider_data_employer_reasonlist =
                        objdataformlist.consider_data.employer_reasonlist
                            .checkdocemp_reason_list;

                    var empNeedEdit = objdataformlist.consider_data.employer_reasonlist.checkdocemp == "N";
                    if ((consider_data_employer_reasonlist != null && consider_data_employer_reasonlist.length > 0) || empNeedEdit) {
                        $(".reason_alien").removeClass("d-none");
                        $(".listreason_alien").removeClass("d-none");
                        $(".text_reason_alien").addClass("lang_list_of_reasons_emp");
                        if (consider_data_employer_reasonlist != null) {
                            for (var v = 0; v < consider_data_employer_reasonlist.length; v++) {
                                reasonHtml +=
                                    '<li data_th="' +
                                    consider_data_employer_reasonlist[v].reason_checkdoc_name_th +
                                    '" data_en="' +
                                    consider_data_employer_reasonlist[v].reason_checkdoc_name_en +
                                    '"></li>';
                            }
                        }
                        if (empNeedEdit) {
                            reasonHtml += '<li style="color:#E24C86;" data_th="กรุณาแก้ไขเอกสารนายจ้าง" data_en="Please edit the employer\'s documents" data_vn="Vui lòng chỉnh sửa tài liệu của người sử dụng lao động" data_lo="ກະລຸນາແກ້ໄຂເອກະສານນາຍຈ້າງ" data_mm="ကျေးဇူးပြု၍ အလုပ်ရှင်၏ စာရွက်စာတမ်းများကို ပြင်ဆင်ပါ" data_cb="សូមកែសម្រួលឯកសាររបស់និយោជក"></li>';
                        }
                        checkreason = false;
                    }
                }

                if (objdataformlist.consider_data.agency_reasonlist != null) {
                    var consider_data_agency_reasonlist =
                        objdataformlist.consider_data.agency_reasonlist
                            .checkdocagency_reason_list;

                    var agencyNeedEdit = objdataformlist.consider_data.agency_reasonlist.checkdocagency == "N";

                    if ((consider_data_agency_reasonlist != null && consider_data_agency_reasonlist.length > 0) || agencyNeedEdit) {
                        $(".reason_emp").removeClass("d-none");
                        $(".listreason_emp").removeClass("d-none");
                        $(".text_reason_emp").addClass("lang_list_of_reasons_import");
                        if (consider_data_agency_reasonlist != null) {
                            for (var v = 0; v < consider_data_agency_reasonlist.length; v++) {
                                reasonHtml_emp +=
                                    '<li data_th="' +
                                    consider_data_agency_reasonlist[v].reason_checkdoc_name_th +
                                    '" data_en="' +
                                    consider_data_agency_reasonlist[v].reason_checkdoc_name_en +
                                    '"></li>';
                            }
                        }
                        if (agencyNeedEdit) {
                            reasonHtml_emp += '<li style="color:#E24C86;" data_th="กรุณาแก้ไขเอกสาร บนจ." data_en="Please edit the recruitment agency\'s documents" data_vn="Vui lòng chỉnh sửa tài liệu của công ty tuyển dụng" data_lo="ກະລຸນາແກ້ໄຂເອກະສານບໍລິສັດຈັດຫາງານ" data_mm="ကျေးဇူးပြု၍ အလုပ်အကိုင်ရှာဖွေရေးအေဂျင်စီ၏ စာရွက်စာတမ်းများကို ပြင်ဆင်ပါ" data_cb="សូមកែសម្រួលឯកសាររបស់ភ្នាក់ងារជ្រើសរើសបុគ្គលិក"></li>';
                        }
                        checkreason = false;
                    }
                }

                if (objdataformlist.consider_data.proxy_reasonlist != null) {
                    var consider_data_proxy_reasonlist =
                        objdataformlist.consider_data.proxy_reasonlist
                            .checkdocproxy_reason_list;

                    var proxyNeedEdit = objdataformlist.consider_data.proxy_reasonlist.checkdocproxy == "N";
                    if ((consider_data_proxy_reasonlist != null && consider_data_proxy_reasonlist.length > 0) || proxyNeedEdit) {
                        $(".reason_proxy").removeClass("d-none");
                        $(".listreason_proxy").removeClass("d-none");
                        $(".text_reason_proxy").addClass("lang_reasons_representative");
                        if (consider_data_proxy_reasonlist != null) {
                            for (var v = 0; v < consider_data_proxy_reasonlist.length; v++) {
                                reasonHtml_proxy +=
                                    '<li data_th="' +
                                    consider_data_proxy_reasonlist[v].reason_checkdoc_name_th +
                                    '" data_en="' +
                                    consider_data_proxy_reasonlist[v].reason_checkdoc_name_en +
                                    '"></li>';
                            }
                        }
                        if (proxyNeedEdit) {
                            reasonHtml_proxy += '<li style="color:#E24C86;" data_th="กรุณาแก้ไขเอกสารผู้ดำเนินการแทน" data_en="Please edit the representative\'s documents" data_vn="Vui lòng chỉnh sửa tài liệu của người đại diện" data_lo="ກະລຸນາແກ້ໄຂເອກະສານຜູ້ດຳເນີນການແທນ" data_mm="ကျေးဇူးပြု၍ ကိုယ်စားလှယ်၏ စာရွက်စာတမ်းများကို ပြင်ဆင်ပါ" data_cb="សូមកែសម្រួលឯកសាររបស់អ្នកតំណាង"></li>';
                        }
                        checkreason = false;
                    }
                }

                var reasonHtmlMore = "";


                if (objdataformlist.consider_data.agency_reasonlist != null) {
                    if (
                        objdataformlist.consider_data.agency_reasonlist
                            .checkdocagency_reason != ""
                    ) {
                        $(".diveressonmore").removeClass("d-none");
                        reasonHtmlMore += `<li><b><span class='lang_authorized_person'></span> :</b>${objdataformlist.consider_data.agency_reasonlist.checkdocagency_reason}</li>`;
                        checkreasonMore = false;
                    }


                }
                if (objdataformlist.consider_data.employer_reasonlist != null) {
                    if (
                        objdataformlist.consider_data.employer_reasonlist
                            .checkdocemp_reason != ""
                    ) {

                        $(".diveressonmore").removeClass("d-none");
                        reasonHtmlMore += `<li><b><span class='lang_lbl_employer'></span> : </b>${objdataformlist.consider_data.employer_reasonlist.checkdocemp_reason}</li>`;
                        checkreasonMore = false;
                    }
                }

                if (objdataformlist.consider_data.proxy_reasonlist != null) {
                    if (
                        objdataformlist.consider_data.proxy_reasonlist
                            .checkdocproxy_reason != ""
                    ) {
                        $(".diveressonmore").removeClass("d-none");
                        reasonHtmlMore += `<li><b><span class='lang_employer_representative'>นายจ้าง (ผู้ดำเนินการเเทน)</span> : </b>${objdataformlist.consider_data.proxy_reasonlist.checkdocproxy_reason}</li>`;
                        checkreasonMore = false;
                    }


                }
         
                if (typeof applyOtherDocDocRequest === "function") {
                    applyOtherDocDocRequest(objdataformlist.consider_data);
                }

                if (objdataformlist.consider_data.docrequestlist.docrequest == "Y") {
                    $(".diveressonmore").removeClass("d-none");
                    reasonHtmlMore += `<li><b><span class='lang_registrar_documents'>ขอเอกสารเพิ่มเติม</span> : </b>${objdataformlist.consider_data.docrequestlist.docrequest_reason}</li>`;
                    checkreasonMore = false;
                }

                $("#divShowreason").html(reasonHtml);
                $("#divShowreason_emp").html(reasonHtml_emp);
                $("#divShowreasonMore").html(reasonHtmlMore);
                $("#divShowreason_proxy").html(reasonHtml_proxy);
                if (
                    checkreason == true &&
                    checkreasonMore == true &&
                    checkreason_emp == true
                ) {
                    $(".hr_shoereason").addClass("d-none");
                }


                $("#nationality_id_mou").val(demand_letter_data.nationality_id);
                GetDataNationalityInfoEdit(
                    demand_letter_data.nationality_id,
                    demand_letter_data.recruitment_agency_id,
                    demand_letter_data.fsc_id,
                    demand_letter_data.immigration_checkpoint_id
                );


                $("#address_recruitment").val(
                    demand_letter_data.recruitment_agency_address
                );

                var recruitment_agency_assign_at =
                    demand_letter_data.recruitment_agency_assign_at;
                var sp_recruitment_agency_assign_at =
                    recruitment_agency_assign_at.split(" ");
                $("#recruitment_agency_assign_at").val(
                    sp_recruitment_agency_assign_at[0]
                );

                $("#employer_type").val(objDataEmp.u_map_id);
                $("#employer_type").attr("disabled", true);
                $("#employerID_63").val(objDataEmp.employer_no);
                $("#employerID_63").attr("disabled", true);
                $("#div_DataEmployer").css("display", "flex");

                if (objDataEmp.emp_type_id == "1") {
                    $(".name_th_employer").addClass("lang_txt_employer_name_th");
                    $(".name_en_employer").addClass("lang_txt_employer_name_eng");

                    if (objDataEmp.firstname_th) {
                        $(".full_name_th").text(
                            objDataEmp.prefix_name_th +
                            objDataEmp.firstname_th +
                            " " +
                            objDataEmp.middlename_th +
                            " " +
                            objDataEmp.lastname_th
                        );
                    } else {
                        $(".full_name_th").text("-");
                    }

                    if (objDataEmp.firstname_en) {
                        $(".full_name_en").text(
                            objDataEmp.prefix_name_en +
                            objDataEmp.firstname_en +
                            " " +
                            objDataEmp.middlename_en +
                            " " +
                            objDataEmp.lastname_en
                        );
                    } else {
                        $(".full_name_en").text("-");
                    }
                    if (objDataEmp.u_map_id == "8" && objDataEmp.user_type_id == "4") {
                        $(".text_employer_type").addClass("lang_foreign_individual_label");
                        $(".no_employer").addClass("lang_tax_id");
                    } else {
                        $(".text_employer_type").attr(
                            "data_th",
                            objDataEmp.emp_type_name_th
                        );
                        $(".text_employer_type").attr(
                            "data_en",
                            objDataEmp.emp_type_name_en
                        );
                        $(".no_employer").addClass("lang_id_card_and_legal_entity_number");
                    }
                } else if (objDataEmp.emp_type_id == "2") {
                    $(".name_th_employer").addClass("lang_bu_name_th");
                    $(".name_en_employer").addClass("lang_bu_name_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);

                    if (objDataEmp.u_map_id == "9" && objDataEmp.user_type_id == "4") {
                        $(".text_employer_type").addClass("lang_registered_in_thailand");
                        $(".no_employer").addClass("lang_tax_id");
                    } else {
                        $(".no_employer").addClass("lang_id_card_and_legal_entity_number");
                        $(".text_employer_type").attr(
                            "data_th",
                            objDataEmp.emp_type_name_th
                        );
                        $(".text_employer_type").attr(
                            "data_en",
                            objDataEmp.emp_type_name_en
                        );
                    }
                } else if (
                    objDataEmp.emp_type_id == "4" ||
                    objDataEmp.emp_type_id == "5" ||
                    objDataEmp.emp_type_id == "6"
                ) {
                    $(".no_employer").addClass("lang_government_no");
                    $(".name_th_employer").addClass("lang_name_organization_th");
                    $(".name_en_employer").addClass("lang_name_organization_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);
                    $(".text_employer_type").attr("data_th", objDataEmp.emp_type_name_th);
                    $(".text_employer_type").attr("data_en", objDataEmp.emp_type_name_en);
                } else if (objDataEmp.emp_type_id == "8") {
                    $(".no_employer").addClass("lang_government_no");
                    $(".name_th_employer").addClass("lang_bu_name_th");
                    $(".name_en_employer").addClass("lang_bu_name_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);
                    $(".text_employer_type").attr("data_th", objDataEmp.emp_type_name_th);
                    $(".text_employer_type").attr("data_en", objDataEmp.emp_type_name_en);
                } else if (objDataEmp.emp_type_id == "7") {
                    $(".no_employer").addClass("lang_government_no");
                    $(".name_th_employer").addClass("lang_company_th");
                    $(".name_en_employer").addClass("lang_company_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);
                    $(".text_employer_type").attr("data_th", objDataEmp.emp_type_name_th);
                    $(".text_employer_type").attr("data_en", objDataEmp.emp_type_name_en);
                } else if (objDataEmp.emp_type_id == "9") {
                    $(".no_employer").addClass("lang_tax_id");
                    $(".name_th_employer").addClass("lang_bu_name_th");
                    $(".name_en_employer").addClass("lang_bu_name_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);
                    if (objDataEmp.u_map_id == "10" && objDataEmp.user_type_id == "4") {
                        $(".text_employer_type").addClass("lang_registered_abroad");
                    } else {
                        $(".text_employer_type").attr(
                            "data_th",
                            objDataEmp.emp_type_name_th
                        );
                        $(".text_employer_type").attr(
                            "data_en",
                            objDataEmp.emp_type_name_en
                        );
                    }
                }


                if (objdataformlist.proxy_id) {
                    $(".div_agent").removeClass("d-none");
                    var dataproxy = objdataformlist.proxy_data;
                    if (dataproxy.firstname_th) {
                        $(".text_namethai_agent").text(
                            dataproxy.prefix_name_th +
                            dataproxy.firstname_th +
                            " " +
                            dataproxy.midname_th +
                            " " +
                            dataproxy.lastname_th
                        );
                    } else {
                        $(".text_namethai_agent").text("-");
                    }

                    if (dataproxy.firstname_en) {
                        $(".text_nameeng_agent").text(
                            dataproxy.prefix_name_en +
                            dataproxy.firstname_en +
                            " " +
                            dataproxy.midname_en +
                            " " +
                            dataproxy.lastname_en
                        );
                    } else {
                        $(".text_nameeng_agent").text("-");
                    }

                    $(".text_idcard_agent").text(dataproxy.id_card);
                }
                $(".text_employer_idcard").text(objDataEmp.employer_id);
                $("#text_employer_idcard").text(objDataEmp.employer_id);
                $(".text_employer_id").text(objDataEmp.employer_id);
                $(".text_employer_no").text(objDataEmp.employer_no);

                $(".email_em").text(objDataEmp.email || "-");
                $(".lbl_fax_em").text(objDataEmp.fax || "-");
                if (objDataEmp.addrReqList.length != 0) {
                    $(".addrReqList_emp").text(
                        objDataEmp.addrReqList[0].employer_full_address_th
                    );
                } else {
                    $(".addrReqList_emp").text("-");
                }

                $(".tel_em").text(objDataEmp.phone || "-");
                $(".email").text(agency_data.email || "-");
                $(".tel_em").text(agency_data.phone || "-");
                $(".lbl_fax").text(agency_data.fax || "-");
                $(".text_employer_type41").attr(
                    "data_th",
                    agency_data.emp_type_name_th
                );
                $(".text_employer_type41").attr(
                    "data_en",
                    agency_data.emp_type_name_en
                );
                $(".agency_no").text(agency_data.agency_no);
                $(".company_thai").text(agency_data.companyname_th);
                $(".company_eng").text(agency_data.companyname_en);

                $(".company_regis_date").attr(
                    "data_th",
                    convert_to_thai_date(agency_data.company_regis_date)
                );
                $(".company_regis_date").attr(
                    "data_en",
                    convert_to_eng_date(agency_data.company_regis_date)
                );

                $(".company_objt").text(agency_data.company_objt);
                $(".company_regis_capital").text(agency_data.company_regis_capital);
                if (agency_data.addrWrkList.length > 0) {
                    $(".fulladdress_th").text(agency_data.addrWrkList[0].fulladdress_th);
                }
                sessionStorage.setItem(
                    "request_by_user_id",
                    dataObj.request_by_user_id
                );


                if (objDataEmp.addrCurList.length > 0) {
                    $(".textaddress_em").attr(
                        "data_th",
                        objDataEmp.addrCurList[0].employer_full_address_th
                    );
                    $(".textaddress_em").attr(
                        "data_en",
                        objDataEmp.addrCurList[0].employer_full_address_en
                    );
                }

                $(".text_employer_type41").attr(
                    "data_th",
                    agency_data.user_type_name_th
                );
                $(".text_employer_type41").attr(
                    "data_en",
                    agency_data.user_type_name_en
                );
                $(".company_type_id").text(agency_data.company_type_id);
                $(".company_status_id").text(agency_data.company_status_id);
                $(".company_id").text(agency_data.company_id);
                $(".company_regis_capital").text(agency_data.company_regis_capital);
                $(".director_fullname").html(
                    agency_data.director_fullname.replace(",", "<br>")
                );
            } else {
                $("#loading").delay().fadeOut("");
                alert(data.msg);
            }
        },
        "json"
    );
    $("#loading").delay().fadeOut("");
    await Promise.all([
        OnloadFileMOUEdit(),
        onloadfileDetail(),
        changeDatalang(),
    ]);
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
function GetDataImmigrationCheckpoint(
    value_nationality,
    immigration_checkpoint_id
) {
    $("#DataImmigrationCheckpoint").removeAttr("disabled");
    $("#fsc_id").removeAttr("disabled");

    let objRequest = {
        nationality_id: value_nationality,
    };

    var request = jQuery.post(
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
                            objData[i].immigration_checkpoint_name_th +
                            "' data_en='" +
                            objData[i].immigration_checkpoint_name_en +
                            "'></option>";
                    }
                }

                jQuery("#DataImmigrationCheckpoint").html(setData);
                if (immigration_checkpoint_id != "") {
                    $("#DataImmigrationCheckpoint").val(immigration_checkpoint_id);
                }
                changeDatalang();


            }
        },
        "json"
    );
}
function OncGetDataImmigrationCheckpoint(value_nationality) {
    $("#DataImmigrationCheckpoint").removeAttr("disabled");
    $("#fsc_id").removeAttr("disabled");

    let objRequest = {
        nationality_id: value_nationality,
    };

    var request = jQuery.post(
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
                            objData[i].immigration_checkpoint_name_th +
                            "' data_en='" +
                            objData[i].immigration_checkpoint_name_en +
                            "'></option>";
                    }
                }

                jQuery("#DataImmigrationCheckpoint").html(setData);

                GetDataFirstServiceCenter(value_nationality);
                GetDataRecruitmentAgency(value_nationality);
                changeDatalang();
            }
        },
        "json"
    );
}
async function GetDataFirstServiceCenter(value_nationality, fsc_id) {
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
                            "' data_en='" +
                            objData[i].fsc_name_en +
                            "'></option>";
                    }
                }

                jQuery("#fsc_id").html(setData);
                $("#fsc_id").val(fsc_id);
            }
        },
        "json"
    );
    await changeDatalang();
}
function NextStepTwoPageOneMOU() {
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
            OnloadFileSellEdit();
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
            showdetailedit();
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
function OnloadFileMOUEdit() {
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
    var request = jQuery.post(
        url_GetDataRequestFormEtrackingMOU41_4_59_ByGroupID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                $("#loading").delay().fadeOut("");
                var objData = data.data[0];
                var employer_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.employer_n_doclist;
                var saleperson_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.saleperson_doclist;
                var agency_n_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.agency_n_doclist;
                var setData = "";
                var index_array = 0;
                agency_id_show = objData.demand_letter_data.agency_id;


                var setData_agency = "";

                var agency_n_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.agency_n_doclist;
                for (let c = 0; c < agency_n_doclist.length; c++) {
                    $("#uploadfile_emp").removeClass("d-none");
                    $(".div_uploadfile_emp").removeClass("d-none");

                    const requestDocList = agency_n_doclist[c].requestdoclist;
                    setData_agency += "<div class='col-md-12'>";
                    if (requestDocList.length >= 2) {
                        setData_agency += `<p class='text-upload'>${c + 1
                            }${"."}<span class='lang_select_at_least_one_file'></span>`;
                        if (agency_n_doclist[c].required == "Y") {
                            setData_agency += "<span class='asterisk'>*</span>";
                        }
                        setData_agency +=
                            "<span id='err_file_name_" +
                            agency_n_doclist[c].group_form_type_doc_id +
                            "' class='lang_please_add_file text-err'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                    }

                    for (let j = 0; j < requestDocList.length; j++) {
                        if (requestDocList.length < 2) {
                            setData_agency += "<div>";
                            setData_agency += `<p class='text-upload'>${c + 1
                                }${"."}<span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()
                                }' data_en='${requestDocList[j].document_name_en}'></span>`;
                            if (agency_n_doclist[c].required == "Y") {
                                setData_agency += "<span class='asterisk'>*</span>";
                            }
                            setData_agency +=
                                "<span id='err_file_name_" +
                                requestDocList[j].document_type_id +
                                "' class='lang_please_add_file text-err'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                        } else {
                            setData_agency += '<div style="margin-left:15px;">';

                            setData_agency += `<p class='text-upload'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_th}'></span>`;
                        }

                        setData_agency += "</p>";
                        setData_agency += "</p>";
                        setData_agency += "<div class='row'>";
                        setData_agency += "<div class='col-lg-12'>";
                        setData_agency +=
                            "<p class='text-upload-subtext lang_support_pdf'>รองรับไฟล์ .PDF .PNG เเละ .JPG ขนาดไม่เกิน 25 mb</p>";
                        setData_agency += "<div class='row mb-3' style='margin-left:0px;'>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "C") {
                                setData_agency +=
                                    "<div class='upload d-none' style='padding-right:14px;'>";
                            } else {
                                setData_agency +=
                                    "<div class='upload' style='padding-right:14px;'>";
                            }
                        } else {
                            setData_agency +=
                                "<div class='upload' style='padding-right:14px;'>";
                        }

                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList.length >= 2) {
                                setData_agency +=
                                    "<input type='file' name='file_name_group' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;,&quot;" +
                                    requestDocList[j].doclist[0].document_id +
                                    "&quot;)' hidden/>";
                                setData_agency +=
                                    "<input   id='file_name_check_" +
                                    requestDocList[j].document_type_id +
                                    "' onclick='checknewfile(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            } else {
                                setData_agency +=
                                    "<input type='file' name='file_name' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;,&quot;" +
                                    requestDocList[j].doclist[0].document_id +
                                    "&quot;)' hidden/>";
                                setData_agency +=
                                    "<input  id='file_name_check_" +
                                    requestDocList[j].document_type_id +
                                    "' onclick='checknewfile(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            }
                        } else {
                            if (requestDocList.length >= 2) {
                                setData_agency +=
                                    "<input type='file' name='file_name_group' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                                setData_agency +=
                                    "<input  id='file_name_check_" +
                                    requestDocList[j].document_type_id +
                                    "' onclick='checknewfile(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            } else {
                                setData_agency +=
                                    "<input type='file' name='file_name' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                                setData_agency +=
                                    "<input  id='file_name_check_" +
                                    requestDocList[j].document_type_id +
                                    "' onclick='checknewfile(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            }
                        }

                        setData_agency +=
                            "<label for='file_name_check_" +
                            requestDocList[j].document_type_id +
                            "' style='cursor:pointer;' class='lang_lbl_selectfile'>เลือกไฟล์</label>";


                        setData_agency += "</div>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                setData_agency +=
                                    "<div class='block-doc justify-content-between border_file_reject' id='divshow_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "'>";
                            } else {
                                setData_agency +=
                                    "<div class='block-doc justify-content-between border_file_success' id='divshow_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "'>";
                            }
                        } else {
                            setData_agency +=
                                "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                requestDocList[j].document_type_id +
                                "' style='display:none;'>";
                        }
                        setData_agency += "<div class='d-flex'>";
                        setData_agency += "<div>";
                        setData_agency += "<img src='/assets/images/document.svg'>";
                        setData_agency += "</div>";
                        setData_agency += "<div class='filename'>";
                        if (requestDocList[j].doclist.length != 0) {
                            setData_agency +=
                                "<a id='taga_" +
                                requestDocList[j].document_type_id +
                                "' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                objData.demand_letter_data.agency_id +
                                " &quot;, &quot; " +
                                requestDocList[j].doclist[0].document_id +
                                " &quot;, &quot; " +
                                requestDocList[j].document_type_id +
                                " &quot;)'><span class='name-file namedoc' id='show_file_name_" +
                                requestDocList[j].document_type_id +
                                "'>" +
                                requestDocList[j].doclist[0].doc_title +
                                "</span></a>";

                            setData_agency +=
                                '<input type="hidden" id="DocIDValue_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].document_id +
                                '">';
                        }
                        setData_agency +=
                            "<p class='name-file namedoc' id='show_file_name_" +
                            requestDocList[j].document_type_id +
                            "'></p>";
                        setData_agency += "</div>";

                        setData_agency += "<div>";
                        if (requestDocList[j].doclist.length != 0) {


                        } else {
                            setData_agency +=
                                "<img class='close' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                requestDocList[j].document_type_id +
                                "&quot;)'>";
                        }

                        setData_agency += "</div>";
                        setData_agency += "</div>";
                        setData_agency += "</div>";
                        setData_agency += "</div>";
                        setData_agency += "</div>";
                        setData_agency += "</div>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                setData_agency +=
                                    "<div class='resson' id='resson_" +
                                    requestDocList[j].document_type_id +
                                    "'><p class='textresson lang_text_reason'>หมายเหตุ</p><p> " +
                                    requestDocList[j].doclist[0].doc_comment +
                                    " </p></div>";
                                setData_agency +=
                                    '<input type="hidden" id="hidden_check_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="' +
                                    requestDocList[j].doclist[0].doc_status_id +
                                    '">';

                                if (requestDocList[j].doclist[0].doc_new_lastest == "Y") {
                                    setData_agency +=
                                        "<br><div class='show_new_emp_created_timestamp_" +
                                        requestDocList[j].document_type_id +
                                        " d-none'><span style='color:#535F6B;'>อัปเดตล่าสุด : <span data_th='" +
                                        convert_to_thai_datetime(
                                            requestDocList[j].doclist[0].new_emp_created_timestamp
                                        ) +
                                        "' data_en='" +
                                        convert_to_eng_datetime(
                                            requestDocList[j].doclist[0].new_emp_created_timestamp
                                        ) +
                                        "'></span></span> <span id='form_activity_detail' class='badge badge-pill statusAPSS2'>เอกสารอัปเดตใหม่</span></div>";
                                }
                            }
                        }

                        setData_agency += "</div>";
                        setData_agency +=
                            '<input type="hidden" id="requiredvalue_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            agency_n_doclist[c].required +
                            '">';
                        setData_agency +=
                            '<input type="hidden" id="hidden_form_id_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            objData.demand_letter_data.demand_letter_id +
                            '">';
                        setData_agency +=
                            '<input type="hidden" id="hidden_form_type_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            requestDocList[j].form_type_doc_id +
                            '">';
                        setData_agency +=
                            '<input type="hidden" id="required_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            agency_n_doclist[c].group_form_type_doc_id +
                            '">';

                        if (requestDocList[j].doclist.length != 0) {
                            setData_agency +=
                                '<input type="hidden" id="have_file_current_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].doclist[0].document_id +
                                '">';
                            setData_agency +=
                                '<input type="hidden" id="hidden_select_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="">';
                            setData_agency +=
                                '<input type="hidden" id="doc_new_lastest_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].doclist[0].doc_new_lastest +
                                '">';

                            if (requestDocList[j].doclist[0].doc_new_lastest == "Y") {
                                setData_agency +=
                                    '<input type="hidden" id="new_emp_created_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="' +
                                    requestDocList[j].doclist[0].new_emp_created_timestamp +
                                    '">';

                                objString.push(requestDocList[j]);
                            }
                        } else {
                            setData_agency +=
                                '<input type="hidden" id="have_file_current_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="N">';
                        }
                    }

                    setData_agency += "</div>";
                    setData_agency += "<div class='col-md-12 mb-2 mt-2'>";
                    setData_agency += "<hr>";
                    setData_agency += "</div>";
                }
                jQuery("#uploadfile_agency_n_doclist").html(setData_agency);

                var setDataEmpUp = "";
                var agency_y_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.agency_y_doclist;
                for (let c = 0; c < agency_y_doclist.length; c++) {
                    $("#uploadfile_emp").removeClass("d-none");
                    $(".div_uploadfile_emp").removeClass("d-none");

                    const requestDocList = agency_y_doclist[c].requestdoclist;
                    setDataEmpUp += "<div class='col-md-12'>";
                    if (requestDocList.length >= 2) {
                        setDataEmpUp += `<p class='text-upload'>${c + 1
                            }${"."}<span class='lang_select_at_least_one_file'></span>`;
                        if (agency_y_doclist[c].required == "Y") {
                            setDataEmpUp += "<span class='asterisk'>*</span>";
                        }
                        setDataEmpUp +=
                            "<span id='err_file_name_" +
                            agency_y_doclist[c].group_form_type_doc_id +
                            "' class='lang_please_add_file text-err'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                    }

                    for (let j = 0; j < requestDocList.length; j++) {
                        if (requestDocList.length < 2) {
                            setDataEmpUp += "<div>";
                            setDataEmpUp += `<p class='text-upload'>${c + 1
                                }${"."}<span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()
                                }' data_en='${requestDocList[j].document_name_en}'></span>`;
                            if (agency_y_doclist[c].required == "Y") {
                                setDataEmpUp += "<span class='asterisk'>*</span>";
                            }
                            setDataEmpUp +=
                                "<span id='err_file_name_" +
                                requestDocList[j].document_type_id +
                                "' class='lang_please_add_file text-err'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                        } else {
                            setDataEmpUp += '<div style="margin-left:15px;">';

                            setDataEmpUp += `<p class='text-upload'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_th}'></span>`;
                        }

                        setDataEmpUp += "</p>";
                        setDataEmpUp += "</p>";
                        setDataEmpUp += "<div class='row'>";
                        setDataEmpUp += "<div class='col-lg-12'>";
                        setDataEmpUp +=
                            "<p class='text-upload-subtext lang_support_pdf'>รองรับไฟล์ .PDF .PNG เเละ .JPG ขนาดไม่เกิน 25 mb</p>";
                        setDataEmpUp += "<div class='row mb-3' style='margin-left:0px;'>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "C") {
                                setDataEmpUp +=
                                    "<div class='upload d-none' style='padding-right:14px;'>";
                            } else {
                                setDataEmpUp +=
                                    "<div class='upload' style='padding-right:14px;'>";
                            }
                        } else {
                            setDataEmpUp +=
                                "<div class='upload' style='padding-right:14px;'>";
                        }

                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList.length >= 2) {
                                setDataEmpUp +=
                                    "<input type='file' name='file_name_group' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;,&quot;" +
                                    requestDocList[j].doclist[0].document_id +
                                    "&quot;)' hidden/>";
                            } else {
                                setDataEmpUp +=
                                    "<input type='file' name='file_name' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;,&quot;" +
                                    requestDocList[j].doclist[0].document_id +
                                    "&quot;)' hidden/>";
                            }
                        } else {
                            if (requestDocList.length >= 2) {
                                setDataEmpUp +=
                                    "<input type='file' name='file_name_group' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            } else {
                                setDataEmpUp +=
                                    "<input type='file' name='file_name' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            }
                        }

                        setDataEmpUp +=
                            "<label for='file_name_" +
                            requestDocList[j].document_type_id +
                            "' style='cursor:pointer;' class='lang_lbl_selectfile'>เลือกไฟล์</label>";


                        setDataEmpUp += "</div>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                setDataEmpUp +=
                                    "<div class='block-doc justify-content-between border_file_reject' id='divshow_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "'>";
                            } else {
                                setDataEmpUp +=
                                    "<div class='block-doc justify-content-between border_file_success' id='divshow_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "'>";
                            }
                        } else {
                            setDataEmpUp +=
                                "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                requestDocList[j].document_type_id +
                                "' style='display:none;'>";
                        }
                        setDataEmpUp += "<div class='d-flex'>";
                        setDataEmpUp += "<div>";
                        setDataEmpUp += "<img src='/assets/images/document.svg'>";
                        setDataEmpUp += "</div>";
                        setDataEmpUp += "<div class='filename'>";
                        if (requestDocList[j].doclist.length != 0) {
                            setDataEmpUp +=
                                "<a id='taga_" +
                                requestDocList[j].document_type_id +
                                "' onclick='GetDocFilePathByDocumentID(&quot;" +
                                requestDocList[j].doclist[0].document_id +
                                "&quot;,&quot;" +
                                requestDocList[j].form_type_doc_id +
                                "&quot,&quot;" +
                                objData.demand_letter_data.employer_id +
                                "&quot)'><span class='name-file namedoc' id='show_file_name_" +
                                requestDocList[j].document_type_id +
                                "'>" +
                                requestDocList[j].doclist[0].doc_title +
                                "</span></a>";

                            setDataEmpUp +=
                                '<input type="hidden" id="DocIDValue_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].document_id +
                                '">';
                        }
                        setDataEmpUp +=
                            "<p class='name-file namedoc' id='show_file_name_" +
                            requestDocList[j].document_type_id +
                            "'></p>";
                        setDataEmpUp += "</div>";

                        setDataEmpUp += "<div>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "C") {
                                setDataEmpUp +=
                                    "<img class='close d-none' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)'>";
                            } else {
                                setDataEmpUp +=
                                    "<img class='close' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;,&quot;" +
                                    requestDocList[j].doclist[0].document_id +
                                    "&quot;)'>";
                            }
                        } else {
                            setDataEmpUp +=
                                "<img class='close' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                requestDocList[j].document_type_id +
                                "&quot;)'>";
                        }

                        setDataEmpUp += "</div>";
                        setDataEmpUp += "</div>";
                        setDataEmpUp += "</div>";
                        setDataEmpUp += "</div>";
                        setDataEmpUp += "</div>";
                        setDataEmpUp += "</div>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                setDataEmpUp +=
                                    "<div class='resson' id='resson_" +
                                    requestDocList[j].document_type_id +
                                    "'><p class='textresson lang_text_reason'>หมายเหตุ</p><p> " +
                                    requestDocList[j].doclist[0].doc_comment +
                                    " </p></div>";
                                setDataEmpUp +=
                                    '<input type="hidden" id="hidden_check_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="' +
                                    requestDocList[j].doclist[0].doc_status_id +
                                    '">';
                            }
                        }

                        setDataEmpUp += "</div>";
                        setDataEmpUp +=
                            '<input type="hidden" id="requiredvalue_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            agency_y_doclist[c].required +
                            '">';
                        setDataEmpUp +=
                            '<input type="hidden" id="hidden_form_id_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            objData.demand_letter_data.demand_letter_id +
                            '">';
                        setDataEmpUp +=
                            '<input type="hidden" id="hidden_form_type_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            requestDocList[j].form_type_doc_id +
                            '">';
                        setDataEmpUp +=
                            '<input type="hidden" id="required_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            agency_y_doclist[c].group_form_type_doc_id +
                            '">';

                        if (requestDocList[j].doclist.length != 0) {
                            setDataEmpUp +=
                                '<input type="hidden" id="have_file_current_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].doclist[0].document_id +
                                '">';
                            setDataEmpUp +=
                                '<input type="hidden" id="hidden_select_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="">';
                        } else {
                            setDataEmpUp +=
                                '<input type="hidden" id="have_file_current_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="N">';
                        }
                    }

                    setDataEmpUp += "</div>";
                    setDataEmpUp += "<div class='col-md-12 mb-2 mt-2'>";
                    setDataEmpUp += "<hr>";
                    setDataEmpUp += "</div>";
                }
                jQuery("#uploadfile_emp").html(setDataEmpUp);

                $("#headfile").attr(
                    "data_th",
                    "เอกสารนายจ้าง : " +
                    objData.demand_letter_data.employer_data.emp_type_name_th
                );
                $("#headfile").attr(
                    "data_en",
                    "Employer Documents : " +
                    objData.demand_letter_data.employer_data.emp_type_name_en
                );

                $(".head_employer").html('<span class="lang_employer_documents"></span> : <span data_th="' + objData.demand_letter_data.employer_data.emp_type_name_th + '" data_en="' + objData.demand_letter_data.employer_data.emp_type_name_en +'"></span>')

                var setDataemp = "";
                var index_array = 0;
                var emp_id = objData.demand_letter_data.employer_data.employer_id;

                for (var e = 0; e < employer_doclist.length; e++) {
                    let alreadyLogged = false;
                    for (var f = 0; f < employer_doclist[e].requestdoclist.length; f++) {
                        setDataemp += "<tr>";

                        if (!alreadyLogged && employer_doclist[e].requestdoclist[0]) {
                            setDataemp += "<td >" + (e + 1) + "</td>";
                            alreadyLogged = true;
                        } else {
                            setDataemp += "<td></td>";
                        }

                        if (employer_doclist[e].requestdoclist[f].doclist.length != 0) {
                            if (
                                employer_doclist[e].requestdoclist[f].doclist[0]
                                    .emp_is_expired == "Y"
                            ) {
                                if (
                                    employer_doclist[e].requestdoclist[f].doclist[0]
                                        .doc_status_id == "UC"
                                ) {
                                    checkempfile = false;
                                    setDataemp +=
                                        "<td class='text-left' style='color:#11111F;'><span data_th='" +
                                        employer_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                        "' data_en='" +
                                        employer_doclist[e].requestdoclist[f].document_name_en +
                                        "'></span><br><p id='form_activity_detail' class='badge badge-pill statusNCOC'> <span class='lang_over'>เกิน</span> " +
                                        employer_doclist[e].requestdoclist[f].doclist[0]
                                            .emp_duration_day +
                                        " <span class='lang_lbl_day'>วัน</span> </p>" +
                                        "<br><p id='form_activity_detail' class='badge badge-pill statusNCOC lang_incorrect_doc'> เอกสารไม่ถูกต้อง</p>" +
                                        "<div class='resson'><p class='textresson lang_text_reason'>หมายเหตุ</p><p> " +
                                        employer_doclist[e].requestdoclist[f].doclist[0]
                                            .doc_comment +
                                        " </p><p class='textwarningemp lang_contact_employer_update_doc'>** กรุณาติดต่อนายจ้างให้ดำเนินการแก้ไขเอกสาร</p></div>" +
                                        "</td>";
                                } else {
                                    setDataemp +=
                                        "<td  class='text-left' style='color:#11111F;'><span data_th='" +
                                        employer_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                        "' data_en='" +
                                        employer_doclist[e].requestdoclist[f].document_name_en +
                                        "'></span> </span>" +
                                        "<p id='form_activity_detail' class='badge badge-pill statusNCOC'><span class='lang_over'>เกิน</span> " +
                                        employer_doclist[e].requestdoclist[f].doclist[0]
                                            .emp_duration_day +
                                        " <span class='lang_lbl_day'>วัน</span> </p>" +
                                        "</td>";
                                }
                            } else {
                                if (
                                    employer_doclist[e].requestdoclist[f].doclist[0]
                                        .doc_status_id == "UC"
                                ) {
                                    checkempfile = false;
                                    setDataemp +=
                                        "<td  class='text-left' style='color:#11111F;'><span data_th='" +
                                        employer_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                        "' data_en='" +
                                        employer_doclist[e].requestdoclist[f].document_name_en +
                                        "'></span><br><p id='form_activity_detail' class='badge badge-pill statusNCOC lang_incorrect_doc'> เอกสารไม่ถูกต้อง</p>" +
                                        "<div class='resson'><p class='textresson lang_text_reason'>หมายเหตุ</p><p> " +
                                        employer_doclist[e].requestdoclist[f].doclist[0]
                                            .doc_comment +
                                        " </p><p class='textwarningemp lang_contact_employer_update_doc'>** กรุณาติดต่อนายจ้างให้ดำเนินการแก้ไขเอกสาร</p></div>" +
                                        "</td>";
                                } else {
                                    setDataemp +=
                                        "<td class='text-left' style='color:#11111F;'><span data_th='" +
                                        employer_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                        "' data_en='" +
                                        employer_doclist[e].requestdoclist[f].document_name_en +
                                        "'></span></td>";
                                }
                            }
                            setDataemp +=
                                "<td  class='link-table'><a  href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook'  onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                emp_id +
                                " &quot;, &quot; " +
                                employer_doclist[e].requestdoclist[f].doclist[0].document_id +
                                " &quot;, &quot; " +
                                employer_doclist[e].requestdoclist[f].document_type_id +
                                " &quot;)'></a></td>";
                        } else {
                            setDataemp +=
                                "<td  class='text-left' style='color:#11111F;'><span data_th='" +
                                employer_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                "' data_en='" +
                                employer_doclist[e].requestdoclist[f].document_name_en +
                                "'></span></td>";
                            setDataemp +=
                                "<td  class='link-table'><a id='file_link_file_name_" +
                                employer_doclist[e].requestdoclist[f].document_type_id +
                                "'></a > - </td>";
                        }

                        setDataemp += "</tr>";

                        jQuery("#uploadfile_employer_doclist").html(setDataemp);
                    }
                }

                var setproxy = "";
                var proxy_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.proxy_doclist;
                for (let c = 0; c < proxy_doclist.length; c++) {
                    $("#uploadfile_emp").removeClass("d-none");
                    $(".div_uploadfile_emp").removeClass("d-none");

                    const requestDocList = proxy_doclist[c].requestdoclist;
                    setproxy += "<div class='col-md-12'>";
                    if (requestDocList.length >= 2) {
                        setproxy += `<p class='text-upload'>${c + 1
                            }${"."}<span class='lang_select_at_least_one_file'></span>`;
                        if (proxy_doclist[c].required == "Y") {
                            setproxy += "<span class='asterisk'>*</span>";
                        }
                        setproxy +=
                            "<span id='err_file_name_" +
                            proxy_doclist[c].group_form_type_doc_id +
                            "' class='lang_please_add_file text-err'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                    }

                    for (let j = 0; j < requestDocList.length; j++) {
                        if (requestDocList.length < 2) {
                            setproxy += "<div>";
                            setproxy += `<p class='text-upload'>${c + 1
                                }${"."}<span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()
                                }' data_en='${requestDocList[j].document_name_en}'></span>`;
                            if (proxy_doclist[c].required == "Y") {
                                setproxy += "<span class='asterisk'>*</span>";
                            }
                            setproxy +=
                                "<span id='err_file_name_" +
                                requestDocList[j].document_type_id +
                                "' class='lang_please_add_file text-err'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                        } else {
                            setproxy += '<div style="margin-left:15px;">';

                            setproxy += `<p class='text-upload'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_th}'></span>`;
                        }

                        setproxy += "</p>";
                        setproxy += "</p>";
                        setproxy += "<div class='row'>";
                        setproxy += "<div class='col-lg-12'>";
                        setproxy +=
                            "<p class='text-upload-subtext lang_support_pdf'>รองรับไฟล์ .PDF .PNG เเละ .JPG ขนาดไม่เกิน 25 mb</p>";
                        setproxy += "<div class='row mb-3' style='margin-left:0px;'>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "C") {
                                setproxy +=
                                    "<div class='upload d-none' style='padding-right:14px;'>";
                            } else {
                                setproxy += "<div class='upload' style='padding-right:14px;'>";
                            }
                        } else {
                            setproxy += "<div class='upload' style='padding-right:14px;'>";
                        }

                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList.length >= 2) {
                                setproxy +=
                                    "<input type='file' name='file_name_group' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;,&quot;" +
                                    requestDocList[j].doclist[0].document_id +
                                    "&quot;)' hidden/>";
                                setproxy +=
                                    "<input   id='file_name_check_" +
                                    requestDocList[j].document_type_id +
                                    "' onclick='checknewfile(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            } else {
                                setproxy +=
                                    "<input type='file' name='file_name' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;,&quot;" +
                                    requestDocList[j].doclist[0].document_id +
                                    "&quot;)' hidden/>";
                                setproxy +=
                                    "<input  id='file_name_check_" +
                                    requestDocList[j].document_type_id +
                                    "' onclick='checknewfile(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            }
                        } else {
                            if (requestDocList.length >= 2) {
                                setproxy +=
                                    "<input type='file' name='file_name_group' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                                setproxy +=
                                    "<input  id='file_name_check_" +
                                    requestDocList[j].document_type_id +
                                    "' onclick='checknewfile(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            } else {
                                setproxy +=
                                    "<input type='file' name='file_name' id='file_name_" +
                                    requestDocList[j].document_type_id +
                                    "' onchange='SelectFileEdit(event,&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                                setproxy +=
                                    "<input  id='file_name_check_" +
                                    requestDocList[j].document_type_id +
                                    "' onclick='checknewfile(&quot;" +
                                    requestDocList[j].document_type_id +
                                    "&quot;)' hidden/>";
                            }
                        }

                        setproxy +=
                            "<label for='file_name_check_" +
                            requestDocList[j].document_type_id +
                            "' style='cursor:pointer;' class='lang_lbl_selectfile'>เลือกไฟล์</label>";


                        setproxy += "</div>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                setproxy +=
                                    "<div class='block-doc justify-content-between border_file_reject' id='divshow_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "'>";
                            } else {
                                setproxy +=
                                    "<div class='block-doc justify-content-between border_file_success' id='divshow_file_name_" +
                                    requestDocList[j].document_type_id +
                                    "'>";
                            }
                        } else {
                            setproxy +=
                                "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                requestDocList[j].document_type_id +
                                "' style='display:none;'>";
                        }
                        setproxy += "<div class='d-flex'>";
                        setproxy += "<div>";
                        setproxy += "<img src='/assets/images/document.svg'>";
                        setproxy += "</div>";
                        setproxy += "<div class='filename'>";
                        if (requestDocList[j].doclist.length != 0) {
                            setproxy +=
                                "<a id='taga_" +
                                requestDocList[j].document_type_id +
                                "' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                objData.demand_letter_data.agency_id +
                                " &quot;, &quot; " +
                                requestDocList[j].doclist[0].document_id +
                                " &quot;, &quot; " +
                                requestDocList[j].document_type_id +
                                " &quot;)'><span class='name-file namedoc' id='show_file_name_" +
                                requestDocList[j].document_type_id +
                                "'>" +
                                requestDocList[j].doclist[0].doc_title +
                                "</span></a>";

                            setproxy +=
                                '<input type="hidden" id="DocIDValue_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].document_id +
                                '">';
                        }
                        setproxy +=
                            "<p class='name-file namedoc' id='show_file_name_" +
                            requestDocList[j].document_type_id +
                            "'></p>";
                        setproxy += "</div>";

                        setproxy += "<div>";
                        if (requestDocList[j].doclist.length != 0) {


                        } else {
                            setproxy +=
                                "<img class='close' src='/assets/images/Mark.svg' onclick='closefileEdit63(&quot;" +
                                requestDocList[j].document_type_id +
                                "&quot;)'>";
                        }

                        setproxy += "</div>";
                        setproxy += "</div>";
                        setproxy += "</div>";
                        setproxy += "</div>";
                        setproxy += "</div>";
                        setproxy += "</div>";
                        if (requestDocList[j].doclist.length != 0) {
                            if (requestDocList[j].doclist[0].doc_status_id == "UC") {
                                setproxy +=
                                    "<div class='resson' id='resson_" +
                                    requestDocList[j].document_type_id +
                                    "'><p class='textresson lang_text_reason'>หมายเหตุ</p><p> " +
                                    requestDocList[j].doclist[0].doc_comment +
                                    " </p></div>";
                                setproxy +=
                                    '<input type="hidden" id="hidden_check_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="' +
                                    requestDocList[j].doclist[0].doc_status_id +
                                    '">';

                                if (requestDocList[j].doclist[0].doc_new_lastest == "Y") {
                                    setproxy +=
                                        "<br><div class='show_new_emp_created_timestamp_" +
                                        requestDocList[j].document_type_id +
                                        " d-none'><span style='color:#535F6B;'>อัปเดตล่าสุด : <span data_th='" +
                                        convert_to_thai_datetime(
                                            requestDocList[j].doclist[0].new_emp_created_timestamp
                                        ) +
                                        "' data_en='" +
                                        convert_to_eng_datetime(
                                            requestDocList[j].doclist[0].new_emp_created_timestamp
                                        ) +
                                        "'></span></span> <span id='form_activity_detail' class='badge badge-pill statusAPSS2'>เอกสารอัปเดตใหม่</span></div><br>";
                                }
                            }
                        }

                        setproxy += "</div>";
                        setproxy +=
                            '<input type="hidden" id="requiredvalue_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            proxy_doclist[c].required +
                            '">';
                        setproxy +=
                            '<input type="hidden" id="hidden_form_id_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            objData.demand_letter_data.demand_letter_id +
                            '">';
                        setproxy +=
                            '<input type="hidden" id="hidden_form_type_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            requestDocList[j].form_type_doc_id +
                            '">';
                        setproxy +=
                            '<input type="hidden" id="required_file_name_' +
                            requestDocList[j].document_type_id +
                            '" value="' +
                            proxy_doclist[c].group_form_type_doc_id +
                            '">';

                        if (requestDocList[j].doclist.length != 0) {
                            setproxy +=
                                '<input type="hidden" id="have_file_current_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].doclist[0].document_id +
                                '">';
                            setproxy +=
                                '<input type="hidden" id="hidden_select_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="">';
                            setproxy +=
                                '<input type="hidden" id="doc_new_lastest_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="' +
                                requestDocList[j].doclist[0].doc_new_lastest +
                                '">';

                            if (requestDocList[j].doclist[0].doc_new_lastest == "Y") {
                                setproxy +=
                                    '<input type="hidden" id="new_emp_created_file_name_' +
                                    requestDocList[j].document_type_id +
                                    '" value="' +
                                    requestDocList[j].doclist[0].new_emp_created_timestamp +
                                    '">';

                                objString.push(requestDocList[j]);
                            }
                        } else {
                            setproxy +=
                                '<input type="hidden" id="have_file_current_file_name_' +
                                requestDocList[j].document_type_id +
                                '" value="N">';
                        }
                    }

                    setproxy += "</div>";
                    setproxy += "<div class='col-md-12 mb-2 mt-2'>";
                    setproxy += "<hr>";
                    setproxy += "</div>";
                }
                jQuery("#uploadfile_proxy_doclist").html(setproxy);

                var setData_OR = "";
                var objother_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.other_doclist;
                var demand_letter_id = objData.demand_letter_data.demand_letter_id;

                if (objother_doclist[0].requestdoclist[0].doclist.length == 0) {
                    $("#AddFileOrther").modal("hide");
                    $("#hidden_form_type").val(
                        objother_doclist[0].requestdoclist[0].form_type_doc_id
                    );
                    setData_OR += '<div class="row col-12">';
                    setData_OR +=
                        '<h3 class="head-step lang_other_doc" style="width:48%;">เอกสารอื่นๆ เพิ่มเติม</h3>';
                    setData_OR +=
                        '<div class="buttonaddfile" style="width:50%;"><button type="button" class="btn btn-add-file action-button float-right" onclick="OncOpenModalAddFile()"><img src="/assets/images/plus.svg"><span class="lang_add_document"></span></button></div>';
                    setData_OR += "</div>";

                    setData_OR +=
                        '<span class="head"><p class="text-upload-subtext lang_subtext_upload">เอกสารที่ใช้ประกอบการพิจารณาทั้งลูกจ้างและนายจ้าง สูงสุด 5 ไฟล์</p></span> ';
                    $(".div_show_empty").removeClass("d-none");
                    $("#uploadfile_orther").removeClass("d-none");
                } else {
                    $("#uploadfile_orther").removeClass("d-none");
                    for (var o = 0; o < objother_doclist.length; o++) {
                        const requestDocList_doclist =
                            objother_doclist[o].requestdoclist[0].doclist;
                        const requestDocList = objother_doclist[o].requestdoclist[0];
                        $("#hidden_form_type").val(
                            objother_doclist[0].requestdoclist[0].form_type_doc_id
                        );
                        setData_OR +=
                            "<div class='row col-md-12' style='margin-left:0px;'>";
                        setData_OR +=
                            '<h3 class="head-step lang_other_doc" style="width:48%;"></h3>';
                        setData_OR +=
                            '<input type="file" id="hidden_file_0" style="display:none;">';
                        setData_OR +=
                            '<input type="file" id="hidden_file_1" style="display:none;">';
                        setData_OR +=
                            '<input type="file" id="hidden_file_2" style="display:none;">';
                        setData_OR +=
                            '<input type="file" id="hidden_file_3" style="display:none;">';
                        setData_OR +=
                            '<input type="file" id="hidden_file_4" style="display:none;">';
                        setData_OR +=
                            '<input type="file" id="hidden_file_5" style="display:none;">';

                        setData_OR +=
                            '<input type="text" id="hidden_form_type_0" value="' +
                            requestDocList.form_type_doc_id +
                            '" style="display:none;">';
                        setData_OR +=
                            '<input type="text" id="hidden_form_type_1" value="' +
                            requestDocList.form_type_doc_id +
                            '" style="display:none;">';
                        setData_OR +=
                            '<input type="text" id="hidden_form_type_2" value="' +
                            requestDocList.form_type_doc_id +
                            '" style="display:none;">';
                        setData_OR +=
                            '<input type="text" id="hidden_form_type_3" value="' +
                            requestDocList.form_type_doc_id +
                            '" style="display:none;">';
                        setData_OR +=
                            '<input type="text" id="hidden_form_type_4" value="' +
                            requestDocList.form_type_doc_id +
                            '" style="display:none;">';
                        setData_OR +=
                            '<input type="text" id="hidden_form_type_5" value="' +
                            requestDocList.form_type_doc_id +
                            '" style="display:none;">';
                        setData_OR +=
                            '<input type="text" id="hidden_form_id" value="' +
                            demand_letter_id +
                            '" style="display:none;">';
                        setData_OR +=
                            '<div class="buttonaddfile" style="width:50%;"><button type="button" class="btn btn-add-file action-button float-right" onclick="OncOpenModalAddFile()"><img src="/assets/images/plus.svg"><span class="lang_add_document"></span></button></div>';
                        setData_OR +=
                            '<span class="head"><p class="text-upload-subtext lang_subtext_upload">เอกสารที่ใช้ประกอบการพิจารณาทั้งลูกจ้างและนายจ้าง สูงสุด 5 ไฟล์</p></span></h3>';
                        setData_OR +=
                            '<div class="col-md-12" id="ShowtxtBox_File" style="padding-left:0px;">';
                        for (let p = 0; p < requestDocList_doclist.length; p++) {
                            setData_OR +=
                                '<div id="getElementDoc' +
                                p +
                                '" class="checklength gutters getElementDoc">';
                            setData_OR += '<div class="col-md-12" style="padding-left:0px;">';
                            setData_OR += `<p class='text-upload'>${requestDocList_doclist[p].doc_title}`;
                            if (requestDocList_doclist.length != 0) {
                                if (requestDocList_doclist[p].doc_status_id == "UC") {
                                    setData_OR +=
                                        "<span id='err_" +
                                        p +
                                        "' class='lang_please_add_file text-err'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                                }
                            }
                            setData_OR += "</p>";

                            setData_OR += "<div class='row'>";
                            setData_OR += "<div class='col-lg-12'>";
                            setData_OR +=
                                "<p class='text-upload-subtext lang_support_pdf'>รองรับไฟล์ .PDF .PNG .ZIP เเละ .JPG ขนาดไม่เกิน 25 mb</p>";
                            setData_OR += "<div class='row' style='margin-left:0px;'>";
                            if (requestDocList_doclist[p].doc_status_id != "C") {
                                setData_OR +=
                                    "<div class='upload' id='upload_" +
                                    p +
                                    "' style='padding-right:14px;'>";
                            } else {
                                setData_OR += "<div class='upload' style='display:none;'>";
                            }

                            setData_OR +=
                                "<input type='file' name='file_name_orther' id='file_other_" +
                                p +
                                "' onchange='SelectFileOrtherEdit(event,&quot;" +
                                p +
                                "&quot;,&quot;" +
                                requestDocList_doclist[p].document_id +
                                "&quot;)' hidden/>";

                            setData_OR +=
                                "<label class='lang_lbl_selectfile' for='file_other_" +
                                p +
                                "' style='cursor:pointer;'>เลือกไฟล์</label>";

                            if (requestDocList_doclist.length != 0) {
                                setData_OR +=
                                    "<span class='lang_nofile_chosen' id='divselect_file_name_" +
                                    p +
                                    "' style='display:none;'>No file chosen</span>";
                            } else {
                                setData_OR +=
                                    "<span class='lang_nofile_chosen' id='divselect_file_name_" +
                                    p +
                                    "'>No file chosen</span>";
                            }
                            setData_OR += "</div>";
                            if (requestDocList_doclist.length != 0) {
                                if (requestDocList_doclist[p].doc_status_id == "UC") {
                                    setData_OR +=
                                        "<div class='block-doc justify-content-between border_file_reject' id='divshow_file_name_" +
                                        p +
                                        "'>";
                                } else {
                                    setData_OR +=
                                        "<div class='block-doc justify-content-between border_file_success' id='divshow_file_name_" +
                                        p +
                                        "'>";
                                }
                            } else {
                                setData_OR +=
                                    "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                    p +
                                    "' style='display:none;'>";
                            }
                            setData_OR += "<div class='d-flex'>";
                            setData_OR +=
                                "<div><img src='/assets/images/document.svg'></div>";

                            setData_OR += "<div class='filename'>";
                            if (requestDocList_doclist.length != 0) {
                                setData_OR +=
                                    "<a id='tagaorther_" +
                                    p +
                                    "' onclick='GetDocFilePathByDocumentID(&quot;" +
                                    requestDocList_doclist[p].document_id +
                                    "&quot;,&quot;" +
                                    requestDocList.form_type_doc_id +
                                    "&quot,&quot;" +
                                    objData.demand_letter_data.demand_letter_id +
                                    "&quot)'><span class='name-file namedoc' id='show_file_title_" +
                                    p +
                                    "' value='" +
                                    requestDocList_doclist[p].doc_title +
                                    "'>" +
                                    requestDocList_doclist[p].doc_title +
                                    "</span></a>";

                                setData_OR +=
                                    '<input type="hidden" id="DocIDValue_file_name_' +
                                    p +
                                    '" value="' +
                                    requestDocList_doclist[p].document_id +
                                    '">';
                            } else {
                                setData_OR +=
                                    "<p class='name-file namedoc' id='show_file_title_" +
                                    p +
                                    "'></p>";
                            }

                            setData_OR += "</div>";

                            if (requestDocList_doclist[p].doc_status_id != "C") {
                                setData_OR +=
                                    "<div><img class='close' src='/assets/images/Mark.svg' onclick='deleteboxfile63(&quot;" +
                                    p +
                                    "&quot;,&quot;" +
                                    requestDocList_doclist[p].document_id +
                                    "&quot;)'></div>";
                            }
                            setData_OR += "</div>";
                            setData_OR += "</div>";
                            setData_OR += "</div>";
                            setData_OR += "</div>";
                            setData_OR += "</div>";
                            if (requestDocList_doclist.length != 0) {
                                if (requestDocList_doclist[p].doc_status_id == "UC") {
                                    setData_OR +=
                                        "<div class='resson mt-3' id='resson_" +
                                        p +
                                        "'><p class='textresson'>หมายเหตุ</p><p> " +
                                        requestDocList_doclist[p].doc_comment +
                                        " </p></div>";
                                    setData_OR +=
                                        '<input type="hidden" id="hidden_check_' +
                                        p +
                                        '" value="' +
                                        requestDocList_doclist[p].doc_status_id +
                                        '">';
                                }
                            }
                            setData_OR +=
                                '<input type="hidden" id="hidden_current_file_' +
                                p +
                                '" value="' +
                                requestDocList_doclist[p].document_id +
                                '">';
                            setData_OR +=
                                '<input type="hidden" id="hidden_current_file_new' +
                                p +
                                '" value="Y">';
                            setData_OR +=
                                '<input type="hidden" id="hidden_current_file_name_' +
                                p +
                                '" value="' +
                                requestDocList_doclist[p].doc_title +
                                '">';
                            setData_OR +=
                                '<input type="hidden" id="doc_other_no_file_name_' +
                                p +
                                '" value="' +
                                requestDocList_doclist[p].doc_other_no +
                                '">';
                            setData_OR +=
                                '<input type="hidden" id="hidden_form_type_up" value="' +
                                requestDocList.form_type_doc_id +
                                '" />';
                            setData_OR +=
                                "<div class='col-md-12 mb-2 mt-2' style='padding-left:0px;' id='line_" +
                                p +
                                "'><hr></div>";
                            setData_OR += "</div>";
                            setData_OR += "</div>";
                        }
                    }
                }

                jQuery("#uploadfile_orther").html(setData_OR);


                if (typeof captureOtherDocBaseline === "function") captureOtherDocBaseline();
                if (typeof refreshOtherDocSubtext === "function") refreshOtherDocSubtext(statusGroup);
                if (typeof markOtherDocRequired === "function") markOtherDocRequired(statusGroup);
                if (typeof updateAddFileButtonVisibility === "function") updateAddFileButtonVisibility(statusGroup);
            } else {
                $("#loading").delay().fadeOut("");
                alert(data.msg);
            }

            changeDatalang();
        },
        "json"
    );
}
function onloadfileDetail() {
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
    var request = jQuery.post(
        url_GetDataRequestFormEtrackingMOU41_4_59_ByGroupID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                $("#loading").delay().fadeOut("");
                var app = getParameterByName("App");
                var objData = data.data[0];
                var employer_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.employer_n_doclist;
                var saleperson_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.saleperson_doclist;
                var agency_n_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.agency_n_doclist;
                var agency_y_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.agency_y_doclist;
                var proxydoclist =
                    objData.demand_letter_data.demand_letter_doc_list.proxy_doclist;
                var setData = "";

                var index_array = 0;

                $("#headfileEmp").text("นายจ้าง" + $(".text_employer_type").text());

                var setData = "";
                var emp_ID = objData.demand_letter_data.employer_data.employer_id;


                $("#head_employer").html('<span class="lang_menu_helpservice"></span> : <span data_th="' + objData.demand_letter_data.employer_data.emp_type_name_th + '" data_en="' + objData.demand_letter_data.employer_data.emp_type_name_en + '"></span>')

                for (let c = 0; c < employer_doclist.length; c++) {
                    const requestDocList = employer_doclist[c].requestdoclist;
                    setData += "<tr>";
                    setData += `<td>${c + 1}</td>`;
                    for (let k = 0; k < requestDocList.length; k++) {
                        if (k != index_array) {
                            setData += `<td></td>`;
                        }
                        setData += `<td class='text-left'><span data_th='${requestDocList[k].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[k].document_name_en}'></span></td>`;
                        if (requestDocList[k].doclist.length > 0) {
                            setData +=
                                "<td class='link-table'><a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                emp_ID +
                                " &quot;, &quot; " +
                                requestDocList[k].doclist[0].document_id +
                                " &quot;, &quot; " +
                                requestDocList[k].document_type_id +
                                " &quot;)'></a></td>";

                        } else {
                            setData += `<td class='link-table' tyle='text-decoration: none;'>-</td>`;
                        }

                        setData += "</tr>";
                    }
                }

                jQuery("#listuploadfile_employer46_59").html(setData);

                var setproxy = "";
                if (proxydoclist.length != 0) {
                    for (let c = 0; c < proxydoclist.length; c++) {
                        $(".div_head_show_empUP").css("display", "block");
                        $(".div_head_show_file_empUP").css("display", "block");
                        const requestDocList = proxydoclist[c].requestdoclist;
                        setproxy += "<tr>";
                        setproxy += `<td data-label='ลำดับ'>${c + 1}</td>`;

                        for (let j = 0; j < requestDocList.length; j++) {
                            if (j != index_array) {
                                setproxy += `<td data-label='ลำดับ'></td>`;
                            }
                            setproxy += `<td data-label='รายการ'  class='text-left' data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en}'></td>`;
                            setproxy += `<td data-label='เอกสาร'  class='link-table'><a id='file_link_file_name_${requestDocList[j].document_type_id}'></a></td>`;
                            setproxy += "</tr>";
                        }
                    }
                    jQuery("#listuploadfile_proxy").html(setproxy);
                }

                var setDataAgency = "";

                for (let c = 0; c < agency_n_doclist.length; c++) {
                    $(".div_head_show_empUP_edit").css("display", "block");
                    $(".div_head_show_file_empUP_edit").css("display", "block");
                    const requestDocList = agency_n_doclist[c].requestdoclist;
                    setDataAgency += "<tr>";
                    setDataAgency += `<td>${c + 1}</td>`;

                    for (let j = 0; j < requestDocList.length; j++) {
                        if (j != index_array) {
                            setDataAgency += `<td></td>`;
                        }
                        setDataAgency += `<td class='text-left'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en}'></span></td>`;
                        setDataAgency += `<td class='link-table'><span id='file_link_file_name_${requestDocList[j].document_type_id}'></span></td>`;
                        setDataAgency += "</tr>";
                    }
                }
                jQuery("#list_agency").html(setDataAgency);

                var setDataempUp = "";
                for (let c = 0; c < agency_y_doclist.length; c++) {
                    $(".div_head_show_empUP_edit").css("display", "block");
                    $(".div_head_show_file_empUP_edit").css("display", "block");
                    const requestDocList = agency_y_doclist[c].requestdoclist;
                    setDataempUp += "<tr>";
                    setDataempUp += `<td>${c + 1}</td>`;

                    for (let j = 0; j < requestDocList.length; j++) {
                        if (j != index_array) {
                            setDataempUp += `<td></td>`;
                        }
                        setDataempUp += `<td class='text-left'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en}'></span></td>`;
                        setDataempUp += `<td class='link-table'><span id='file_link_file_name_${requestDocList[j].document_type_id}'></span></td>`;
                        setDataempUp += "</tr>";
                    }
                }
                jQuery("#empuploadfile_edit").html(setDataempUp);

                changeDatalang();
            } else {
                alert(data.msg);
                $("#loading").delay().fadeOut("");
            }
        },
        "json"
    );
}
function OnloadFileSellEdit() {
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
    var request = jQuery.post(
        url_GetDataRequestFormEtrackingMOU41_4_59_ByGroupID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                $("#loading").delay().fadeOut("");
                var objData = data.data[0];
                var employer_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.employer_n_doclist;
                var saleperson_doclist =
                    objData.demand_letter_data.demand_letter_doc_list.saleperson_doclist;
                var employer_id = $("#employer_id").val();
                var objdoclist =
                    objData.demand_letter_data.demand_letter_doc_list.saleperson_doclist;
                var setDataemp = "";
                var index_array = 0;
                var issale_data = $("#issale").val();

                if (issale_data == "Y") {
                    $(".div_head_sell").css("display", "block");
                    $(".listuploadfile_sell").removeClass("d-none");


                    for (var e = 0; e < saleperson_doclist.length; e++) {
                        let alreadyLogged = false;
                        for (
                            var f = 0;
                            f < saleperson_doclist[e].requestdoclist.length;
                            f++
                        ) {
                            setDataemp += "<tr>";

                            if (!alreadyLogged && saleperson_doclist[e].requestdoclist[0]) {
                                setDataemp += "<td >" + (e + 1) + "</td>";
                                alreadyLogged = true;
                            } else {
                                setDataemp += "<td></td>";
                            }

                            if (saleperson_doclist[e].requestdoclist[f].doclist.length != 0) {
                                if (
                                    saleperson_doclist[e].requestdoclist[f].doclist[0]
                                        .emp_is_expired == "Y"
                                ) {
                                    if (
                                        saleperson_doclist[e].requestdoclist[f].doclist[0]
                                            .doc_status_id == "UC"
                                    ) {
                                        checkempfile = false;
                                        setDataemp +=
                                            "<td class='text-left' style='color:#11111F;'>" +
                                            saleperson_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                            "<br><p id='form_activity_detail' class='badge badge-pill statusNCOC'><span class='lang_over'>เกิน</span> " +
                                            saleperson_doclist[e].requestdoclist[f].doclist[0]
                                                .emp_duration_day +
                                            " <span class='lang_lbl_day'>วัน</span> </p>" +
                                            "<br><p id='form_activity_detail' class='badge badge-pill statusNCOC lang_incorrect_doc'> เอกสารไม่ถูกต้อง</p>" +
                                            "<div class='resson'><p class='textresson'>หมายเหตุ</p><p> " +
                                            saleperson_doclist[e].requestdoclist[f].doclist[0]
                                                .doc_comment +
                                            " </p><p class='textwarningemp'>** กรุณาติดต่อนายจ้างให้ดำเนินการแก้ไขเอกสาร</p></div>" +
                                            "</td>";
                                    } else {
                                        setDataemp +=
                                            "<td  class='text-left' style='color:#11111F;'>" +
                                            saleperson_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                            " </span>" +
                                            "<p id='form_activity_detail' class='badge badge-pill statusNCOC'><span class='lang_over'>เกิน</span> " +
                                            saleperson_doclist[e].requestdoclist[f].doclist[0]
                                                .emp_duration_day +
                                            " <span class='lang_lbl_day'>วัน</span> </p>" +
                                            "</td>";
                                    }
                                } else {
                                    if (
                                        saleperson_doclist[e].requestdoclist[f].doclist[0]
                                            .doc_status_id == "UC"
                                    ) {
                                        checkempfile = false;
                                        setDataemp +=
                                            "<td  class='text-left' style='color:#11111F;'>" +
                                            saleperson_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                            "<br><p id='form_activity_detail' class='badge badge-pill statusNCOC lang_incorrect_doc'> เอกสารไม่ถูกต้อง</p>" +
                                            "<div class='resson'><p class='textresson'>หมายเหตุ</p><p> " +
                                            saleperson_doclist[e].requestdoclist[f].doclist[0]
                                                .doc_comment +
                                            " </p><p class='textwarningemp'>** กรุณาติดต่อนายจ้างให้ดำเนินการแก้ไขเอกสาร</p></div>" +
                                            "</td>";
                                    } else {
                                        setDataemp +=
                                            "<td class='text-left' style='color:#11111F;'>" +
                                            saleperson_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                            "</td>";
                                    }
                                }
                                setDataemp +=
                                    "<td  class='link-table'><a  href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook'  onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                    employer_id +
                                    " &quot;, &quot; " +
                                    saleperson_doclist[e].requestdoclist[f].doclist[0]
                                        .document_id +
                                    " &quot;, &quot; " +
                                    saleperson_doclist[e].requestdoclist[f].document_type_id +
                                    " &quot;)'></a></td>";
                            } else {
                                setDataemp +=
                                    "<td  class='text-left' style='color:#11111F;'>" +
                                    saleperson_doclist[e].requestdoclist[f].document_name_th?.toSafeJsonHtml() +
                                    "</td>";
                                setDataemp +=
                                    "<td  class='link-table'><a id='file_link_file_name_" +
                                    saleperson_doclist[e].requestdoclist[f].document_type_id +
                                    "'></a > - </td>";
                            }

                            setDataemp += "</tr>";

                            jQuery("#uploadfile_employer_doclist").html(setDataemp);
                        }
                    }
                    jQuery("#listuploadfile_sell").html(setDataemp);


                } else {
                    $(".div_head_sell").css("display", "none");
                    $(".listuploadfile_sell").addClass("d-none");
                }

                var setDatasell = "";
                var checksale_edit = true;
                var bus_type_id = $("#bus_type_id :selected").attr("data-sale");
                var checksale_edit = $("input[name='permit_type']");
                if (
                    checksale_edit.is(":checked") &&
                    checksale_edit.data("sale") === "Y"
                ) {
                    checksale_edit = true;
                } else {
                    checksale_edit = false;
                }

                if (checksale_edit && bus_type_id == "Y") {
                    $(".div_head_show_sell_edit").css("display", "block");
                    $(".div_head_show_file_sell_edit").css("display", "block");

                    for (let v = 0; v < objdoclist.length; v++) {
                        const requestDocList = objdoclist[v].requestdoclist;
                        setDatasell += "<tr>";
                        setDatasell += `<td data-label='ลำดับ'>${v + 1}</td>`;

                        for (let n = 0; n < requestDocList.length; n++) {
                            if (n != index_array) {
                                setDatasell += `<td data-label='ลำดับ'></td>`;
                            }
                            setDatasell += `<td data-label='รายการ'  class='text-left' data_th='${requestDocList[n].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[n].document_name_en}'></td>`;
                            setDatasell += `<td data-label='เอกสาร'  class='link-table'><span id='file_link_file_name_${requestDocList[n].document_type_id}'></span></td>`;
                            if (requestDocList[n].doclist.length > 0) {
                                setDatasell +=
                                    "<td class='link-table'><a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                    employer_id +
                                    " &quot;, &quot; " +
                                    requestDocList[n].doclist[0].document_id +
                                    " &quot;, &quot; " +
                                    requestDocList[n].document_type_id +
                                    " &quot;)'></a></td>";

                            } else {
                                setDatasell += `<td class='link-table' tyle='text-decoration: none;'>-</td>`;
                            }
                            setDatasell += "</tr>";
                        }
                    }
                } else {
                    $(".div_head_show_sell_edit").css("display", "none");
                    $(".div_head_show_file_sell_edit").css("display", "none");
                }
                jQuery("#listuploadfile_file_sell_edit").html(setDatasell);
            } else {
                $("#loading").delay().fadeOut("");
                alert(data.msg);
            }

            changeDatalang();
        },
        "json"
    );
}
function OnchangeSumTypeJob(data) {
    var inputID = data.id;
    var sumMen;
    var sumWomen;
    if (
        inputID == "numbermen_1" ||
        inputID == "numbermen_2" ||
        inputID == "numbermen_3" ||
        inputID == "numbermen_4"
    ) {
        var numbermen_1 = $("#numbermen_1").val();
        var numbermen_2 = $("#numbermen_2").val();
        var numbermen_3 = $("#numbermen_3").val();
        var numbermen_4 = $("#numbermen_4").val();

        if (numbermen_1 == "") {
            var changeToInt1 = parseInt("0");
        } else {
            var changeToInt1 = parseInt(numbermen_1);
        }
        if (numbermen_2 == "") {
            var changeToInt2 = parseInt("0");
        } else {
            var changeToInt2 = parseInt(numbermen_2);
        }
        if (numbermen_3 == "") {
            var changeToInt3 = parseInt("0");
        } else {
            var changeToInt3 = parseInt(numbermen_3);
        }
        if (numbermen_4 == "") {
            var changeToInt4 = parseInt("0");
        } else {
            var changeToInt4 = parseInt(numbermen_4);
        }
        sumMen = changeToInt1 + changeToInt2 + changeToInt3 + changeToInt4;
        $("#number_of_alien_male").val(sumMen);
    } else if (
        inputID == "numberwomen_1" ||
        inputID == "numberwomen_2" ||
        inputID == "numberwomen_3" ||
        inputID == "numberwomen_4"
    ) {
        var numberwomen_1 = $("#numberwomen_1").val();
        var numberwomen_2 = $("#numberwomen_2").val();
        var numberwomen_3 = $("#numberwomen_3").val();
        var numberwomen_4 = $("#numberwomen_4").val();

        if (numberwomen_1 == "") {
            var changeToInt1 = parseInt("0");
        } else {
            var changeToInt1 = parseInt(numberwomen_1);
        }
        if (numberwomen_2 == "") {
            var changeToInt2 = parseInt("0");
        } else {
            var changeToInt2 = parseInt(numberwomen_2);
        }
        if (numberwomen_3 == "") {
            var changeToInt3 = parseInt("0");
        } else {
            var changeToInt3 = parseInt(numberwomen_3);
        }
        if (numberwomen_4 == "") {
            var changeToInt4 = parseInt("0");
        } else {
            var changeToInt4 = parseInt(numberwomen_4);
        }
        sumWomen = changeToInt1 + changeToInt2 + changeToInt3 + changeToInt4;
        $("#number_of_alien_female").val(sumWomen);
    }

    var IntMen = $("#number_of_alien_male").val();
    var IntWomen = $("#number_of_alien_female").val();
    if (IntMen == "" || IntMen == undefined) {
        IntMen = parseInt("0");
    } else {
        IntMen = parseInt($("#number_of_alien_male").val());
    }
    if (IntWomen == "" || IntWomen == undefined) {
        IntWomen = parseInt("0");
    } else {
        IntWomen = parseInt($("#number_of_alien_female").val());
    }
    $("#number_of_alien_all").val(IntMen + IntWomen);
}
$("#bus_type_id").on("change", function () {
    OnloadFileMOUEdit();
    onloadfileDetail();
});
$("#checkboxjob_1").on("click", function () {
    var checksale = $("#checkboxjob_1").attr("data-sale");
    var bus_type_id = $("#bus_type_id :selected").attr("data-sale");

    if (bus_type_id == "Y" && checksale == "Y") {
        $("#issale").val(bus_type_id);
        $(".div_head_sell").css("display", "block");
        $(".listuploadfile_sell").css("display", "contents");
    } else {
        $("#issale").val("N");
        $(".div_head_sell").css("display", "none");
        $(".listuploadfile_sell").css("display", "none");
    }
    OnloadFileMOUEdit();
    onloadfileDetail();
});
$("#NextStepThreePageTwo").on("click", function () {
    current_fs = $(this).parent();
    next_fs = $(this).parent().next();

    $("#progressbar li").eq($("fieldset").index(next_fs)).addClass("active");

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
});
function showdetailedit() {
    initializePageCheckbox();

    $(".text_workhour_per_day").html(
        $("#workhour_per_day").val() + " " + "<span class='lang_lbl_hour'></span>"
    );
    $(".text_dayoff_per_week").html(
        $("#dayoff_per_week").val() + " " + "<span class='lang_lbl_day'></span>"
    );
    $(".text_holiday_per_year").html(
        $("#holiday_per_year").val() + " " + "<span class='lang_lbl_day'></span>"
    );
    $(".text_wage_rate_per_day").html(
        $("#wage_rate_per_day").val() +
        " " +
        "<span class='lang_bath_per_day'></span>"
    );
    $(".text_traditional_holiday").html(
        $("#traditional_holiday").val() + " " + "<span class='lang_lbl_day'></span>"
    );

    var year = $("#year_work :selected").val();
    var month = $("#month_work :selected").val();
    var day = $("#day_work :selected").val();
    var chkData = true;

    if (year == "0" && month !== "0" && day !== "0") {
        $(".text_year_month_day_work").html(
            month +
            " " +
            "<span class='lang_lbl_month'></span>" +
            " " +
            day +
            " " +
            "<span class='lang_lbl_day'></span>"
        );
    } else if (year !== "0" && month == "0" && day !== "0") {
        $(".text_year_month_day_work").html(
            year +
            " " +
            "<span class='lang_lbl_year'></span>" +
            " " +
            day +
            " " +
            "<span class='lang_lbl_day'></span>"
        );
    } else if (year !== "0" && month !== "0" && day == "0") {
        $(".text_year_month_day_work").html(
            year +
            " " +
            "<span class='lang_lbl_year'></span>" +
            " " +
            month +
            " " +
            "<span class='lang_lbl_month'></span>"
        );
    } else if (year !== "0" && month == "0" && day == "0") {
        $(".text_year_month_day_work").html(
            year + " " + "<span class='lang_lbl_year'></span>"
        );
    } else if (year == "0" && month !== "0" && day == "0") {
        $(".text_year_month_day_work").html(
            month + " " + "<span class='lang_lbl_month'></span>"
        );
    } else if (year == "0" && month == "0" && day !== "0") {
        $(".text_year_month_day_work").html(
            day + " " + "<span class='lang_lbl_day'></span>"
        );
    } else {
        $(".text_year_month_day_work").html(
            year +
            " " +
            "<span class='lang_lbl_year'></span>" +
            " " +
            month +
            " " +
            "<span class='lang_lbl_month'></span>" +
            " " +
            day +
            " " +
            "<span class='lang_lbl_day'></span>"
        );
    }

    $(".text_number_of_alien_all").html(
        "<span>" +
        $("#number_of_alien_all").val() +
        "</span> <span class='lang_person'></span>"
    );

    const number_of_alien_female = $("#number_of_alien_female").val();
    number_of_alien_female.length === 0
        ? $(".text_number_of_alien_female").text("-")
        : $(".text_number_of_alien_female").html(
            "<span>" +
            $("#number_of_alien_female").val() +
            "</span>" +
            " <span class='lang_person'></span>"
        );
    const number_of_alien_male = $("#number_of_alien_male").val();
    number_of_alien_male.length === 0
        ? $(".text_number_of_alien_male").text("-")
        : $(".text_number_of_alien_male").html(
            "<span>" +
            $("#number_of_alien_male").val() +
            "</span>" +
            " <span class='lang_person'></span>"
        );

    $(".text_nationality_id_mou").text($("#nationality_id_mou :selected").text());
    $(".text_nationality_id_mou").attr(
        "data_th",
        $("#nationality_id_mou :selected").attr("data_th")
    );
    $(".text_nationality_id_mou").attr(
        "data_en",
        $("#nationality_id_mou :selected").attr("data_en")
    );

    $(".text_recruitment_agency").text($("#recruitment_agency :selected").text());
    $(".text_recruitment_agency").attr(
        "data_th",
        $("#recruitment_agency :selected").attr("data_th")
    );
    $(".text_recruitment_agency").attr(
        "data_en",
        $("#recruitment_agency :selected").attr("data_en")
    );

    $(".text_address_recruitment").text($("#address_recruitment").val());

    if ($("#recruitment_agency_assign_at").val() != "") {
        $(".text_recruitment_agency_assign_at").attr(
            "data_th",
            convert_to_thai_date($("#recruitment_agency_assign_at").val())
        );
        $(".text_recruitment_agency_assign_at").attr(
            "data_en",
            convert_to_eng_date($("#recruitment_agency_assign_at").val())
        );
    }

    $(".text_name_address_work").text($("#name_address_work :selected").text());
    $(".text_name_address_work").attr(
        "data_th",
        $("#name_address_work :selected").attr("data_th")
    );
    $(".text_name_address_work").attr(
        "data_en",
        $("#name_address_work :selected").attr("data_en")
    );

    $(".text_address_work").text($("#address_work").val());

    $(".text_permit_catework_id").text($("#permit_catework_id :selected").text());
    $(".text_permit_catework_id").attr(
        "data_th",
        $("#permit_catework_id :selected").attr("data_th")
    );
    $(".text_permit_catework_id").attr(
        "data_en",
        $("#permit_catework_id :selected").attr("data_en")
    );

    $(".text_bus_type_id").text($("#bus_type_id :selected").text());
    $(".text_bus_type_id").attr(
        "data_th",
        $("#bus_type_id :selected").attr("data_th")
    );
    $(".text_bus_type_id").attr(
        "data_en",
        $("#bus_type_id :selected").attr("data_en")
    );

    $(".text_age").html(
        "<span class='lang_from'>ตั้งเเต่</span>" +
        " " +
        $("#start_age").val() +
        " " +
        "<span class='lang_lbl_to'>ถึง</span>" +
        " " +
        $("#end_age").val() +
        " " +
        "<span class='lang_lbl_year'></span>"
    );
    $(".text_height").html(
        "<span class='lang_from'>ตั้งเเต่</span>" +
        " " +
        $("#start_height").val() +
        " " +
        "<span class='lang_lbl_to'>ถึง</span>" +
        " " +
        $("#end_height").val() +
        " " +
        "<span class='lang_lbl_cm'></span>"
    );
    $(".text_weight").html(
        "<span class='lang_from'>ตั้งเเต่</span>" +
        " " +
        $("#start_weight").val() +
        " " +
        "<span class='lang_lbl_to'>ถึง</span>" +
        " " +
        $("#end_weight").val() +
        " " +
        "<span class='lang_lbl_kg'></span>"
    );

    $(".text_DataImmigrationCheckpoint").text(
        $("#DataImmigrationCheckpoint :selected").text()
    );
    $(".text_DataImmigrationCheckpoint").attr(
        "data_th",
        $("#DataImmigrationCheckpoint :selected").attr("data_th")
    );
    $(".text_DataImmigrationCheckpoint").attr(
        "data_en",
        $("#DataImmigrationCheckpoint :selected").attr("data_en")
    );

    $(".text_fsc_id").text($("#fsc_id :selected").text());
    $(".text_fsc_id").attr("data_th", $("#fsc_id :selected").attr("data_th"));
    $(".text_fsc_id").attr("data_en", $("#fsc_id :selected").attr("data_en"));

    $(".text_address_now").text($("#text_address_emp_nonw").text());


    var arrayData = [];
    var arrayWomen = [];
    var arrayMen = [];
    var arraySave_Men = [];
    var arraySave_Women = [];
    if ($("#checkboxjob_4").is(":checked")) {
        var permit_catework_Dropdown = $("#permit_catework_Dropdown :selected").text();
        arrayData.push(permit_catework_Dropdown);
        if ($("#numbermen_checkboxjob_4").val() == "") {
            var dataMen1 = "-";
            var dataMenSave1 = "0";
        } else {
            var dataMen1 = $("#numbermen_checkboxjob_4").val();
            var dataMenSave1 = $("#numbermen_checkboxjob_4").val();
        }
        if ($("#numberwomen_checkboxjob_4").val() == "") {
            var dataWomen1 = "-";
            var dataWomenSave1 = "0";
        } else {
            var dataWomen1 = $("#numberwomen_checkboxjob_4").val();
            var dataWomenSave1 = $("#numberwomen_checkboxjob_4").val();
        }
        arrayMen.push(dataMen1);
        arrayWomen.push(dataWomen1);
        arraySave_Men.push(dataMenSave1);
        arraySave_Women.push(dataWomenSave1);
        $("#All_Men").val(arraySave_Men);
        $("#All_Women").val(arraySave_Women);
    } else {
        $('input[name="permit_type"]:checked').each(function () {
            var ID_checked = this.id;
            if ($("#numbermen_" + ID_checked).val() == "") {
                var dataMen1 = "-";
                var dataMenSave1 = "0";
            } else {
                var dataMen1 = $("#numbermen_" + ID_checked).val();
                var dataMenSave1 = $("#numbermen_" + ID_checked).val();
            }
            if ($("#numberwomen_" + ID_checked).val() == "") {
                var dataWomen1 = "-";
                var dataWomenSave1 = "0";
            } else {
                var dataWomen1 = $("#numberwomen_" + ID_checked).val();
                var dataWomenSave1 = $("#numberwomen_" + ID_checked).val();
            }
            arrayMen.push(dataMen1);
            arrayWomen.push(dataWomen1);
            arraySave_Men.push(dataMenSave1);
            arraySave_Women.push(dataWomenSave1);
            var dataValue = $(this).attr("value");
            arrayData.push(dataValue);
            $("#All_Men").val(arraySave_Men);
            $("#All_Women").val(arraySave_Women);
        });
    }

    var setdata = "";
    var num = 1;
    for (var c = 0; c < arrayData.length; c++) {
        setdata += "<tr>";
        setdata += '<td data-label="ลำดับ" >' + num++ + "</td>";
        if ($("#checkboxjob_4").is(":checked")) {
            var data_th = $("#permit_catework_Dropdown :selected").attr("data_th");
            var data_en = $("#permit_catework_Dropdown :selected").attr("data_en");

            setdata +=
                '<td data-label="ลักษณะงาน" class="text-left" data_th="' +
                data_th +
                '" data_en="' +
                data_en +
                '"></td>';
        } else {
            var data_th = $("#checkboxjob_" + arrayData[c]).attr("data_th")
            var data_en = $("#checkboxjob_" + arrayData[c]).attr("data_en")
            setdata +=
                setdata += '<td data-label="ลักษณะงาน" class="text-left" data_th="' + data_th + '" data_en="' + data_en + '"></td>'
        }

        setdata +=
            '<td data-label="เพศชาย"  class="link-table text-md-right text-left"><span>' +
            arrayMen[c] +
            "</td>";
        setdata +=
            '<td data-label="เพศหญิง" class="link-table text-md-right text-left">' +
            arrayWomen[c] +
            "</td>";
        setdata += "</tr>";
    }
    $("#tabelshowtypejob").html(setdata);
    changeDatalang();
}
async function UpdateRequestDemandLetter41_4_59() {
    const isTokenValid = await validate_token_presubmit();

    if (!isTokenValid) {
        return false;
    }
    $("#loading").css("display", "block");
    var formData = new FormData();
    var dataarray = [];
    var arraytype = [];
    var tmpfiletitle = [];

    jQuery("[name='file_name']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file = fileUpload[0].files[0];
        if (file != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#hidden_form_type_" + this.id).val();
            var have_file_current_file = $("#have_file_current_" + this.id).val();


            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file);
        }
    });
    jQuery("[name='file_name_group']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file = fileUpload[0].files[0];
        if (file != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#hidden_form_type_" + this.id).val();
            var have_file_current_file_or = $("#have_file_current_" + this.id).val();


            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file);
        }
    });
    jQuery("#ShowtxtBox_File > .checklength").each(function (index) {
        var uploadid = jQuery(this)[0].id.replace("getElementDoc", "");
        var fileUp2 = $("#hidden_file_up_" + uploadid)[0].files;
        var hidden_check_file_name = $("#hidden_check_" + uploadid).val();
        var doc_other_no = $("#doc_other_no_file_name_" + uploadid).val();
        if (fileUp2.length != 0) {
            formData.append("docfile", fileUp2[0]);
            if (hidden_check_file_name == "UC") {
                var filename_orther = $("#hidden_current_file_name_" + uploadid).val();
                tmpfiletitle.push(filename_orther);
                var hidden_form_type = $("#hidden_form_type_up").val();
                var result_type = hidden_form_type + "-" + doc_other_no;
                arraytype.push(result_type);
            } else {
                var filename_orther = $("#text_file_name_" + uploadid).val();
                tmpfiletitle.push(filename_orther);
                var hidden_form_type = $("#hidden_form_type_up").val();
                arraytype.push(hidden_form_type);
            }
        }
    });

    var array_job_desc_id = [];
    var dataarray = "";
    $('input[name="permit_type"]:checked').each(function () {
        var ID_checked = this.id;
        if (ID_checked != "checkboxjob_4") {
            array_job_desc_id.push(this.value);
        }
    });
    if ($("#checkboxjob_4").is(":checked")) {
        var hidden_checkboxtypejob = $("#permit_catework_Dropdown :selected").val();
        array_job_desc_id.push(hidden_checkboxtypejob);
    }

    formData.append("document_id", currentfile.toString());
    formData.append("document_title", tmpfiletitle.toString());
    formData.append("form_type_doc_id", arraytype.toString());
    formData.append("permit_cate_id", array_job_desc_id.toString());
    if (core_datas.default_profile.user_type_id == "5") {
        formData.append("user_id", core_datas.default_profile.user_id);
    } else {
        formData.append("user_id", core_datas.current_profile.user_id);
    }
    formData.append("group_id", $("#group_id").val());
    formData.append("bus_type_id", $("#bus_type_id :selected").val());
    var nationality_id = $("#nationality_id_mou :selected").val().split(",");
    formData.append("nationality_id", nationality_id[0]);
    formData.append(
        "number_of_alien_all",
        $("#number_of_alien_all").val() || "0"
    );
    formData.append(
        "number_of_alien_male",
        $("#number_of_alien_male").val() || "0"
    );
    formData.append(
        "number_of_alien_female",
        $("#number_of_alien_female").val() || "0"
    );
    formData.append("start_age", $("#start_age").val());
    formData.append("end_age", $("#end_age").val() || "-");
    formData.append("start_height", $("#start_height").val() || "-");
    formData.append("end_height", $("#end_height").val());
    formData.append("start_weight", $("#start_weight").val() || "-");
    formData.append("end_weight", $("#end_weight").val() || "-");
    formData.append("workplace_addr_id", $("#name_address_work :selected").val());
    formData.append("work_duration_year", $("#year_work :selected").val() || "-");
    formData.append(
        "work_duration_month",
        $("#month_work :selected").val() || "-"
    );
    formData.append("work_duration_day", $("#day_work :selected").val() || "-");
    formData.append("work_duration_since_at", "-");
    formData.append("work_duration_until_at", "-");
    formData.append("workhour_per_day", $("#workhour_per_day").val() || "-");
    formData.append("dayoff_per_week", $("#dayoff_per_week").val() || "-");
    formData.append("holiday_per_year", $("#holiday_per_year").val() || "-");
    formData.append("wage_rate_per_day", $("#wage_rate_per_day").val());
    formData.append("fsc_id", $("#fsc_id :selected").val());
    formData.append(
        "immigration_checkpoint_id",
        $("#DataImmigrationCheckpoint :selected").val()
    );
    formData.append("is_saleperson", $("#issale").val());
    formData.append(
        "recruitment_agency_id",
        $("#recruitment_agency :selected").val()
    );

    if ($("#recruitment_agency_assign_at").val() != "") {
        var split_date = $("#recruitment_agency_assign_at").val().split("/");
        formData.append(
            "recruitment_agency_assign_dt",
            split_date[2] + "-" + split_date[1] + "-" + split_date[0]
        );
    }

    formData.append("job_number_of_alien_male", $("#All_Men").val() || "0");
    formData.append("job_number_of_alien_female", $("#All_Women").val() || "0");
    formData.append("traditional_holiday", $("#traditional_holiday").val());

    $.ajax({
        method: "post",
        processData: false,
        contentType: false,
        cache: false,
        data: formData,
        url: url_UpdateRequestDemandLetter41_4_59,
        success: function (response) {
            var dataresponse = JSON.parse(response);
            if (dataresponse.status == "success") {
                $("#loading").delay().fadeOut("");
                $("#payment").addClass("active");
                $("#summaries").addClass("active");
                GetDataAddNewRequestFormSuccess41($("#group_id").val());
            } else {
                $("#loading").delay().fadeOut("");
                $("#payment").removeClass("active");
                $("#summaries").addClass("active");
                $("#stepfour").css("display", "none");
                $("#stepthree").css("display", "block");
                $("#stepthree").css("opacity", "1");
                Swal.fire({
                    imageUrl: important_data.errorIconUrl,
                    title: `<span class="lang_failed_request">ยื่นคำขออนุญาตไม่สำเร็จ</span>`,
                    html: `<span>${dataresponse.msg}</span>`,
                    showCancelButton: false,
                    confirmButtonText: "<span class='lang_lbl_ok'>ตกลง</span>",
                    reverseButtons: true,
                    didRender: () => {
                        changeDatalang();
                    },
                }).then((result) => {
                    $("#payment").removeClass("active");
                    $("#summaries").addClass("active");
                    $("#stepfour").css("display", "none");
                    $("#stepthree").css("display", "block");
                    $("#stepthree").css("opacity", "1");
                });
            }
        },
        error: function (xhr) {
            $("#loading").delay().fadeOut("");
            $("#payment").removeClass("active");
            $("#summaries").addClass("active");
            $("#stepfour").css("display", "none");
            $("#stepthree").css("display", "block");
            $("#stepthree").css("opacity", "1");
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
    changeDatalang();
}
function GetDataAddNewRequestFormSuccess41(demand_letter_id) {
    let objRequest = {
        group_id: demand_letter_id,
        user_id: core_datas.current_profile.user_id,
    };

    var request = jQuery.post(
        url_GetDataAddNewRequestFormDLSuccess41,
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
                    "<span class='lang_lbl_year'></span>" +
                    " " +
                    objData.work_duration_month +
                    " " +
                    "<span class='lang_lbl_month'></span>" +
                    " " +
                    objData.work_duration_day +
                    " " +
                    "<span class='lang_lbl_day'></span>"
                );
                changeDatalang();
            }
        },
        "json"
    );
}
function GetDocFilePathByDocumentIDMOU(
    key_id_data,
    document_id,
    document_type_id
) {
    var app = important_data.is_mobile;
    let objRequest = {
        key_id: key_id_data,
        document_id: document_id,
        form_type_doc_id: document_type_id,
        version_id: "",
    };
    var request = jQuery.post(
        url_GetDocFilePathByDocumentID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                if (app == "Y") {
                    mobileScriptChannel.postMessage("openFile:" + objData.document_path);
                } else {
                    window.open(objData.document_path, "_blank");
                }
            }
        },
        "json"
    );
}
function ConvertDate(dateValue) {
    var sp_dateValue = dateValue.split("/");
    var resultDate =
        sp_dateValue[2] + "-" + sp_dateValue[1] + "-" + sp_dateValue[0];
    return resultDate;
}


async function GetDataRequestFormEtrackingMOU_ByGroupID() {
    $("#loading").css("display", "block");

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
            var setReFresh = "";
            if (data.status == "success") {
                var dataObj = data.data[0];
                var objData = data.data[0];
                statusGroup = dataObj.group_status_id;
                objData_setPayment = data.data[0];
                $("#statusGroup").val(statusGroup);


                setReFresh += `<div id='payment_recheck_notice' class='mb-3'><span class='payment_issue_text'>หากพบปัญหาการชำระเงิน ท่านสามารถส่งเรื่องเพื่อตรวจสอบการชำระเงินอีกครั้งได้ โดยกดปุ่ม</span> <span class='payment_check_button btncheckpayment' onclick='GetPaymentByGroupId("${objData.group_id}", this)'>“ตรวจสอบการชำระเงิน”</span></div>`;

                $("#payment_recheck_notice").remove();
                $(`<div>`)
                    .html(setReFresh)
                    .insertBefore("#div_no_payment");

                var demand_letter_data = dataObj.demand_letter_data;
                var objDataEmp = dataObj.demand_letter_data.employer_data;
                var agency_data = dataObj.demand_letter_data.agency_data;
                localStorage.setItem("demand_id", demand_letter_data.demand_letter_id);
                sessionStorage.setItem("group_status", objData.group_status_id);
                $("#number_formid").html(
                    "<span class='lang_lbl_req_number'></span> : " + objData.group_id
                );
                $("#title_form").html(
                    "<span class='lang_lbl_list'></span> : <span style='color:#11111F;'><span data_th='" +
                    objData.form_type_name_th +
                    "' data_en='" +
                    objData.form_type_name_en +
                    "'></span></span>"
                );

                $("#text_home").html(
                    "<span class='lang_first'></span> /"
                );
                $("#text_index").html(
                    '<a class="text-secondary font-weight-normal" href="/Permit/Tracking"><span class="lang_track_request_status">ติดตามสถานะคำขอ</span></a> /'
                );
                $("#text_header").html(
                    "<span class='lang_lbl_req_number'></span> :" + " " + objData.group_id
                );

                $("#date_re").html(
                    "<span class='lang_lbl_submit_date'></span> : <span style='color:#11111F;' data_th='" +
                    convert_to_thai_datetime(objData.created_timestamp) +
                    "' data_en='" +
                    convert_to_eng_datetime(objData.created_timestamp) +
                    "'></span>"
                );

                var getstatus = objData.group_status_id;
                var arraystatus_W = [
                    "WCOSNA",
                    "WCOSNA1",
                    "WCOCNA",
                    "WCOCNA1",
                    "WCOCAS",
                    "WCOCAS1",
                    "WCOCKKJ",
                    "WCOC1KKJ",
                    "WCOCOC",
                    "WCOCOC1",
                    "WCOSAS",
                    "WCOSAS1",
                    "WCOCOCNA",
                    "WCOCOCNA1",
                ];
                var arraystatus_W2 = [
                    "WCOSNA2",
                    "WCOSAS2",
                    "WCOCNA2",
                    "WCOCAS2",
                    "WCOCOC2",
                    "WCOCOCNA2",
                ];

                var status_data_AP = [
                    "AP",
                    "APPSS",
                    "APPRG2",
                    "APSS",
                    "APPP",
                    "APCC",
                    "APWD",
                    "APPRG2NL",
                    "WCPOS",
                    "WBI",
                ];
                var status_data_WA = ["WA"];
                var status_data_NC = ["NCOS2", "NCOC2"];
                var status_data_SS = ["SS", "CC", "CCPC", "WPC"];
                var status_data_CCOS = [
                    "CCOS2",
                    "CCOS",
                    "CCP3O",
                    "CCP2O",
                    "CCP4O",
                    "CCP5O",
                    "CCPO",
                ];
                var status_data_NCOS3 = ["NCOS3"];
                var status_data_WP = ["WP", "WP3"];
                var status_data_WC = ["WCOSNA3", "WCOSAS3"];
                var status_data_NCOS = [
                    "NCOS",
                    "NCOS1",
                    "NCOC",
                    "NCOC1",
                    "NCOCR1",
                    "NCOCRD1",
                ];
                var status_data_CCOS1 = ["CCOS1"];
                if (status_data_AP.includes(getstatus)) {
                    $("#btnaddEmp").css("display", "none");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#divpayment").css("display", "none");

                    if (getstatus == "AP") {
                        reloadQueueMOU();
                        openActiveSection("5", "active");
                    } else if (
                        getstatus == "APSS" ||
                        getstatus == "WCPOS" ||
                        getstatus == "WBI"
                    ) {
                        openActiveSection("3", "active");
                        $("#Alien_Table").css("display", "none");
                        $("#statusAPSS").css("display", "block");
                        $("#head_step_2").css("display", "none");
                    } else {
                        openActiveSection("5", "active");
                    }
                } else if (status_data_WA.includes(getstatus)) {
                    openActiveSection("3", "active");
                    $("#div_wait_tap_three").addClass("d-none");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#text_button").addClass("lang_attach_name_list");
                    $(".div_hidden_button").removeClass("d-none");
                    $("#btnpayment").removeClass("d-none");
                } else if (status_data_NC.includes(getstatus)) {
                    openActiveSection("3", "active");
                    $("#div_wait_tap_three").addClass("d-none");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#divpayment").hide();
                    $("#head_step_2").css("display", "flex");
                    $("#btnpayment").css("display", "none");
                    $("#text_button").addClass("lang_process_edit");
                } else if (arraystatus_W.includes(getstatus)) {
                    openActiveSection("1", "active");
                    $("#div_wait_tap_three").removeClass("d-none");
                } else if (arraystatus_W2.includes(getstatus)) {
                    openActiveSection("3", "active");
                    $("#div_wait_tap_three").addClass("d-none");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#btnaddEmp").css("display", "none");
                } else if (status_data_SS.includes(getstatus)) {
                    openActiveSection("1", "active");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#Alien_Table").css("display", "inline-table");
                    $("#btnpayment_3").css("display", "none");
                    if (getstatus == "CC" || getstatus == "WPC") {
                        $("#btnaddEmp").hide();
                    }

                } else if (status_data_CCOS.includes(getstatus)) {
                    openActiveSection("7", "active");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#Alien_Table").css("display", "inline-table");
                    $("#addEmp").css("display", "none");
                    $("#btnaddEmp").css("display", "none");


                    $(".div_hidden_appeal_button").removeClass("d-none");
                    $(".div_hidden_button").addClass("d-none");
                    $("#reject_appeal").show();
                    AppealList(group_id);
                } else if (status_data_NCOS3.includes(getstatus)) {
                    openActiveSection("3", "active");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#divpayment").css("display", "none");
                    $("#Alien_Table").css("display", "none");
                    $("#statusAPSS").css("display", "block");
                    $("#head_step_2").css("display", "none");

                    $("#Alien_Table_None").css("display", "none");
                } else if (status_data_WP.includes(getstatus)) {
                    openActiveSection("4", "active");
                    $("#div_wait_tap_three").addClass("d-none");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#divpayment").hide();
                    $("#Alien_Table").css("display", "inline-table");
                    $("#btnaddEmp").css("display", "none");
                    $(".div_hidden_button").removeClass("d-none");
                    $("#text_button").addClass("lang_make_payment");

                    if (
                        (dataObj.request_by_user_id == core_datas.user_id &&
                            dataObj.request_by_user_id_proxy == "") ||
                        dataObj.request_by_user_id_proxy == core_datas.user_id
                    ) {
                        $(".show_for_proxy").css("display", "none");
                        GetDataPaymentInquiryTransactionMOU(dataObj.group_id, false);
                    } else {
                        GetDataPaymentInquiryTransactionMOU(dataObj.group_id, true);
                    }

                } else if (status_data_WC.includes(getstatus)) {
                    openActiveSection("3", "active");
                    $("#div_all_tap_three").removeClass("d-none");
                    $("#divpayment").css("display", "none");
                    $("#Alien_Table").css("display", "none");
                    $("#statusAPSS").css("display", "block");
                    $("#head_step_2").css("display", "none");

                    $("#div_wait_tap_three").addClass("d-none");
                } else if (status_data_NCOS.includes(getstatus)) {
                    openActiveSection("1", "active");
                    $(".div_hidden_button").removeClass("d-none");
                    $("#div_wait_tap_three").removeClass("d-none");
                    $("#text_button").addClass("lang_process_edit");
                } else if (status_data_CCOS1.includes(getstatus)) {
                    openActiveSection("1", "active");
                    $("#div_all_tap_three").addClass("d-none");
                    $("#div_wait_tap_three").removeClass("d-none");
                } else {
                    openActiveSection("1", "active");
                }

                if (objData.is_appointment == "N") {
                    if (getstatus == "AP") {
                        $("#div_appointment_true").removeClass("d-none");
                    } else {
                        $("#div_appointment_true").addClass("d-none");
                        $(".div_none_ap").removeClass("d-none");
                    }
                } else {
                    $(".div_none_ap").addClass("d-none");
                    $("#div_appointment_true").removeClass("d-none");
                    GetDataFormAppointmentMoreDocByGroupID();
                }

                var arraystatuscancel = [
                    "WA",
                    "WP",
                    "WP3",
                    "WP2",
                    "NCOS",
                    "NCOS1",
                    "NCOC",
                    "NCOC1",
                    "AP",
                    "NCOS2",
                    "NCOC2",
                    "NCOS3",
                    "NCOCR1",
                    "NCOCRD1",
                    "NCOCR2",
                    "NCOCRD2",
                ];
                var arraystatusnocancel = [
                    "NCOS",
                    "NCOS1",
                    "NCOC",
                    "NCOC1",
                    "NCOCR1",
                    "NCOCRD1",
                ];

                if (arraystatuscancel.includes(getstatus)) {
                    $(".cancelRequestForm").removeClass("d-none");
                } else {
                    $(".cancelRequestForm").addClass("d-none");
                }

                $("#transection_activity_th").attr(
                    "data_th",
                    objData.transection_activity_th
                );
                $("#transection_activity_th").attr(
                    "data_en",
                    objData.transection_activity_en
                );
                $("#tracking_desc_th").html(
                    "<span class='lang_text_reason'></span> : " +
                    "<span data_th='" +
                    objData.tracking_desc_th?.toSafeJsonHtml() +
                    "' data_en='" +
                    objData.tracking_desc_en +
                    "'></span>"
                );


                $("#nationality_al").val(demand_letter_data.nationality_id);

                var arraystatus_cc = [
                    "CC",
                    "CCOS",
                    "CCOS1",
                    "CCOS2",
                    "NCOC",
                    "NCOC1",
                    "NCOC2",
                    "NCOS",
                    "NCOS1",
                    "NCOS2",
                    "NCOCR1",
                    "NCOCRD1",
                    "NCOCR2",
                    "NCOCRD2",
                    "CCP2O",
                    "CCP3O",
                    "CCP4O",
                    "CCP5O",
                    "CCPO",
                ];
                var arraystatus_ss = ["SS"];
                var arraystatus_ap = ["APSS"];
                $("#form_activity_detail").addClass("badgestatus");
                $("#form_activity_detail").css({
                    color: objData.status_color_font,
                    background: objData.status_color,
                });
                if (arraystatus_cc.includes(objData.group_status_id)) {
                    $("#divforreasonshowdetail").show();
                } else if (arraystatus_ss.includes(objData.group_status_id)) {
                    $("#line_status").hide();
                } else if (arraystatus_ap.includes(objData.group_status_id)) {
                    $("#line_status").hide();
                } else {
                    $("#line_status").hide();
                    $("#divforreasonshowdetail").hide();
                }

                $("#form_activity_detail").attr(
                    "data_th",
                    objData.transection_status_th
                );
                $("#form_activity_detail").attr(
                    "data_en",
                    objData.transection_status_en
                );
                $("#div_DataEmployer").css("display", "flex");
                if (objDataEmp.emp_type_id == "1") {
                    $(".name_th_employer").addClass("lang_txt_employer_name_th");
                    $(".name_en_employer").addClass("lang_txt_employer_name_eng");

                    if (objDataEmp.firstname_th) {
                        $(".full_name_th").text(
                            objDataEmp.prefix_name_th +
                            objDataEmp.firstname_th +
                            " " +
                            objDataEmp.middlename_th +
                            " " +
                            objDataEmp.lastname_th
                        );
                    } else {
                        $(".full_name_th").text("-");
                    }

                    if (objDataEmp.firstname_en) {
                        $(".full_name_en").text(
                            objDataEmp.prefix_name_en +
                            objDataEmp.firstname_en +
                            " " +
                            objDataEmp.middlename_en +
                            " " +
                            objDataEmp.lastname_en
                        );
                    } else {
                        $(".full_name_en").text("-");
                    }
                    if (objDataEmp.u_map_id == "8" && objDataEmp.user_type_id == "4") {
                        $(".text_employer_type").addClass("lang_foreign_individual_label");
                        $(".no_employer").addClass("lang_tax_id");
                    } else {
                        $(".text_employer_type").attr(
                            "data_th",
                            objDataEmp.emp_type_name_th
                        );
                        $(".text_employer_type").attr(
                            "data_en",
                            objDataEmp.emp_type_name_en
                        );
                        $(".no_employer").addClass("lang_id_card_and_legal_entity_number");
                    }
                } else if (objDataEmp.emp_type_id == "2") {
                    $(".name_th_employer").addClass("lang_bu_name_th");
                    $(".name_en_employer").addClass("lang_bu_name_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);

                    if (objDataEmp.u_map_id == "9" && objDataEmp.user_type_id == "4") {
                        $(".text_employer_type").addClass("lang_registered_in_thailand");
                        $(".no_employer").addClass("lang_tax_id");
                    } else {
                        $(".no_employer").addClass("lang_id_card_and_legal_entity_number");
                        $(".text_employer_type").attr(
                            "data_th",
                            objDataEmp.emp_type_name_th
                        );
                        $(".text_employer_type").attr(
                            "data_en",
                            objDataEmp.emp_type_name_en
                        );
                    }
                } else if (
                    objDataEmp.emp_type_id == "4" ||
                    objDataEmp.emp_type_id == "5" ||
                    objDataEmp.emp_type_id == "6"
                ) {
                    $(".no_employer").addClass("lang_government_no");
                    $(".name_th_employer").addClass("lang_name_organization_th");
                    $(".name_en_employer").addClass("lang_name_organization_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);
                    $(".text_employer_type").attr("data_th", objDataEmp.emp_type_name_th);
                    $(".text_employer_type").attr("data_en", objDataEmp.emp_type_name_en);
                } else if (objDataEmp.emp_type_id == "8") {
                    $(".no_employer").addClass("lang_government_no");
                    $(".name_th_employer").addClass("lang_bu_name_th");
                    $(".name_en_employer").addClass("lang_bu_name_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);
                    $(".text_employer_type").attr("data_th", objDataEmp.emp_type_name_th);
                    $(".text_employer_type").attr("data_en", objDataEmp.emp_type_name_en);
                } else if (objDataEmp.emp_type_id == "7") {
                    $(".no_employer").addClass("lang_government_no");
                    $(".name_th_employer").addClass("lang_company_th");
                    $(".name_en_employer").addClass("lang_company_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);
                    $(".text_employer_type").attr("data_th", objDataEmp.emp_type_name_th);
                    $(".text_employer_type").attr("data_en", objDataEmp.emp_type_name_en);
                } else if (objDataEmp.emp_type_id == "9") {
                    $(".no_employer").addClass("lang_tax_id");
                    $(".name_th_employer").addClass("lang_bu_name_th");
                    $(".name_en_employer").addClass("lang_bu_name_en");
                    $(".full_name_th").text(objDataEmp.companyname_th);
                    $(".full_name_en").text(objDataEmp.companyname_en);
                    if (objDataEmp.u_map_id == "10" && objDataEmp.user_type_id == "4") {
                        $(".text_employer_type").addClass("lang_registered_abroad");
                    } else {
                        $(".text_employer_type").attr(
                            "data_th",
                            objDataEmp.emp_type_name_th
                        );
                        $(".text_employer_type").attr(
                            "data_en",
                            objDataEmp.emp_type_name_en
                        );
                    }
                }

                $(".text_employer_idcard").text(objDataEmp.employer_id);
                $("#text_employer_idcard").text(objDataEmp.employer_id);
                $(".text_employer_id").text(objDataEmp.employer_id);
                $(".text_employer_no").text(objDataEmp.employer_no);


                if (objData.demand_letter_data.proxy_id) {
                    $(".div_agent").removeClass("d-none");
                    var dataproxy = objData.demand_letter_data.proxy_data;
                    if (dataproxy.firstname_th) {
                        $(".text_namethai_agent").text(
                            dataproxy.prefix_name_th +
                            dataproxy.firstname_th +
                            " " +
                            dataproxy.midname_th +
                            " " +
                            dataproxy.lastname_th
                        );
                    } else {
                        $(".text_namethai_agent").text("-");
                    }

                    if (dataproxy.firstname_en) {
                        $(".text_nameeng_agent").text(
                            dataproxy.prefix_name_en +
                            dataproxy.firstname_en +
                            " " +
                            dataproxy.midname_en +
                            " " +
                            dataproxy.lastname_en
                        );
                    } else {
                        $(".text_nameeng_agent").text("-");
                    }

                    $(".text_idcard_agent").text(dataproxy.id_card);
                }

                if (dataObj.formlist.length != 0) {
                    $("#numberform").text(dataObj.formlist.length);
                    $("#numberpayment").text(dataObj.formlist.length);
                    $("#numberpayment2").text(dataObj.formlist.length);
                    $("#alien_show_num").text(dataObj.formlist.length);
                    var status = dataObj.group_status_id;
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

                        if (dataObj.formlist.length > 0) {
                            for (var f = 0; f < dataObj.formlist.length; f++) {
                                if (dataObj.formlist[f].form_status_id === statusDefult) {
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


                        $("#head_step_2").show();

                        if (dataObj.formlist.length > 0) {
                            $("#Alien_Table_None").css("display", "none");
                            $("#Alien_Table").show();
                        } else {
                            $("#Alien_Table_None").css("display", "block");
                            $("#Alien_Table").hide();
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

                        if (dataObj.formlist.length > 0) {
                            for (var f = 0; f < dataObj.formlist.length; f++) {
                                if (dataObj.formlist[f].form_status_id === statusDefult) {
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
                                    return `<span>${row.alien_fullname_en}</span><br><span style="color:#808080;"><span class="lang_ref_code">หมายเลขอ้างอิง</span> : ${row.alien_id}</span>`;
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
                                    return `<div class='badge badge-pill' style='word-wrap: break-word;white-space: normal; font-size: 14px; color:${row.status_color_font
                                        }; background-color:${row.status_color
                                        }' data_th='${row.transection_status_th?.toSafeJsonHtml()}' data_en='${row.transection_status_en
                                        }' id='${row.form_id}'></div>`;
                                },
                            },

                            {
                                targets: 5,
                                render: function (data, type, row, meta) {
                                    var objData_data = dataObj;
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
                                    var getstatus = dataObj.group_status_id;


                                    var objData_data = dataObj;
                                    if (row.form_status_id == "WA") {
                                        return `<a href="javascript:void(0)" onclick="DeleteDataNameList(&quot;${objData_data.group_id}&quot;,&quot;${row.form_id}&quot;,&quot;${objData_data.request_by_user_id}&quot;)" style="vertical-align:middle;">
                                            <img src='/assets/images/cancel_namelist.png'>
                                       </a>`;
                                    } else if (
                                        row.form_status_id == "NCOS2" ||
                                        row.form_status_id == "NCOC2" ||
                                        row.form_status_id == "NCOCR2" ||
                                        row.form_status_id == "NCOCRD2" ||
                                        row.form_status_id == "NCOS3" ||
                                        row.form_status_id == "AP"
                                    ) {
                                        return `<div class="dropdown show">
  <a class="btn btn-secondary dropdown-toggle button-dropdown-height" href="#" role="button" id="dropdownMenuLink" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
   <i class="fa-solid fa-ellipsis-vertical"></i>
  </a>  <div class="dropdown-menu" aria-labelledby="dropdownMenuLink">
    <a href="javascript:void(0)" class="dropdown-item offcanvas-edit-toggle lang_watch_detail" style="font-weight:700" onclick="UploadFileMOU(&quot;${row.form_id}&quot;,&quot;${row.form_status_id}&quot;,&quot;${getstatus}&quot;)">ดูรายละเอียด</a>
                                                                  <a href="javascript:void(0)" class="dropdown-item lang_reject_work" id="cancelAssign"  style="color:#FF3B30;font-weight:700" onclick="UpdateRejectAcceptWork(&quot;${objData_data.group_id}&quot;,&quot;${row.form_id}&quot;,&quot;${objData_data.request_by_user_id}&quot;)">ปฏิเสธการจ้างงาน</a>
  </div>
</div>`;
                                    } else {
                                        return `<div class="dropdown show">
  <a class="btn btn-secondary dropdown-toggle button-dropdown-height" href='javascript:void(0)' role="button" aria-haspopup="true" aria-expanded="false" onclick="UploadFileMOU(&quot;${row.form_id}&quot;,&quot;${row.form_status_id}&quot;,&quot;${getstatus}&quot;)">
   <i class="fa-solid fa-ellipsis-vertical"></i>
  </a>


</div>`;
                                    }
                                },
                            },
                        ],

                        data: dataObj.formlist,
                        drawCallback: function (settings) {

                            changeDatalang();
                            if (dataObj.group_status_id == "CC") {
                                GetCancelComment(dataObj.group_id, dataObj.form_id, dataObj.form_type_id, "", dataObj.formlist);
                            }

                        },
                    });

                } else {
                    $("#divpayment").css("display", "none");
                    $("#divpayment_2").css("display", "none");

                    if (getstatus == "APSS") {
                        $("#Alien_Table").css("display", "none");
                        $("#Alien_Table_None").hide();
                        $("#head_step_2").css("display", "none");
                    } else if (getstatus == "WA") {
                        $("#div_wait_tap_three").addClass("d-none");
                    } else {
                        $("#div_wait_tap_three").removeClass("d-none");
                    }
                }

                roundedNumber = convertToMonths(
                    demand_letter_data.work_duration_year,
                    demand_letter_data.work_duration_month,
                    demand_letter_data.work_duration_day
                );
                for (var f = 0; f < dataObj.formlist.length; f++) {
                    if (!list_alien_id.includes(dataObj.formlist[f].alien_id)) {
                        if (dataObj.formlist[f].form_status_id == "WA") {
                            list_alien_id.push(dataObj.formlist[f].alien_id);
                        }

                    }
                }


                if (objDataEmp.addrReqList.length != 0) {
                    $(".addrReqList_emp").text(
                        objDataEmp.addrReqList[0].employer_full_address_th
                    );
                } else {
                    $(".addrReqList_emp").text("-");
                }

                $(".tel_em").text(objDataEmp.phone || "-");
                $(".email_em").text(objDataEmp.email || "-");
                $(".lbl_fax_em").text(objDataEmp.fax || "-");
                if (objDataEmp.addrCurList.length != 0) {
                    $(".addrCurList_emp").text(
                        objDataEmp.addrCurList[0].employer_full_address_th
                    );
                } else {
                    $(".addrCurList_emp").text("-");
                }

                var setData = "";
                setData += '<option selected="" value="">กรุณาเลือก</option>';
                for (var i = 0; i < objDataEmp.addrWrkList.length; i++) {
                    var StringData = JSON.stringify(
                        objDataEmp.addrWrkList[i].employer_addr_id
                    );
                    setData +=
                        '<option value="' +
                        objDataEmp.addrWrkList[i].employer_addr_id +
                        '" data-value=' +
                        StringData +
                        ">" +
                        objDataEmp.addrWrkList[i].employer_full_address_th +
                        "</option>";
                }
                $("#name_address_work").html(setData);

                if (objDataEmp.addrWrkList.length == 1) {
                    $("#name_address_work").prop("disabled", true);
                    $("#name_address_work").prop("selectedIndex", 1);
                } else {
                    $("#name_address_work").prop("disabled", false);
                }

                $(".juristic_permit_no_detail").text(
                    agency_data.juristic_permit_no || "-"
                );
                if (agency_data.juristic_permit_issue_dt != "") {
                    $(".juristic_permit_issue_dt_detail").attr(
                        "data_th",
                        convert_to_thai_date(agency_data.juristic_permit_issue_dt)
                    );
                    $(".juristic_permit_issue_dt_detail").attr(
                        "data_en",
                        convert_to_eng_date(agency_data.juristic_permit_issue_dt)
                    );
                } else {
                    $(".juristic_permit_issue_dt_detail").addClass("lang_not_specified");
                }

                if (agency_data.juristic_permit_until_dt != "") {
                    $(".juristic_permit_until_dt_detail").attr(
                        "data_th",
                        convert_to_thai_date(agency_data.juristic_permit_until_dt)
                    );
                    $(".juristic_permit_until_dt_detail").attr(
                        "data_en",
                        convert_to_eng_date(agency_data.juristic_permit_until_dt)
                    );
                } else {
                    $(".juristic_permit_until_dt_detail").addClass("lang_not_specified");
                }

                var arraystatus_wait = [
                    "WCOSNA2",
                    "WCOSNA",
                    "WCOSNA1",
                    "WCOSNA3",
                    "WCOSAS",
                    "WCOSAS1",
                    "WCOSAS3",
                    "WCOSAS2",
                    "WCOCNA2",
                    "WCOCNA1",
                    "WCOCNA",
                    "WCOCNA1",
                    "WCOCOCNA1",
                    "WCOCOCNA2",
                    "WCOCOCNA",
                    "WCOCOC",
                    "WCOCOC1",
                    "WCOCOC2",
                    "WCOC1KKJ",
                    "WCOCKKJ",
                    "AP",
                    "WCOCAS",
                    "WCOCAS1",
                ];

                var arraystatus_wait2 = [
                    "WCOSNAR1SS",
                    "WCOSNA3RSS",
                    "WCOSNA2RSS",
                    "WP3ADSS",
                    "WCOSNAR5SS",
                    "WCOCAS1POS",
                    "WCOCAS2POS",
                    "WCOCASPOS",
                    "WCOCNA2POS",
                    "WCOCNAPOS",
                    "WCOCNA1POS",
                    "WCOCOC1PAS",
                    "WCOCOCPAS",
                    "WCOCOCNA1PAS",
                    "WCOCOC2PAS",
                    "WCOCOCNAPAS",
                    "WCOCOCNA2PAS",
                    "WAPRG",
                    "WADPRG",
                    "SS2",
                    "WPPRG2",
                    "WPPRG",
                    "WAPRG",
                    "WP5PRG",
                    "APPRG2",
                    "SS3",
                    "WCOSNA2P3SS",
                    "WP6SS",
                    "PS",
                    "APPSS",
                    "APSS",
                    "WCDAPSS",
                    "SS",
                    "WCPOS",
                    "WPC",
                ];

                var setTimeline = "";
                for (var t = 0; t < dataObj.timelinelist.length; t++) {

                    setTimeline +=
                        '<div class="col-xl-3 col-lg-2 col-sm-2 col-12 step-status ' +
                        (t === dataObj.timelinelist.length - 1 ? "hide-before" : "") +
                        '">';
                    setTimeline +=
                        '<p class="label-form-info step"><span data_th="' +
                        convert_to_thai_datetime(
                            dataObj.timelinelist[t].created_timestamp
                        ) +
                        '" data_en="' +
                        convert_to_eng_datetime(dataObj.timelinelist[t].created_timestamp) +
                        '"></span></p>';
                    setTimeline += "</div>";
                    setTimeline +=
                        '<div class="col-xl-3 col-lg-3 col-sm-3 col-6">' +
                        '<p class="form-info" data_th="' +
                        dataObj.timelinelist[t].tracking_status_th?.toSafeJsonHtml() +
                        '">' +
                        dataObj.timelinelist[t].tracking_status_th +
                        "</p></div>";


                    setTimeline +=
                        '<div class="col-xl-6 col-lg-7 col-sm-7 col-6"><span class="badgestatus" style="color:' +
                        dataObj.timelinelist[t].status_color_font +
                        ";background-color:" +
                        dataObj.timelinelist[t].status_color +
                        '"><span data_th="' +
                        dataObj.timelinelist[t].tracking_activity_th?.toSafeJsonHtml() +
                        '" data_en="' +
                        dataObj.timelinelist[t].tracking_activity_en +
                        '"></span></span></div>';
                }
                $("#Timeline").html(setTimeline);

                if (
                    important_data.is_mobile == "Y" ||
                    important_data.is_mobile == "y"
                ) {
                    $(".step-status").removeClass("step-status");
                }

                sessionStorage.setItem(
                    "demand_id",
                    demand_letter_data.demand_letter_id
                );

                $("#employer_id").val(objDataEmp.employer_id);
                $("#emp_type_id").val(objDataEmp.emp_type_id);
                $("#emp_type_name_th").val(objDataEmp.emp_type_name_th);
                $("#user_id").val(objDataEmp.user_id);

                $(".text_workhour_per_day").html(
                    demand_letter_data.workhour_per_day +
                    " " +
                    "<span class='lang_lbl_hour'></span>"
                );
                $(".text_dayoff_per_week").html(
                    demand_letter_data.dayoff_per_week +
                    " " +
                    "<span class='lang_lbl_day'></span>"
                );
                $(".text_holiday_per_year").html(
                    demand_letter_data.holiday_per_year +
                    " " +
                    "<span class='lang_lbl_day'></span>"
                );
                $(".text_wage_rate_per_day").html(
                    demand_letter_data.wage_rate_per_day +
                    " " +
                    "<span class='lang_bath_per_day'></span>"
                );
                $(".text_traditional_holiday").html(
                    demand_letter_data.traditional_holiday +
                    " " +
                    "<span class='lang_lbl_day'></span>"
                );

                var year = demand_letter_data.work_duration_year;
                var month = demand_letter_data.work_duration_month;
                var day = demand_letter_data.work_duration_day;
                if (year == "0" && month !== "0" && day !== "0") {
                    $(".text_year_month_day_work").html(
                        month +
                        " " +
                        "<span class='lang_lbl_month'></span>" +
                        " " +
                        day +
                        " " +
                        "<span class='lang_lbl_day'></span>"
                    );
                } else if (year !== "0" && month == "0" && day !== "0") {
                    $(".text_year_month_day_work").html(
                        year +
                        " " +
                        "<span class='lang_lbl_year'></span>" +
                        " " +
                        day +
                        " " +
                        "<span class='lang_lbl_day'></span>"
                    );
                } else if (year !== "0" && month !== "0" && day == "0") {
                    $(".text_year_month_day_work").html(
                        year +
                        " " +
                        "<span class='lang_lbl_year'></span>" +
                        " " +
                        month +
                        " " +
                        "<span class='lang_lbl_month'></span>"
                    );
                } else if (year !== "0" && month == "0" && day == "0") {
                    $(".text_year_month_day_work").html(
                        year + " " + "<span class='lang_lbl_year'></span>"
                    );
                } else if (year == "0" && month !== "0" && day == "0") {
                    $(".text_year_month_day_work").html(
                        month + " " + "<span class='lang_lbl_month'></span>"
                    );
                } else if (year == "0" && month == "0" && day !== "0") {
                    $(".text_year_month_day_work").html(
                        day + " " + "<span class='lang_lbl_day'></span>"
                    );
                } else {
                    $(".text_year_month_day_work").html(
                        year +
                        " " +
                        "<span class='lang_lbl_year'></span>" +
                        " " +
                        month +
                        " " +
                        "<span class='lang_lbl_month'></span>" +
                        " " +
                        day +
                        " " +
                        "<span class='lang_lbl_day'></span>"
                    );
                }

                $(".text_number_of_alien_all").html(
                    "<span>" +
                    demand_letter_data.number_of_alien_all +
                    "</span> <span class='lang_person'></span>"
                );
                $(".text_number_of_alien_female").html(
                    "<span>" +
                    demand_letter_data.number_of_alien_female +
                    "</span> <span class='lang_person'></span>"
                );
                $(".text_number_of_alien_male").html(
                    "<span>" +
                    demand_letter_data.number_of_alien_male +
                    "</span> <span class='lang_person'></span>"
                );

                $(".text_nationality_id_mou").attr(
                    "data_th",
                    demand_letter_data.nationality_th
                );
                $(".text_nationality_id_mou").attr(
                    "data_en",
                    demand_letter_data.nationality_en
                );
                $(".text_recruitment_agency").attr(
                    "data_th",
                    demand_letter_data.recruitment_agency_name_th
                );
                $(".text_recruitment_agency").attr(
                    "data_en",
                    demand_letter_data.recruitment_agency_name_en
                );
                $(".text_address_recruitment").text(
                    demand_letter_data.recruitment_agency_address
                );

                $(".text_recruitment_agency_assign_at").attr(
                    "data_th",
                    convert_to_thai_date(demand_letter_data.recruitment_agency_assign_at)
                );
                $(".text_recruitment_agency_assign_at").attr(
                    "data_en",
                    convert_to_eng_date(demand_letter_data.recruitment_agency_assign_at)
                );

                var workplace_addr_id = demand_letter_data.workplace_addr_id;
                for (var w = 0; w < objDataEmp.addrWrkList.length; w++) {
                    if (workplace_addr_id == objDataEmp.addrWrkList[w].employer_addr_id) {
                        $(".text_name_address_work").attr(
                            "data_th",
                            objDataEmp.addrWrkList[w].employer_full_address_th
                        );
                        $(".text_name_address_work").attr(
                            "data_en",
                            objDataEmp.addrWrkList[w].employer_full_address_en
                        );
                    }
                }

                $("#bus_type_id").attr("data_th", demand_letter_data.bus_type_name_th);
                $("#bus_type_id").attr("data_en", demand_letter_data.bus_type_name_en);
                $("#bus_type_id_value").val(demand_letter_data.bus_type_id);
                $("#hidden_bus_type_id").val(demand_letter_data.bus_type_id);
                $("#hidden_bus_type_text").val(demand_letter_data.bus_type_name_th);
                $(".text_bus_type_id").attr(
                    "data_th",
                    demand_letter_data.bus_type_name_th
                );
                $(".text_bus_type_id").attr(
                    "data_en",
                    demand_letter_data.bus_type_name_en
                );

                $(".text_age").html(
                    "<span class='lang_from'> </span>" +
                    " " +
                    demand_letter_data.start_age +
                    " " +
                    "<span class='lang_lbl_to'>ถึง</span>" +
                    " " +
                    demand_letter_data.end_age +
                    " " +
                    "<span class='lang_lbl_year'></span>"
                );
                $(".text_height").html(
                    "<span class='lang_from'>ตั้งเเต่</span>" +
                    " " +
                    demand_letter_data.start_height +
                    " " +
                    "<span class='lang_lbl_to'>ถึง</span>" +
                    " " +
                    demand_letter_data.end_height +
                    " " +
                    "<span class='lang_lbl_cm'></span>"
                );
                $(".text_weight").html(
                    "<span class='lang_from'>ตั้งเเต่</span>" +
                    " " +
                    demand_letter_data.start_weight +
                    " " +
                    "<span class='lang_lbl_to'>ถึง</span>" +
                    " " +
                    demand_letter_data.end_weight +
                    " " +
                    "<span class='lang_lbl_kg'></span>"
                );

                $(".text_DataImmigrationCheckpoint").attr(
                    "data_th",
                    demand_letter_data.immigration_checkpoint_name_th
                );
                $(".text_DataImmigrationCheckpoint").attr(
                    "data_en",
                    demand_letter_data.immigration_checkpoint_name_en
                );
                $(".text_fsc_id").attr("data_th", demand_letter_data.fsc_name_th);
                $(".text_fsc_id").attr("data_en", demand_letter_data.fsc_name_en);
                if (objDataEmp.addrCurList.length > 0) {
                    $(".textaddress_em").attr(
                        "data_th",
                        objDataEmp.addrCurList[0].employer_full_address_th
                    );
                    $(".textaddress_em").attr(
                        "data_en",
                        objDataEmp.addrCurList[0].employer_full_address_en
                    );
                }

                $("#nationality_text").attr(
                    "data_th",
                    demand_letter_data.nationality_th
                );
                $("#nationality_text").attr(
                    "data_en",
                    demand_letter_data.nationality_en
                );

                $(".company_regis_date").attr(
                    "data_th",
                    convert_to_thai_date(agency_data.company_regis_date)
                );
                $(".company_regis_date").attr(
                    "data_en",
                    convert_to_eng_date(agency_data.company_regis_date)
                );

                $(".company_objt").text(agency_data.company_objt || "-");
                $(".email").text(agency_data.email || "-");
                $(".tel_em").text(agency_data.phone || "-");
                $(".lbl_fax").text(agency_data.fax || "-");
                $(".agency_no").text(agency_data.agency_no || "-");
                $(".company_thai").text(agency_data.companyname_th);
                $(".company_eng").text(agency_data.companyname_en);

                if (agency_data.addrWrkList.length > 0) {
                    $(".addrWrkList").attr(
                        "data_th",
                        agency_data.addrWrkList[0].fulladdress_th
                    );
                    $(".addrWrkList").attr(
                        "data_en",
                        agency_data.addrWrkList[0].fulladdress_en
                    );
                }
                sessionStorage.setItem(
                    "request_by_user_id",
                    dataObj.request_by_user_id
                );

                $(".text_employer_type41").attr(
                    "data_th",
                    agency_data.user_type_name_th
                );
                $(".text_employer_type41").attr(
                    "data_en",
                    agency_data.user_type_name_en
                );
                $(".company_type_id").text(agency_data.company_type_id);
                $(".company_status_id").text(agency_data.company_status_id);
                $(".company_id").text(agency_data.company_id);
                $(".company_regis_capital").text(agency_data.company_regis_capital);
                $(".director_fullname").html(
                    agency_data.director_fullname.replace(",", "<br>")
                );

                var demand_letter_job_list = demand_letter_data.demand_letter_job_list;
                var setdatajob = "";
                var num = 1;
                for (var c = 0; c < demand_letter_job_list.length; c++) {
                    setdatajob += "<tr>";
                    setdatajob += '<td data-label="ลำดับ">' + num++ + "</td>";
                    if (
                        demand_letter_job_list[c].permit_catework_name_th == "งานอื่น ๆ"
                    ) {
                        setdatajob +=
                            '<td data-label="ลักษณะงาน" class="text-left">งานอื่นๆ : <span data_th="' +
                            demand_letter_job_list[c].permit_catework_name_th +
                            '" data_en="' +
                            demand_letter_job_list[c].permit_catework_name_en +
                            '"></span></td>';
                    } else {
                        setdatajob +=
                            '<td data-label="ลักษณะงาน" class="text-left"><span data_th="' +
                            demand_letter_job_list[c].permit_catework_name_th +
                            '" data_en="' +
                            demand_letter_job_list[c].permit_catework_name_en +
                            '"></span></td>';
                    }

                    setdatajob +=
                        '<td data-label="เพศชาย" class="link-table text-md-right text-left">' +
                        demand_letter_job_list[c].number_of_alien_male +
                        "</td>";
                    setdatajob +=
                        '<td data-label="เพศหญิง" class="link-table text-md-right text-left">' +
                        demand_letter_job_list[c].number_of_alien_female +
                        "</td>";
                    setdatajob += "</tr>";
                }
                $("#tabelshowtypejob").html(setdatajob);

                var objdoclist = demand_letter_data.demand_letter_doc_list;
                var setDataemp = "";
                var index_array = 0;
                var EmployerID = objDataEmp.employer_id;


                $("#head_employer").html('<span class="lang_menu_helpservice"></span> : <span data_th="' + objDataEmp.emp_type_name_th + '" data_en="' + objDataEmp.emp_type_name_en + '"></span>')
                for (let a = 0; a < objdoclist.employer_n_doclist.length; a++) {
                    const requestDocList =
                        objdoclist.employer_n_doclist[a].requestdoclist;
                    setDataemp += "<tr>";
                    setDataemp += `<td data-label='ลำดับ'>${a + 1}</td>`;
                    for (let k = 0; k < requestDocList.length; k++) {
                        if (k != index_array) {
                            setDataemp += `<td data-label='ลำดับ'></td>`;
                        }
                        setDataemp += `<td data-label='รายการ' class='text-left'><span data_th='${requestDocList[k].document_name_th?.toSafeJsonHtml()
                            } ' data_en='${requestDocList[k].document_name_en} '></span></td>`;
                        if (requestDocList[k].doclist.length > 0) {
                            setDataemp +=
                                "<td  class='link-table'><a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                EmployerID +
                                " &quot;, &quot; " +
                                requestDocList[k].doclist[0].document_id +
                                " &quot;, &quot; " +
                                requestDocList[k].document_type_id +
                                " &quot;)'></a></td>";
                        } else {
                            setDataemp += `<td  class='link-table' tyle='text-decoration: none;'>-</td>`;
                        }

                        setDataemp += "</tr>";
                    }

                    jQuery("#listuploadfile_employer46_59").html(setDataemp);

                    var setproxy = "";
                    if (objdoclist.proxy_doclist.length != 0) {
                        $(".show_proxy").removeClass("d-none");
                        for (let a = 0; a < objdoclist.proxy_doclist.length; a++) {
                            const requestDocList = objdoclist.proxy_doclist[a].requestdoclist;
                            setproxy += "<tr>";
                            setproxy += `<td data-label='ลำดับ'>${a + 1}</td>`;
                            for (let k = 0; k < requestDocList.length; k++) {
                                if (k != index_array) {
                                    setproxy += `<td data-label='ลำดับ'></td>`;
                                }
                                setproxy += `<td data-label='รายการ' class='text-left' data_th='${requestDocList[k].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[k].document_name_en}'></td>`;
                                if (requestDocList[k].doclist.length > 0) {
                                    setproxy +=
                                        "<td data-label='เอกสาร' class='link-table'><a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                        EmployerID +
                                        " &quot;,&quot; " +
                                        requestDocList[k].doclist[0].document_id +
                                        " &quot;, &quot; " +
                                        requestDocList[k].document_type_id +
                                        " &quot;)'></a></td>";
                                } else {
                                    setproxy += `<td data-label='เอกสาร' class='link-table' tyle='text-decoration: none;'>-</td>`;
                                }

                                setproxy += "</tr>";
                            }
                        }
                        jQuery("#listuploadfile_proxy").html(setproxy);
                    }

                    if (objdoclist.other_doclist[0].requestdoclist[0].doclist.length > 0) {
                        $("#div_showFileOrther59 , .file_or_div").removeClass("d-none");
                        var setDataorther = "";
                        var numor = 1;
                        for (let r = 0; r < objdoclist.other_doclist.length; r++) {
                            const requestDocOR =
                                objdoclist.other_doclist[r].requestdoclist[0].doclist;
                            var doc_type = objdoclist.other_doclist[r].requestdoclist[0];
                            for (let n = 0; n < requestDocOR.length; n++) {
                                setDataorther += "<tr>";
                                setDataorther += `<td data-label='ลำดับ'>${numor++}</td>`;
                                setDataorther += `<td data-label='รายการ' class='text-left'>${requestDocOR[n].doc_title}</td>`;
                                setDataorther +=
                                    "<td  class='link-table'><a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                    dataObj.demand_letter_data.demand_letter_id +
                                    " &quot;, &quot; " +
                                    requestDocOR[n].document_id +
                                    " &quot;, &quot; " +
                                    doc_type.form_type_doc_id +
                                    " &quot;)'></a></td>";
                                setDataorther += "</tr>";

                            }
                        }

                        jQuery("#listFile_or46_59").html(setDataorther);
                    } else {
                        $("#div_showFileOrther59 , .file_or_div").addClass("d-none");
                    }

                    var FileSell =
                        dataObj.demand_letter_data.demand_letter_doc_list.saleperson_doclist;
                    var setDatasell = "";
                    if (dataObj.demand_letter_data.is_saleperson == "Y") {
                        $(".div_head_show_sell").css("display", "block");
                        $(".div_head_show_file_sell").css("display", "block");
                        for (let s = 0; s < FileSell.length; s++) {
                            const requestDocList = FileSell[s].requestdoclist;
                            setDatasell += "<tr>";
                            setDatasell += `<td data-label='ลำดับ'>${s + 1}</td>`;
                            for (let o = 0; o < requestDocList.length; o++) {
                                if (o != index_array) {
                                    setDatasell += `<td data-label='ลำดับ'></td>`;
                                }
                                setDatasell += `<td data-label='รายการ' class='text-left'><span data_th='${requestDocList[o].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[o].document_name_en}'></span></td>`;
                                if (requestDocList[o].doclist.length > 0) {
                                    setDatasell +=
                                        "<td  class='link-table'><a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                        dataObj.demand_letter_data.employer_data.employer_id +
                                        " &quot;, &quot; " +
                                        requestDocList[o].doclist[0].document_id +
                                        " &quot;, &quot; " +
                                        requestDocList[o].document_type_id +
                                        " &quot;)'></a></td>";


                                } else {
                                    setDatasell += `<td  class='link-table' tyle='text-decoration: none;'>-</td>`;
                                }

                                setDatasell += "</tr>";
                            }
                        }
                        jQuery("#listuploadfile_file_sell").html(setDatasell);
                    } else {
                        $(".div_head_show_sell").css("display", "none");
                        $(".div_head_show_file_sell").css("display", "none");
                    }

                    var setdataEmpUp = "";
                    for (let y = 0; y < objdoclist.agency_y_doclist.length; y++) {
                        $(".div_head_show_empup").css("display", "block");
                        $(".div_head_show_file_empup").css("display", "block");
                        const requestDocList = objdoclist.agency_y_doclist[y].requestdoclist;
                        setdataEmpUp += "<tr>";
                        setdataEmpUp += `<td data-label='ลำดับ'>${y + 1}</td>`;
                        for (let k = 0; k < requestDocList.length; k++) {
                            if (k != index_array) {
                                setdataEmpUp += `<td data-label='ลำดับ'></td>`;
                            }
                            setdataEmpUp += `<td data-label='รายการ' class='text-left'><span data_th='${requestDocList[k].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[k].document_name_en}'></span></td>`;
                            if (requestDocList[k].doclist.length > 0) {
                                setdataEmpUp +=
                                    "<td  class='link-table'><a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                    dataObj.demand_letter_data.demand_letter_id +
                                    " &quot;, &quot; " +
                                    requestDocList[k].doclist[0].document_id +
                                    " &quot;, &quot; " +
                                    requestDocList[k].form_type_doc_id +
                                    " &quot;)'></a></td>";


                            } else {
                                setdataEmpUp += `<td  class='link-table' tyle='text-decoration: none;'>-</td>`;
                            }

                            setdataEmpUp += "</tr>";
                        }
                    }
                    jQuery("#setdataEmpUp").html(setdataEmpUp);
                    var setDataAgency = "";
                    for (let a = 0; a < objdoclist.agency_n_doclist.length; a++) {
                        const requestDocList = objdoclist.agency_n_doclist[a].requestdoclist;
                        setDataAgency += "<tr>";
                        setDataAgency += `<td data-label='ลำดับ'>${a + 1}</td>`;
                        for (let r = 0; r < requestDocList.length; r++) {
                            if (r != index_array) {
                                setDataAgency += `<td data-label='ลำดับ'></td>`;
                            }
                            setDataAgency += `<td data-label='รายการ' class='text-left'><span data_th='${requestDocList[r].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[r].document_name_en}'></span></td>`;
                            if (requestDocList[r].doclist.length > 0) {
                                setDataAgency +=
                                    "<td  class='link-table'><a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                    dataObj.demand_letter_data.agency_id +
                                    " &quot;, &quot; " +
                                    requestDocList[r].doclist[0].document_id +
                                    " &quot;, &quot; " +
                                    requestDocList[r].document_type_id +
                                    " &quot;)'></a></td>";

                            } else {
                                setDataAgency += `<td  class='link-table' tyle='text-decoration: none;'>-</td>`;
                            }

                            setDataAgency += "</tr>";
                        }
                    }

                    jQuery("#list_agency_detail").html(setDataAgency);

                    var setDatadocfile_list = "";
                    var docfile_list = dataObj.demand_letter_data.docfile_list;
                    var status = sessionStorage.getItem("group_status");
                    if (status == "APSS" || status == "SS" || status == "WCPOS") {
                        for (let c = 0; c < docfile_list.length; c++) {
                            const requestDocList = docfile_list[c].requestdoclist;

                            setDatadocfile_list += "<div class='col-md-12'>";
                            if (requestDocList.length >= 2) {
                                setDatadocfile_list += `<p class='text-upload'>${c + 1
                                    }${"."}<span class='lang_select_at_least_one_file'></span>`;
                                if (docfile_list[c].required == "Y") {
                                    setDatadocfile_list += "<span class='asterisk'>*</span>";
                                    setDatadocfile_list +=
                                        "<span id='err_file_name_" +
                                        docfile_list[c].group_form_type_doc_id +
                                        "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                                }
                            }

                            for (let j = 0; j < requestDocList.length; j++) {
                                if (requestDocList[j].doclist.length > 0) {
                                    setDatadocfile_list += '<div class="row">';
                                    if (requestDocList.length < 2) {
                                        setDatadocfile_list += "<div class='col-md-6'>";

                                        setDatadocfile_list += `<p class='text-upload'>${c + 1
                                            }${"."}<span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()
                                            }' data_en='${requestDocList[j].document_name_en}'></span>`;
                                    } else {
                                        setDatadocfile_list += '<div style="margin-left:15px;">';
                                        setDatadocfile_list += `<p class='text-upload'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en}'></span>`;
                                    }
                                    setDatadocfile_list += "</p>";
                                    setDatadocfile_list += "</p>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list +=
                                        "<div class='col-md-6' style='text-align:end;'>";
                                    setDatadocfile_list +=
                                        "<a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                        dataObj.demand_letter_data.demand_letter_id +
                                        " &quot;, &quot; " +
                                        requestDocList[j].doclist[0].document_id +
                                        " &quot;, &quot; " +
                                        requestDocList[j].form_type_doc_id +
                                        " &quot;)'></a>";

                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "<div class='col-md-12 mb-2 mt-2'>";
                                    setDatadocfile_list += "<hr>";
                                    setDatadocfile_list += "</div>";
                                }
                            }
                        }
                    } else {
                        for (let c = 0; c < docfile_list.length; c++) {
                            const requestDocList = docfile_list[c].requestdoclist;

                            setDatadocfile_list += "<div class='col-md-12'>";
                            if (requestDocList.length >= 2) {
                                setDatadocfile_list += `<p class='text-upload'>${c + 1
                                    }${"."}<span class='lang_select_at_least_one_file'></span>`;
                                if (docfile_list[c].required == "Y") {
                                    setDatadocfile_list += "<span class='asterisk'>*</span>";
                                    setDatadocfile_list +=
                                        "<span id='err_file_name_" +
                                        docfile_list[c].group_form_type_doc_id +
                                        "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                                }
                            }

                            for (let j = 0; j < requestDocList.length; j++) {
                                lengthfile.push(requestDocList[j].document_type_id);
                                if (status == "WA") {
                                    if (requestDocList.length < 2) {
                                        setDatadocfile_list += "<div>";
                                        setDatadocfile_list += `<p class='text-upload'>${c + 1
                                            }${"."}<span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()
                                            }' data_en='${requestDocList[j].document_name_en}'></span>`;
                                        if (docfile_list[c].required == "Y") {
                                            setDatadocfile_list += "<span class='asterisk'>*</span>";
                                            setDatadocfile_list +=
                                                "<span id='err_file_name_" +
                                                requestDocList[j].document_type_id +
                                                "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                                        }
                                    } else {
                                        setDatadocfile_list += '<div style="margin-left:15px;">';

                                        setDatadocfile_list += `<p class='text-upload'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en}'></span>`;
                                    }

                                    setDatadocfile_list += "</p>";
                                    setDatadocfile_list += "</p>";
                                    setDatadocfile_list += "<div class='row'>";
                                    setDatadocfile_list += "<div class='col-lg-12'>";
                                    setDatadocfile_list +=
                                        "<p class='text-upload-subtext lang_support_pdf'>รองรับไฟล์ .PDF .PNG เเละ .JPG ขนาดไม่เกิน 25 mb</p>";
                                    setDatadocfile_list +=
                                        "<div class='row' style='margin-left:0px;'>";
                                    setDatadocfile_list +=
                                        "<div class='upload' style='padding-right:14px;'>";
                                    if (requestDocList.length >= 2) {
                                        setDatadocfile_list +=
                                            "<input type='file' name='file_name_group_doc' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='selectFile46_59Upload(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;)' hidden/>";
                                    } else {
                                        setDatadocfile_list +=
                                            "<input type='file' name='file_name_doc' id='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' onchange='selectFile46_59Upload(event,&quot;" +
                                            requestDocList[j].document_type_id +
                                            "&quot;)' hidden/>";
                                    }

                                    if (requestDocList[j].doclist.length > 0) {
                                        setDatadocfile_list +=
                                            "<label class='lang_lbl_selectfile' for='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' style='cursor:pointer;'>เลือกไฟล์</label>";
                                        setDatadocfile_list +=
                                            '<input type="hidden" id="validate_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="Y">';
                                    } else {
                                        setDatadocfile_list +=
                                            "<label class='lang_lbl_selectfile' for='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' style='cursor:pointer;'>เลือกไฟล์</label>";
                                        setDatadocfile_list +=
                                            "<span class='lang_nofile_chosen' id='divselect_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "'>No file chosen</span>";
                                        setDatadocfile_list +=
                                            '<input type="hidden" id="validate_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="N">';
                                    }
                                    setDatadocfile_list += "</div>";
                                    if (requestDocList[j].doclist.length > 0) {
                                        setDatadocfile_list +=
                                            "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "'>";
                                    } else {
                                        setDatadocfile_list +=
                                            "<div class='block-doc justify-content-between d-none' id='divshow_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' >";
                                    }

                                    setDatadocfile_list += "<div class='d-flex'>";
                                    setDatadocfile_list += "<div>";
                                    setDatadocfile_list +=
                                        "<img src='/assets/images/document.svg'>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "<div class='filename'>";
                                    if (requestDocList[j].doclist.length > 0) {

                                        setDatadocfile_list +=
                                            "<a id='taga_" +
                                            requestDocList[j].document_type_id +
                                            "' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                            dataObj.demand_letter_data.demand_letter_id +
                                            " &quot;,&quot; " +
                                            requestDocList[j].doclist[0].document_id +
                                            " &quot;, &quot; " +
                                            requestDocList[j].form_type_doc_id +
                                            " &quot;)'><span class='name-file namedoc' id='show_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "'>" +
                                            requestDocList[j].doclist[0].doc_title +
                                            "</span></a>";
                                    } else {
                                        setDatadocfile_list +=
                                            "<p class='name-file namedoc' id='show_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "'></p>";
                                    }

                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "<div>";

                                    setDatadocfile_list +=
                                        '<input type="hidden" id="datamore_' +
                                        requestDocList[j].document_type_id +
                                        '" value="' +
                                        requestDocList[j].document_type_id +
                                        '">';
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list += "</div>";
                                    setDatadocfile_list +=
                                        '<input type="hidden" id="requiredvalue_file_name_' +
                                        requestDocList[j].document_type_id +
                                        '" value="' +
                                        docfile_list[c].required +
                                        '">';
                                    setDatadocfile_list +=
                                        '<input type="hidden" id="required_file_name_' +
                                        requestDocList[j].document_type_id +
                                        '" value="' +
                                        docfile_list[c].group_form_type_doc_id +
                                        '">';
                                    setDatadocfile_list +=
                                        '<input type="hidden" id="form_type_doc_id_file_name_' +
                                        requestDocList[j].document_type_id +
                                        '" value="' +
                                        requestDocList[j].form_type_doc_id +
                                        '">';
                                } else {
                                    if (requestDocList[j].doclist.length > 0) {
                                        setDatadocfile_list += '<div class="row">';
                                        if (requestDocList.length < 2) {
                                            setDatadocfile_list += "<div class='col-md-6'>";

                                            setDatadocfile_list += `<p class='text-upload'>${c + 1
                                                }${"."}<span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()
                                                }' data_en='${requestDocList[j].document_name_en}'></span>`;
                                        } else {
                                            setDatadocfile_list += '<div style="margin-left:15px;">';
                                            setDatadocfile_list += `<p class='text-upload'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en}'></span>`;
                                        }
                                        setDatadocfile_list += "</p>";
                                        setDatadocfile_list += "</p>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list +=
                                            "<div class='col-md-6' style='text-align:end;'>";
                                        setDatadocfile_list +=
                                            "<a href='javascript:void(0)'><img src='/assets/images/document.svg' class='iconbook' onclick='GetDocFilePathByDocumentIDMOU(&quot; " +
                                            dataObj.demand_letter_data.demand_letter_id +
                                            " &quot;, &quot; " +
                                            requestDocList[j].doclist[0].document_id +
                                            " &quot;, &quot; " +
                                            requestDocList[j].form_type_doc_id +
                                            " &quot;)'></a>";

                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "</div>";
                                    } else {
                                        if (requestDocList.length < 2) {
                                            setDatadocfile_list += "<div>";
                                            setDatadocfile_list += `<p class='text-upload'>${c + 1
                                                }${"."}<span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()
                                                }' data_en='${requestDocList[j].document_name_en}'></span>`;
                                            if (docfile_list[c].required == "Y") {
                                                setDatadocfile_list += "<span class='asterisk'>*</span>";
                                                setDatadocfile_list +=
                                                    "<span id='err_file_name_" +
                                                    requestDocList[j].document_type_id +
                                                    "' class='text-err lang_please_add_file'>&nbsp;กรุณาเพิ่มไฟล์</span>";
                                            }
                                        } else {
                                            setDatadocfile_list += '<div style="margin-left:15px;">';

                                            setDatadocfile_list += `<p class='text-upload'><span data_th='${requestDocList[j].document_name_th?.toSafeJsonHtml()}' data_en='${requestDocList[j].document_name_en}'></span>`;
                                        }

                                        setDatadocfile_list += "</p>";
                                        setDatadocfile_list += "</p>";
                                        setDatadocfile_list += "<div class='row'>";
                                        setDatadocfile_list += "<div class='col-lg-12'>";
                                        setDatadocfile_list +=
                                            "<p class='text-upload-subtext lang_support_pdf'>รองรับไฟล์ .PDF .PNG เเละ .JPG ขนาดไม่เกิน 25 mb</p>";
                                        setDatadocfile_list +=
                                            "<div class='row' style='margin-left:0px;'>";
                                        setDatadocfile_list +=
                                            "<div class='upload' style='padding-right:14px;'>";
                                        if (requestDocList.length >= 2) {
                                            setDatadocfile_list +=
                                                "<input type='file' name='file_name_group_doc' id='file_name_" +
                                                requestDocList[j].document_type_id +
                                                "' onchange='selectFile46_59Upload(event,&quot;" +
                                                requestDocList[j].document_type_id +
                                                "&quot;)' hidden/>";
                                        } else {
                                            setDatadocfile_list +=
                                                "<input type='file' name='file_name_doc' id='file_name_" +
                                                requestDocList[j].document_type_id +
                                                "' onchange='selectFile46_59Upload(event,&quot;" +
                                                requestDocList[j].document_type_id +
                                                "&quot;)' hidden/>";
                                        }
                                        setDatadocfile_list +=
                                            "<label class='lang_lbl_selectfile' for='file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' style='cursor:pointer;'>เลือกไฟล์</label>";
                                        setDatadocfile_list +=
                                            "<span class='lang_nofile_chosen' id='divselect_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "'>No file chosen</span>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list +=
                                            "<div class='block-doc justify-content-between' id='divshow_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "' style='display:none;'>";
                                        setDatadocfile_list += "<div class='d-flex'>";
                                        setDatadocfile_list += "<div>";
                                        setDatadocfile_list +=
                                            "<img src='/assets/images/document.svg'>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "<div class='filename'>";
                                        setDatadocfile_list +=
                                            "<p class='name-file namedoc' id='show_file_name_" +
                                            requestDocList[j].document_type_id +
                                            "'></p>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "<div>";

                                        setDatadocfile_list +=
                                            '<input type="hidden" id="datamore_' +
                                            requestDocList[j].document_type_id +
                                            '" value="' +
                                            requestDocList[j].document_type_id +
                                            '">';
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list += "</div>";
                                        setDatadocfile_list +=
                                            '<input type="hidden" id="requiredvalue_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="' +
                                            docfile_list[c].required +
                                            '">';
                                        setDatadocfile_list +=
                                            '<input type="hidden" id="required_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="' +
                                            docfile_list[c].group_form_type_doc_id +
                                            '">';
                                        setDatadocfile_list +=
                                            '<input type="hidden" id="form_type_doc_id_file_name_' +
                                            requestDocList[j].document_type_id +
                                            '" value="' +
                                            requestDocList[j].form_type_doc_id +
                                            '">';
                                    }
                                }
                            }
                            setDatadocfile_list += "</div>";
                            setDatadocfile_list += "<div class='col-md-12 mb-2 mt-2'>";
                            setDatadocfile_list += "<hr>";
                            setDatadocfile_list += "</div>";
                        }
                    }

                    jQuery("#docfile_list").html(setDatadocfile_list);

                    if (dataObj.paymentlist.length > 0) {
                        $("#form_id_success_pay").text(getParameterByName("group_id"));
                        $("#createdtimestamp_success_pay").text(dataObj.created_timestamp);

                        for (var b = 0; b < dataObj.paymentlist.length; b++) {
                            if (dataObj.paymentlist[b].payment_type_id == "REQUESTANDFEE") {
                                $("#ref_id_ss_1").text(dataObj.paymentlist[b].ref_id);
                                $("#ref_id_ss_1_1").text(dataObj.paymentlist[b].ref_id);
                                $("#price_ss_1").text(
                                    dataObj.paymentlist[b].payment_price + " " + "บาท"
                                );
                            } else if (dataObj.paymentlist[b].payment_type_id == "COLLATERAL") {
                                $("#ref_id_ss_2").text(dataObj.paymentlist[b].ref_id);
                                $("#ref_id_ss_1_2").text(dataObj.paymentlist[b].ref_id);
                                $("#price_ss_2").text(
                                    dataObj.paymentlist[b].payment_price + " " + "บาท"
                                );
                            }
                        }
                    }


                    var arraystatus_showbtnadd = [
                        "WCOSNA2",
                        "WCOSAS2",
                        "WCOCNA2",
                        "WCOCAS2",
                        "WCOCOC2",
                        "AP",
                        "CCOS2",
                    ];

                    if (arraystatus_showbtnadd.includes(getstatus)) {
                        $("#btnaddEmp").css("display", "none");
                    } else if (getstatus == "WA") {
                        $("#btnpayment").removeClass("d-none");
                        var filecheck = "Y";


                        $("#head_step_2").show();

                        if (dataObj.formlist.length > 0) {
                            $("#Alien_Table_None").css("display", "none");
                            $("#Alien_Table").show();
                        } else {
                            $("#Alien_Table_None").css("display", "block");
                            $("#Alien_Table").hide();
                        }
                    } else if (
                        getstatus == "NCOS2" ||
                        getstatus == "NCOC2" ||
                        getstatus == "NCOCR2" ||
                        getstatus == "NCOCRD2"
                    ) {
                        $("#head_step_2").css("display", "flex");
                        $("#btnpayment").css("display", "none");
                    } else if (getstatus == "APSS" || getstatus == "NCOS3") {
                        $("#Alien_Table_None").css("display", "none");
                        $("#divpayment_2").css("display", "none");
                    } else if (getstatus == "WCPOS") {
                        $("#Alien_Table").css("display", "none");
                    }

                    var status_reject1 = ["NCOS1", "NCOC1", "NCOCR1", "NCOCRD1", ""];
                    if (
                        (dataObj.request_by_user_id == core_datas.user_id &&
                            dataObj.request_by_user_id_proxy == "") ||
                        dataObj.request_by_user_id_proxy == core_datas.user_id
                    ) {
                        $(".show_for_proxy").css("display", "none");

                    } else {

                        if (status_reject1.includes(getstatus)) {
                            $("#text_button").addClass("d-none");
                        }

                        if (getstatus == "WP3") {
                            $("#text_button").addClass("d-none");
                            $("#div_no_payment_no_proxy").removeClass("d-none");
                            $("#paymentlist41").addClass("d-none");
                        }

                        if (getstatus == "AP") {
                            $("#div_no_appointment_no_proxy").removeClass("d-none");
                            $("#div_appointment_true").addClass("d-none");
                            $("#desc_ap").attr("data_th", objData.tracking_desc_th);
                            $("#desc_ap").attr("data_en", objData.tracking_desc_en);
                        }
                        $("#desc_wp").attr("data_th", objData.tracking_desc_th);
                        $("#desc_wp").attr("data_en", objData.tracking_desc_en);

                        $(".cancelRequestForm").addClass("d-none");
                        $("#btn_SendDoc").addClass("d-none");
                        $("#divpayment").hide();
                    }

                    $("#loading").delay().fadeOut("");
                }
            } else {
                $("#loading").delay().fadeOut("");
                alert(data.msg);
            }
        }, "json");
    await Promise.all([changeDatalang()]);
}

function clearData() {
    $("#form_add_alien")[0].reset();
    $("#id").val("");
    $("#EmployeeModal").modal("hide");
    $("#PASSPORT").attr("checked", true);
    GetDataRequestFormEtrackingMOU_ByGroupID();
}
function SumbitDocFileDemandLetterOrNameList(id_file) {
    $("#loading").css("display", "block");
    var formData = new FormData();
    var dataarray = [];
    var arraytype = [];
    var tmpfiletitle = [];

    jQuery("[name='file_name']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file = fileUpload[0].files[0];
        if (file != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#form_type_doc_id_" + this.id).val();


            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file);
        }
    });

    formData.append("demand_letter_id", localStorage.getItem("demand_id"));
    if (core_datas.default_profile.user_type_id == "5") {
        formData.append("user_id", core_datas.user_id);
    } else {
        formData.append("user_id", core_datas.current_profile.user_id);
    }

    formData.append("document_title", tmpfiletitle.toString());
    formData.append("form_type_doc_id", arraytype.toString());
    formData.append("relate_account_id", core_datas.current_profile.profile_id);

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
                jQuery("[name='file_name']").val("");
                jQuery("#show_" + id_file).text(file.name);
                $("#divshow_" + id_file).css("display", "flex");
                $("#divselect_" + id_file).css("display", "none");
                $("#uploadfile_" + id_file).val(id_file);
                $("#loading").delay().fadeOut("");
            } else {
                $("#loading").delay().fadeOut("");
                $("#success").removeClass("active");
                Swal.fire({
                    icon: "error",
                    title: "เพิ่มไฟล์ไม่สำเร็จ",
                    showCancelButton: false,
                    confirmButtonText: "ตกลง",
                    reverseButtons: true,
                }).then((result) => {

                });

            }
        },
    });
    changeDatalang();
}
$("#AddNewNameList").on("click", function () {
    $("#loading").css("display", "block");


    let insurance_condition = $("#insurance_condition").val();
    let fileup72 = $("#file_name_72")[0]?.files[0];

    var chkData = true;

    if (insurance_condition == "NOT_PASS" && !fileup72) {
        $("#loading").delay().fadeOut("")
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
            html: `<span>เนื่องจากผลการตรวจสุขภาพหรือข้อมูลประกันสุขภาพของท่านไม่เป็นไปตามเกณฑ์ที่กำหนด จึงไม่สามารถยื่นคำขอได้</span>`,
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            allowOutsideClick: false,
            allowEscapeKey: false
        }).then((result) => {
            if (result.isConfirmed) {
                $("#EmployeeModal").modal('hide');
                $("#AddNewNameList").prop('disabled', false)
            }
        });
        return;
    }
    if (insurance_condition == "NO_DATA" && !fileup72) {
        $("#loading").delay().fadeOut("")
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
            html: `<span>เนื่องจากไม่พบข้อมูลและเอกสารของผลการตรวจสุขภาพหรือประกันสุขภาพในระบบ</span>`,
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            allowOutsideClick: false,
            allowEscapeKey: false
        }).then((result) => {
            if (result.isConfirmed) {

                $("#AddNewNameList").prop('disabled', false)

            }
        });
        return;
    }
    if ($("#form_add_alien").valid()) {
        var chkData = true;
        jQuery("[name='file_name']").each(function (index) {
            var fileUpload = jQuery("#" + this.id);
            var ckrequired = jQuery("#requiredvalue_" + this.id).val();
            var file = fileUpload[0].files[0];
            var filebase64 = "";
            if ($("#is_accept_work_N").is(":checked")) {
                $("#err_" + this.id + "").css("display", "none");
            } else {
                if (ckrequired == "Y") {
                    if (this.id == "file_name_72") {
                        if (checkhealth == "H" && insurance_condition !== "PASS") {
                            chkData = false;
                            $("#err_" + this.id + "").css("display", "inline");
                            jQuery("#file_link_" + this.id).text("-");
                            $(".link-table a").css("text-decoration", "none");
                        }
                    } else {
                        if (file != undefined) {
                            $("#err_" + this.id + "").css("display", "none");
                        } else {
                            chkData = false;
                            $("#err_" + this.id + "").css("display", "inline");
                            jQuery("#file_link_" + this.id).text("-");
                            $(".link-table a").css("text-decoration", "none");
                        }
                    }
                }
            }
        });
        var chkDataGroup = true;
        var arrayfileUploadGroup = [];
        var arrayFileRequest = [];
        $("[name='file_name_group']").each(function (index) {
            var fileUploadGroup = jQuery("#" + this.id);
            var FileRequest = $("#required_" + this.id).val();
            var ckrequired = jQuery("#requiredvalue_" + this.id).val();
            var fileGroup = fileUploadGroup[0].files[0];
            if (fileGroup != undefined) {
                arrayfileUploadGroup.push(fileGroup);
            }

            arrayFileRequest.push(FileRequest);

            if (ckrequired == "Y") {
                if (arrayfileUploadGroup.length == 0) {
                    if ($("#is_accept_work_N").is(":checked")) {
                        chkDataGroup = true;
                    } else {
                        chkDataGroup = false;
                        for (var v = 0; v < arrayFileRequest.length; v++) {
                            $("#err_file_name_" + arrayFileRequest[v]).css(
                                "display",
                                "inline"
                            );
                        }
                    }
                } else {
                    chkDataGroup = true;
                    for (var v = 0; v < arrayFileRequest.length; v++) {
                        $("#err_file_name_" + arrayFileRequest[v]).css("display", "none");
                    }
                }
            }
        });

        if (chkData == true && chkDataGroup == true) {
            $("#AddNewNameList").attr("disabled", true);
            AddNewDataNameList41();
        } else {
            $("#loading").delay().fadeOut("");
        }
    } else {
        $("#loading").delay().fadeOut("");
        const inputElements = document.querySelectorAll("input.error");
        const focusedElement = document.getElementById(
            "" + inputElements[0].id + ""
        );


        if (focusedElement) {
            focusedElement.focus();
            focusedElement.scrollIntoView({
                behavior: "smooth",
                block: "center",
                inline: "center",
            });
        }

        var chkData = true;
        jQuery("[name='file_name']").each(function (index) {
            var fileUpload = jQuery("#" + this.id);
            var ckrequired = jQuery("#requiredvalue_" + this.id).val();
            var file = fileUpload[0].files[0];
            var filebase64 = "";
            if ($("#is_accept_work_N").is(":checked")) {
                $("#err_" + this.id + "").css("display", "none");
            } else {
                if (ckrequired == "Y") {
                    if (this.id == "file_name_72") {
                        if (checkhealth == "H" && insurance_condition !== "PASS") {
                            chkData = false;
                            $("#err_" + this.id + "").css("display", "inline");
                            jQuery("#file_link_" + this.id).text("-");
                            $(".link-table a").css("text-decoration", "none");
                        }
                    } else {
                        if (file != undefined) {
                            $("#err_" + this.id + "").css("display", "none");
                        } else {
                            chkData = false;
                            $("#err_" + this.id + "").css("display", "inline");
                            jQuery("#file_link_" + this.id).text("-");
                            $(".link-table a").css("text-decoration", "none");
                        }
                    }
                }
            }
        });
        var chkDataGroup = true;
        var arrayfileUploadGroup = [];
        var arrayFileRequest = [];
        $("[name='file_name_group']").each(function (index) {
            var fileUploadGroup = jQuery("#" + this.id);
            var FileRequest = $("#requiredvalue_" + this.id).val();
            var hidden_form_type = $("#required_" + this.id).val();
            var fileGroup = fileUploadGroup[0].files[0];
            var fileId = this.id;
            var aTag = document.createElement("a");
            aTag.id = "file_link_" + fileId;
            aTag.innerHTML =
                "<img src='/assets/images/document.svg' class='iconbook'>";
            $("#file_link_" + fileId).empty();
            if (fileGroup != undefined) {
                arrayfileUploadGroup.push(fileGroup);
                arrayFileSelect.push(FileRequest);
                var objectUrl = URL.createObjectURL(fileGroup);
                var reader = new FileReader();
                reader.readAsDataURL(fileGroup);
                reader.onload = function () {
                    let base64 = reader.result;

                    if (app == "Y" || app == "y") {
                        document.getElementById("file_link_" + fileId).innerHTML =
                            "<img src='/assets/images/document.svg' style='background: white; border-radius: 6px; border-color: #DDBFFD; padding: 6px; border: 1px solid;' onclick='sendmessage(&quot;" +
                            base64 +
                            "&quot;)'>";
                    } else {
                        $(aTag).attr({
                            href: objectUrl,
                            target: "_blank",
                        });
                        document.getElementById("file_link_" + fileId).appendChild(aTag);
                    }
                };
            } else {
                jQuery("#file_link_" + this.id).text("-");
            }
            if (FileRequest == "Y") {
                if (!arrayFileRequest.includes(hidden_form_type)) {
                    arrayFileRequest.push(hidden_form_type);
                }
            }
        });
        arrayFileRequest.forEach(function (fileName) {
            var errorElement = $("#err_file_name_" + fileName);
            if (filevalidate.length == 0) {
                allFieldsValid = false;
                $("#err_file_name_" + fileName).css("display", "inline");
            } else {
                if (filevalidate.includes(fileName)) {
                    $("#err_file_name_" + fileName).css("display", "none");
                    allFieldsValidtrue = true;
                } else {

                    allFieldsValid = false;
                    $("#err_file_name_" + fileName).css("display", "inline");
                }
            }
        });
    }
});
function AddNewDataNameList41() {
    $("#loading").css("display", "block");

    var formData = new FormData();
    var dataarray = [];
    var arraytype = [];
    var tmpfiletitle = [];

    jQuery("[name='file_name']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file_up = fileUpload[0].files[0];
        if (file_up != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file_up.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#hidden_form_type_" + this.id).val();
            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file_up);
        }
    });
    jQuery("[name='file_name_group']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file_group = fileUpload[0].files[0];
        if (file_group != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file_group.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#hidden_form_type_" + this.id).val();
            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file_group);
        }
    });
    var array_job_desc_id = [];
    var array_job_other = [];
    var array_accept = [];
    var stay_permission = [];
    var stay_permission_name = [];
    var dataarray = "";
    $('input[name="AcceptWork"]:checked').each(function () {
        var ID_checked = this.id;
        if (ID_checked == "work_info_0") {
            array_job_other.push($("#accept_work_other").val());
        } else {
            array_job_other.push("-");
        }
        array_job_desc_id.push(this.value);
    });

    $('input[name="is_accept_work"]:checked').each(function () {
        array_accept.push(this.value);
    });
    $('input[name="stay_permission_type46_59"]:checked').each(function () {
        stay_permission.push(this.id);
        stay_permission_name.push(this.value);
    });

    formData.append("document_title", tmpfiletitle.toString());
    formData.append("form_type_doc_id", arraytype.toString());
    if (core_datas.default_profile.user_type_id == "5") {
        formData.append("user_id", core_datas.user_id);
    } else {
        formData.append("user_id", core_datas.current_profile.user_id);
    }

    formData.append("group_id", getParameterByName("group_id"));
    formData.append("pf_id", $("#prefixNameEnDropdown :selected").val());
    formData.append("alien_first_name_th", $("#name_thcheckreg").val());
    formData.append("alien_middle_name_th", $("#middle_name_thcheckreg").val());
    formData.append("alien_last_name_th", $("#lastname_thcheckreg").val());
    formData.append("alien_first_name_en", $("#name_encheckreg").val());
    formData.append("alien_middle_name_en", $("#middle_name_encheckreg").val());
    formData.append("alien_last_name_en", $("#lastname_encheckreg").val());
    formData.append("alien_nationality", $("#nationality_al").val());
    if ($("#birthdate_al").val() != "") {
        var resultDate = ConvertDate($("#birthdate_al").val());
        formData.append("alien_date_of_birth", resultDate);
    }

    formData.append("alien_address_telephone", $("#tel_al").val());
    formData.append("alien_address_email", $("#email_al").val());
    formData.append("alien_address_lineid", $("#soi_work_63").val());
    formData.append("stay_permis_type_id", stay_permission);
    formData.append("stay_permis_name", stay_permission_name);
    formData.append("stay_permis_no", $("#placepassport_al").val());
    formData.append("stay_permis_issue_at", $("#issued_al").val());
    formData.append("stay_permis_country_id", $("#country_al :selected").val());
    if ($("#dateofexpiry").val() != "") {
        var dateofexpiry = ConvertDate($("#dateofexpiry").val());
        formData.append("stay_permis_issue_dt", dateofexpiry);
    }
    if ($("#enddateofexpiry").val() != "") {
        var enddateofexpiry = ConvertDate($("#enddateofexpiry").val());
        formData.append("stay_permis_expire_dt", enddateofexpiry);
    }

    formData.append("bus_type_id", $("#bus_type_id_value").val());
    formData.append("permit_cate_id", $("#permit_cate_id :selected").val());
    formData.append("job_desc_01", $("#job_description").val());
    formData.append("job_desc_02", "");
    formData.append("job_desc_03", "");
    formData.append("job_desc_04", "");
    formData.append("job_desc_05", "");
    formData.append("is_accept_work", array_accept);
    formData.append("accept_work_id", array_job_desc_id);
    formData.append("accept_work_other", array_job_other);
    formData.append("relate_account_id", core_datas.current_profile.profile_id);

    $.ajax({
        method: "post",
        processData: false,
        contentType: false,
        cache: false,
        data: formData,
        url: url_AddNewDataNameList41_4_59,
        success: function (response) {
            var dataresponse = JSON.parse(response);
            if (dataresponse.status == "success") {
                $("#loading").delay().fadeOut("");
                $("#EmployeeModal").modal("hide");
                GetDataRequestFormNameListByDemandLetterID("");
            } else {
                $("#loading").delay().fadeOut("");
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
                        imageUrl: important_data.warningIconUrl,
                        title:
                            '<span data_th="ไม่สามารถบันทึกได้" data_en="Cannot be saved"></span>',
                        html:
                            '<span data_th="' +
                            dataresponse.msg +
                            '" data_en="' +
                            dataresponse.msgen +
                            '"></span>',
                        confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                        reverseButtons: true,
                    }).then((result) => {
                        $("#AddNewNameList").attr("disabled", false);
                    });
                    changeDatalang();
                }


            }
        },
        error: function (xhr) {
            $("#loading").delay().fadeOut("");

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
    changeDatalang();
}

function handleSave41() {
    let checkinsurance = $("#insurance_file_name").val();
    let insurance_condition = $("#insurance_condition").val();
    let hidden_check_file = $("#hidden_check_file_name_72").val();
    let doc_new_lastest = $("#doc_new_lastest").val();
    let chkis_upload = $("#chkis_upload").val();

    if (checkinsurance == "N") {
        $("#loading").delay().fadeOut("")
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
            html: `<span>เนื่องจากผลการตรวจสุขภาพของท่านไม่ผ่านเกณฑ์ที่กำหนด
กรุณาย้อนกลับไปแนบเอกสารการตรวจสุขภาพฉบับใหม่อีกครั้ง</span>`,
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            allowOutsideClick: false,
            allowEscapeKey: false
        }).then((result) => {
            if (result.isConfirmed) {
                $("#EmployeeModal").modal('hide');

            }
        });
        return;
    }
    if (insurance_condition == "NOT_PASS") {
        $("#loading").delay().fadeOut("")
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
            html: `<span>เนื่องจากระยะเวลาคุ้มครองของประกันสุขภาพนั้นไม่ถึง 6 เดือน กรุณาย้อนกลับไปแนบเอกสารประกันสุขภาพฉบับใหม่อีกครั้ง</span>`,
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            allowOutsideClick: false,
            allowEscapeKey: false
        }).then((result) => {
            if (result.isConfirmed) {
                $("#EmployeeModal").modal('hide');

            }
        });
        return;
    }
    if (insurance_condition == "NO_DATA") {
        $("#loading").delay().fadeOut("")
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
            html: `<span>เนื่องจากไม่พบข้อมูลและเอกสารของ ผลการตรวจสุขภาพหรือประกันสุขภาพในระบบกรุณาย้อนกลับไปดำเนินการแนบเอกสารประกันสุขภาพฉบับใหม่ให้ครบถ้วนอีกครั้ง</span>`,
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            allowOutsideClick: false,
            allowEscapeKey: false
        }).then((result) => {
            if (result.isConfirmed) {


            }
        });
        return;
    }
    if (hidden_check_file == "UC" && doc_new_lastest == "N" && chkis_upload == "N") {
        $("#loading").delay().fadeOut("")
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
            html: `<span>เนื่องจากไม่พบข้อมูลและเอกสารของ ผลการตรวจสุขภาพหรือประกันสุขภาพในระบบกรุณาย้อนกลับไปดำเนินการแนบเอกสารประกันสุขภาพฉบับใหม่ให้ครบถ้วนอีกครั้ง</span>`,
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            allowOutsideClick: false,
            allowEscapeKey: false
        }).then((result) => {
            if (result.isConfirmed) {
                $("#EmployeeModal").modal('hide');

            }
        });
        return;
    }

    if (
        statuscheck == "NCOS2" ||
        statuscheck == "NCOC2" ||
        statuscheck == "NCOCR2"
    ) {
        if (validateFilesUC()) {
            UpdateDataNameList46_59();
        }
    } else {
        if (validateFiles()) {
            if (checkage) {
                UpdateDataNameList46_59();
            } else {
                Swal.fire({
                    imageUrl: important_data.warningIconUrl,
                    title:
                        '<span data_th="ไม่สามารถบันทึกได้" data_en="Cannot be saved"></span>',
                    html:
                        '<span data_th="เนื่องจากอายุต้องอยู่ระหว่าง ' +
                        important_data.alienAge.minAge.toString() +
                        " ถึง " +
                        important_data.alienAge.maxAge.toString() +
                        '" data_en="Because the age must be between ' +
                        important_data.alienAge.minAge.toString() +
                        " to " +
                        important_data.alienAge.maxAge.toString() +
                        '"></span>',
                    showCancelButton: false,
                    confirmButtonText: '<span data_th="ตกลง" data_en="OK"></span>',
                    reverseButtons: true,
                }).then((result) => { });
                changeDatalang();
            }
        }
    }
}
function checkfile() {
    var app = getParameterByName("App");
    var chkData = true;
    jQuery("[name='file_name']").each(function (index) {
        var app = getParameterByName("App");
        var fileUpload = jQuery("#" + this.id);
        var ckrequired = jQuery("#requiredvalue_" + this.id).val();
        var ckhave_file_current = jQuery("#have_file_current_" + this.id).val();
        var hidden_check_file_name = $("#hidden_check_" + this.id).val();
        var doc_new_lastest = $("#doc_new_lastest_" + this.id).val();
        var hidden_select = $("#hidden_select_" + this.id).val();
        var fileUp = fileUpload[0].files[0];
        var chIDFile = fileUpload[0].id;
        var aTag = document.createElement("a");
        aTag.id = "file_link_" + chIDFile;
        aTag.innerHTML = "<img src='/assets/images/document.svg' class='iconbook'>";
        $("#file_link_" + chIDFile).empty();


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
                            } else {
                                chkData = false;
                                $("#err_" + chIDFile + "").css("display", "inline");
                            }
                        }
                    } else {
                    }
                }
            }
        }
    });

    var arrayFileRequest = [];
    var chkDataGroupEdit = true;
    var chkDataGroupEdit_Y = true;
    var group_false = [];
    var group_true = [];
    var group_true_N = [];
    jQuery("[name='file_name_group']").each(function (index) {
        var app = getParameterByName("App");
        var fileUpload = jQuery("#" + this.id);
        var ckrequired = jQuery("#requiredvalue_" + this.id).val();
        var ckhave_file_current = jQuery("#have_file_current_" + this.id).val();
        var group_required = $("#required_" + this.id).val();
        var hidden_check_file_name = $("#hidden_check_" + this.id).val();
        var doc_new_lastest = $("#doc_new_lastest_" + this.id).val();
        var hidden_select = $("#hidden_select_" + this.id).val();
        var fileUp = fileUpload[0].files[0];
        var chIDFile = fileUpload[0].id;

        $("#file_link_" + chIDFile).empty();

        if (!filevalidate.includes(group_required)) {
            filevalidate.push(group_required);
        }


        if (ckrequired == "N") {

            if (ckhave_file_current != "N") {
                if (hidden_check_file_name == "UC") {
                    if (fileUp != undefined) {
                        chkDataGroupEdit = true;
                        KeepFile(group_required);
                    } else {
                        if (doc_new_lastest == "true") {
                            KeepFile(group_required);
                        } else {
                            jQuery("#file_link_" + chIDFile).text("-");
                            $(".link-table a").css("text-decoration", "none");
                            if (hidden_select == "0") {
                                KeepFile(group_required);
                            } else {
                                if (hidden_check_file_name == "UC") {
                                    chkDataGroupEdit = false;
                                    group_false.push(group_required);
                                    Removefiledupilcate(group_required);
                                } else if (chkDataGroupEdit) {
                                    chkDataGroupEdit = true;
                                } else {
                                    chkDataGroupEdit = false;
                                    group_false.push(group_required);
                                    Removefiledupilcate(group_required);
                                }
                            }
                        }
                    }
                } else {
                    KeepFile(group_required);
                }
            } else {
                if (hidden_check_file_name == "UC") {
                    if (fileUp != undefined) {
                        KeepFile(group_required);

                        if (!group_true_N.includes(group_required)) {
                            group_true_N.push(group_required);
                        }
                        chkDataGroupEdit = true;
                    } else {
                        if (doc_new_lastest == "true") {
                        } else {
                            jQuery("#file_link_" + chIDFile).text("-");
                            $(".link-table a").css("text-decoration", "none");

                            if (hidden_select == "0") {
                                KeepFile(group_required);
                            } else {
                                KeepFile(group_required);
                                chkDataGroupEdit = true;
                                group_false = group_false.filter(function (value) {
                                    return !group_true_N.includes(value);
                                });
                            }
                        }
                    }
                } else {
                    if (fileUp != undefined) {
                        KeepFile(group_required);
                    } else {
                        if (doc_new_lastest == "true") {
                        } else {
                            if (
                                ckrequired == "N" &&
                                hidden_check_file_name != "UC" &&
                                hidden_check_file_name !== undefined
                            ) {
                                if (!arrayfileUploadGroupEdit.includes(group_required)) {
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
                                chkDataGroupEdit = true;
                                group_false = group_false.filter(function (value) {
                                    return !group_true_N.includes(value);
                                });
                                KeepFile(group_required);
                            } else {
                                jQuery("#file_link_" + chIDFile).text("-");
                                $(".link-table a").css("text-decoration", "none");
                                Removefiledupilcate(group_required);
                            }
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
                    if (!group_true.includes(group_required)) {
                        group_true.push(group_required);
                    }
                    chkDataGroupEdit_Y = true;
                    KeepFile(group_required);
                } else {
                    if (
                        hidden_select != "0" &&
                        hidden_select != "Y" &&
                        hidden_select != undefined
                    ) {
                        chkDataGroupEdit_Y = true;
                        KeepFile(group_required);
                    } else {
                        if (group_true.length != 0) {
                            chkDataGroupEdit_Y = true;
                            group_false = group_false.filter(function (value) {
                                return !group_true.includes(value);
                            });
                        } else {
                            chkDataGroupEdit_Y = false;
                            group_false.push(group_required);
                        }
                    }
                    jQuery("#file_link_" + chIDFile).text("-");
                    $(".link-table a").css("text-decoration", "none");
                }
            } else {
                if (hidden_check_file_name == "UC") {
                    if (fileUp != undefined) {
                        KeepFile(group_required);
                        if (!group_true.includes(group_required)) {
                            group_true.push(group_required);
                        }
                        chkDataGroupEdit_Y = true;
                    } else {
                        if (doc_new_lastest == "true") {
                            KeepFile(group_required);
                        } else {
                            jQuery("#file_link_" + chIDFile).text("-");
                            $(".link-table a").css("text-decoration", "none");
                            if (hidden_select == "0") {
                                KeepFile(group_required);
                            } else {
                                chkDataGroupEdit_Y = false;
                                group_false.push(group_required);
                                Removefiledupilcate(group_required);
                            }
                        }
                    }
                } else {
                    KeepFile(group_required);
                    arrayFileRequest.push(ckrequired);
                    if (fileUp != undefined) {
                    } else {
                        if (group_false.length != 0) {
                            chkDataGroupEdit_Y = false;
                        } else {
                            chkDataGroupEdit_Y = true;
                        }
                        if (!group_true.includes(group_required)) {
                            group_true.push(group_required);
                        }

                        KeepFile(group_required);
                    }
                }
            }
        }
    });
    var allFieldsValid = true;
    if (chkDataGroupEdit == false || chkDataGroupEdit_Y == false) {
        filevalidate.forEach(function (fileName) {
            var errorElement = $("#err_file_name_" + fileName);
            if (group_false.includes(fileName)) {
                allFieldsValid = false;
                errorElement.css("display", "inline");
            } else {
                errorElement.css("display", "none");
            }
        });
    } else {
        filevalidate.forEach(function (fileName) {
            var errorElement = $("#err_file_name_" + fileName);
            if (arrayfileUploadGroupEdit.includes(fileName)) {
                errorElement.css("display", "none");
            } else {

                allFieldsValid = false;
                errorElement.css("display", "inline");
            }
        });
    }

    if (chkData && allFieldsValid && checkage) {
        UpdateDataNameList46_59();
    } else {
        var age_al = $("#age_al").val();
        if (parseInt(age_al) < 15 || parseInt(age_al) > 55) {
            Swal.fire({
                imageUrl: important_data.warningIconUrl,
                title:
                    '<span data_th="ไม่สามารถบันทึกได้" data_en="Cannot be saved"></span>',
                html:
                    '<span data_th="เนื่องจากอายุต้องอยู่ระหว่าง ' +
                    important_data.alienAge.minAge.toString() +
                    " ถึง " +
                    important_data.alienAge.maxAge.toString() +
                    '" data_en="Because the age must be between ' +
                    important_data.alienAge.minAge.toString() +
                    " to " +
                    important_data.alienAge.maxAge.toString() +
                    '"></span>',
                showCancelButton: false,
                confirmButtonText: '<span data_th="ตกลง" data_en="OK"></span>',
                reverseButtons: true,
            }).then((result) => { });
            changeDatalang();
        }
    }
}
function UpdateDataNameList46_59() {
    $("#loading").css("display", "block");

    var formData = new FormData();
    var dataarray = [];
    var arraytype = [];
    var tmpfiletitle = [];

    jQuery("[name='file_name']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file_up = fileUpload[0].files[0];
        if (file_up != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file_up.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#hidden_form_type_" + this.id).val();
            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file_up);
        }
    });
    jQuery("[name='file_name_group']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file_group = fileUpload[0].files[0];
        if (file_group != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file_group.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#hidden_form_type_" + this.id).val();
            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file_group);
        }
    });
    var array_job_desc_id = [];
    var array_job_other = [];
    var array_accept = [];
    var stay_permission = [];
    var stay_permission_name = [];
    var dataarray = "";
    $('input[name="AcceptWork"]:checked').each(function () {
        var ID_checked = this.id;
        if (ID_checked == "work_info_0") {
            array_job_other.push($("#accept_work_other").val());
        } else {
            array_job_other.push("-");
        }
        array_job_desc_id.push(this.value);
    });

    $('input[name="is_accept_work"]:checked').each(function () {
        array_accept.push(this.value);
    });
    $('input[name="stay_permission_type46_59"]:checked').each(function () {
        stay_permission.push(this.id);
        stay_permission_name.push(this.value);
    });

    formData.append("document_title", tmpfiletitle.toString());
    formData.append("form_type_doc_id", arraytype.toString());
    formData.append("document_id", currentfile.toString());

    formData.append("group_id", getParameterByName("group_id"));
    if (core_datas.default_profile.user_type_id == "5") {
        formData.append("user_id", core_datas.user_id);
    } else {
        formData.append("user_id", core_datas.current_profile.user_id);
    }

    formData.append("form_id", $("#form_id").val());
    formData.append("alien_id", $("#alien_id").val());
    formData.append("pf_id", $("#prefixNameEnDropdown :selected").val());
    formData.append("alien_first_name_th", $("#name_thcheckreg").val());
    formData.append("alien_middle_name_th", $("#middle_name_thcheckreg").val());
    formData.append("alien_last_name_th", $("#lastname_thcheckreg").val());
    formData.append("alien_first_name_en", $("#name_encheckreg").val());
    formData.append("alien_middle_name_en", $("#middle_name_encheckreg").val());
    formData.append("alien_last_name_en", $("#lastname_encheckreg").val());
    formData.append("alien_nationality", $("#nationality_al").val());
    if ($("#birthdate_al").val() != "") {
        var resultDate = ConvertDate($("#birthdate_al").val());
        formData.append("alien_date_of_birth", resultDate);
    }

    formData.append("alien_address_telephone", $("#tel_al").val());
    formData.append("alien_address_email", $("#email_al").val());
    formData.append("alien_address_lineid", $("#soi_work_63").val());
    formData.append("stay_permis_type_id", stay_permission);
    formData.append("stay_permis_name", stay_permission_name);
    formData.append("stay_permis_no", $("#placepassport_al").val());
    formData.append("stay_permis_issue_at", $("#issued_al").val());
    formData.append("stay_permis_country_id", $("#country_al :selected").val());

    if ($("#dateofexpiry").val() != "") {
        var dateofexpiry = ConvertDate($("#dateofexpiry").val());
        formData.append("stay_permis_issue_dt", dateofexpiry);
    }
    if ($("#enddateofexpiry").val() != "") {
        var enddateofexpiry = ConvertDate($("#enddateofexpiry").val());
        formData.append("stay_permis_expire_dt", enddateofexpiry);
    }
    formData.append("bus_type_id", $("#bus_type_id_value").val());
    formData.append("permit_cate_id", $("#permit_cate_id :selected").val());
    formData.append("job_desc_01", $("#job_description").val());
    formData.append("job_desc_02", "");
    formData.append("job_desc_03", "");
    formData.append("job_desc_04", "");
    formData.append("job_desc_05", "");
    formData.append("is_accept_work", array_accept);
    formData.append("accept_work_id", array_job_desc_id);
    formData.append("accept_work_other", array_job_other);
    formData.append("relate_account_id", core_datas.current_profile.profile_id);
    formData.append("update_doc_insurance", $("#doc_new_lastest").val());

    $.ajax({
        method: "post",
        processData: false,
        contentType: false,
        cache: false,
        data: formData,
        url: url_UpdateDataNameList41_4_59,
        success: function (response) {
            var dataresponse = JSON.parse(response);
            if (dataresponse.status == "success") {
                $("#EmployeeModal").modal("hide");


                location.reload();
            } else {
                $("#loading").delay().fadeOut("");
                $("#success").removeClass("active");
                $("#payment").addClass("active");
                $("#step5").css("display", "none");
                $("#step4").css("display", "block");
                $("#step4").css("opacity", "1");
                Swal.fire({
                    imageUrl: important_data.warningIconUrl,
                    title:
                        '<span data_th="ไม่สามารถบันทึกได้" data_en="Cannot be saved"></span>',
                    html:
                        '<span data_th="' +
                        dataresponse.msg +
                        '" data_en="' +
                        dataresponse.msgen +
                        '"></span>',
                    confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                    reverseButtons: true,
                }).then((result) => {
                    $("#success").removeClass("active");
                    $("#payment").addClass("active");
                    $("#step5").css("display", "none");
                    $("#step4").css("display", "block");
                    $("#step4").css("opacity", "1");
                });
                changeDatalang();
            }
        },
        error: function (xhr) {
            $("#loading").delay().fadeOut("");
            $("#success").removeClass("active");
            $("#payment").addClass("active");
            $("#step5").css("display", "none");
            $("#step4").css("display", "block");
            $("#step4").css("opacity", "1");
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
    changeDatalang();
}
function GetDataFormAppointmentByAppointmentDate(appointment_datetime) {
    if (important_data.is_mobile == "Y" || important_data.is_mobile == "y") {
        var userid = getParameterByName("user_id");
        var group_id = getParameterByName("group_id");
        objRequest = {
            user_id: userid,
            group_id: group_id,
            appointment_datetime: appointment_datetime,
        };
    } else {
        objRequest = {
            user_id: core_datas.current_profile.user_id,
            group_id: getParameterByName("group_id"),
            appointment_datetime: appointment_datetime,
        };
    }
    var request = jQuery.post(
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

                    data: objData,
                });
                $("#loading").delay().fadeOut("");

                changeDatalang();
            }
        },
        "json"
    );
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
        objRequest = {
            group_id: getParameterByName("group_id"),
            user_id: core_datas.current_profile.user_id,
        };
    }
    try {
        var request = await jQuery.post(
            url_GetDataFormAppointmentMoreDocByGroupID,
            objRequest,
            function (data) {
                var objData = data.data;
                var setData = "";
                if (data.status == "success") {
                    for (var i = 0; i < objData.length; i++) {
                        $("#locatename_th").text(objData[i].locatename_th);
                        setData += '<div class="col-12 mt-4">';
                        setData += '<div class="payment">';
                        setData += '<label class="cursor-pointer w-100">';
                        setData +=
                            '<div class="border rounded p-4 bg-white shadow-primary">';
                        setData += '<div class="row">';
                        setData += '<div class="col-lg-9 col-md-8 col-12 pr-md-4 pr-auto">';
                        setData +=
                            '<h4 class="mb-3" style="font-size: 18px; font-weight: bold;">จำนวน ' +
                            objData[i].num_alien_appointment +
                            ' <sapn class="lang_person"></span></h4>';
                        setData += '<h5 class="text-left">';
                        if (
                            objData[i].appointment_datetime != "" ||
                            objData[i].appointment_datetime != null
                        ) {
                            var appointment_datetime = objData[
                                i
                            ].appointment_datetime.replaceAll("-", "/");
                            var appointment_datetime2 =
                                objData[i].appointment_datetime.split("-");
                            var sp_appointment_datetime =
                                appointment_datetime2[2] +
                                " " +
                                MonthString63(appointment_datetime2[1]) +
                                " " +
                                appointment_datetime2[0];
                            setData +=
                                '<span style="color: #6F6F85;">วันที่นัดหมาย :</span><span data_th="' +
                                convert_to_thai_date(appointment_datetime) +
                                '" data_en="' +
                                convert_to_eng_date(appointment_datetime) +
                                '"></span> &nbsp;';
                        }

                        setData +=
                            '<span style="color: #EAEAED;">| &nbsp; </span> <span style="color: #6F6F85;" class="lang_location" >สถานที่ </span> <span>:</span> ' +
                            objData[i].locatename_th +
                            " &nbsp;";
                        setData +=
                            '<span style="color: #EAEAED;">| &nbsp; </span> <span style="color: #6F6F85;" class="lang_upload_doc">อัปโหลดเอกสาร </span> <span>:</span> <span style="color:#E24C86">' +
                            objData[i].num_alien_upload +
                            "</span>/" +
                            objData[i].num_alien_appointment +
                            '<sapn class="lang_person"></span> &nbsp;';
                        setData +=
                            '<span style="color: #EAEAED;">| &nbsp; </span> <span style="color: #6F6F85;" class="lang_confirm_send_doc">ยืนยันการส่งเอกสาร</span> <span>:</span> <span style="color:#E24C86">' +
                            objData[i].num_alien_send +
                            "</span>/" +
                            objData[i].num_alien_appointment +
                            '<sapn class="lang_person"></span> &nbsp;';
                        setData += "</h5>";
                        setData += "</div>";
                        setData +=
                            '<div class="col-lg-3 col-md-4 col-12 d-flex justify-content-lg-end justify-content-sm-start justify-content-center align-items-center pt-md-0 pt-3 column-gap-20">';
                        setData +=
                            '<a href="javascript:void(0)" class="add-btn lang_watch_detail" value="" onclick="GetDataFormAppointmentMoreDocByAppointmentDate(&quot;' +
                            objData[i].appointment_datetime +
                            "&quot;,&quot;" +
                            objData[i].locatename_th +
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
                changeDatalang();
            },
            "json"
        );
    } catch (error) {
        console.error(error);
    }
}
function hiddendiv_Detail_more() {
    $("#div_Detail_more").css("display", "none");
    $("#tabShowTableEmployee").css("display", "block");
    $("#statusAPSS").css("display", "block");
    GetDataFormAppointmentMoreDocByGroupID();
}
function AddNewDocMore() {
    $("#loading").css("display", "block");
    $("#AddNewMore").attr("disabled", true);
    let checkhealth = $("#chkis_upload").val();
    let health_condition = $("#health_condition").val();
    let fileup41 = $("#file_name_41")[0]?.files[0];

    var chkData = true;

    if (health_condition == "NOT_PASS" && !fileup41) {
        $("#loading").delay().fadeOut("")
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
            html: `<span>เนื่องจากผลการตรวจสุขภาพหรือข้อมูลประกันสุขภาพของท่านไม่เป็นไปตามเกณฑ์ที่กำหนด จึงไม่สามารถยื่นคำขอได้</span>`,
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            allowOutsideClick: false,
            allowEscapeKey: false
        }).then((result) => {
            if (result.isConfirmed) {
                $("#UploadFileMOU").modal('hide');
                $("#AddNewMore").prop('disabled',false)
            }
        });
        return;
    }
    if (health_condition == "NO_DATA" && !fileup41) {
        $("#loading").delay().fadeOut("")
        Swal.fire({
            imageUrl: important_data.warningIconUrl,
            title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
            html: `<span>เนื่องจากไม่พบข้อมูลและเอกสารของผลการตรวจสุขภาพหรือประกันสุขภาพในระบบ</span>`,
            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
            reverseButtons: true,
            allowOutsideClick: false,
            allowEscapeKey: false
        }).then((result) => {
            if (result.isConfirmed) {

                $("#AddNewMore").prop('disabled', false)

            }
        });
        return;
    }

    jQuery("[name='file_name_emp']").each(function (index) {
        var fileUpload = jQuery("#" + this.id);
        var ckrequired = jQuery("#requiredvalue_" + this.id).val();
        var file = fileUpload[0].files[0];
        var filebase64 = "";

        if (ckrequired == "Y") {
            if (file != undefined) {

                $("#err_" + this.id + "").css("display", "none");

            } else {
                if (this.id == "file_name_41") {
                    if (checkhealth == "H" && health_condition !== "PASS") {
                        chkData = false;
                        $("#err_" + this.id + "").css("display", "inline");
                        jQuery("#file_link_" + this.id).text("-");
                        $(".link-table a").css("text-decoration", "none");
                    }
                } else {
                    if (file != undefined) {
                        $("#err_" + this.id + "").css("display", "none");
                    } else {
                        chkData = false;
                        $("#err_" + this.id + "").css("display", "inline");
                        jQuery("#file_link_" + this.id).text("-");
                        $(".link-table a").css("text-decoration", "none");
                    }

                }

            }
        }
    });

    if (chkData == true) {
        var formData = new FormData();
        var dataarray = [];
        var arraytype = [];
        var tmpfiletitle = [];

        jQuery("[name='file_name_emp']").each(function (index) {
            dataarray.push(this.id);

            var fileUpload = jQuery("#" + this.id);

            var file_up = fileUpload[0].files[0];
            if (file_up != undefined) {
                var columns = this.id;

                var datatype = columns.split("_");
                datatype = datatype[2];
                var filenamere = file_up.name.replace("'", "_").replace(",", "");
                var file_name_form_type_doc_id = $(
                    "#hidden_form_type_" + this.id
                ).val();
                tmpfiletitle.push(filenamere);
                arraytype.push(file_name_form_type_doc_id);
                formData.append("docfile", file_up);
            }
        });
        jQuery("[name='file_name_emp_group']").each(function (index) {
            dataarray.push(this.id);

            var fileUpload = jQuery("#" + this.id);

            var file_group = fileUpload[0].files[0];
            if (file_group != undefined) {
                var columns = this.id;

                var datatype = columns.split("_");
                datatype = datatype[2];
                var filenamere = file_group.name.replace("'", "_").replace(",", "");
                var file_name_form_type_doc_id = $(
                    "#hidden_form_type_" + this.id
                ).val();
                tmpfiletitle.push(filenamere);
                arraytype.push(file_name_form_type_doc_id);
                formData.append("docfile", file_group);
            }
        });

        formData.append("document_title", tmpfiletitle.toString());
        formData.append("form_type_doc_id", arraytype.toString());
        if (core_datas.default_profile.user_type_id == "5") {
            formData.append("user_id", core_datas.user_id);
        } else {
            formData.append("user_id", core_datas.current_profile.user_id);
        }

        formData.append("group_id", getParameterByName("group_id"));
        formData.append("form_id", $("#form_id").val());
        formData.append("relate_account_id", core_datas.current_profile.profile_id);

        $.ajax({
            method: "post",
            processData: false,
            contentType: false,
            cache: false,
            data: formData,
            url: url_AddNewDocMoreRequestForm41_4_59,
            success: function (response) {
                var dataresponse = JSON.parse(response);
                if (dataresponse.status == "success") {
                    $("#AddNewMore").attr("disabled", false);
                    $("#loading").delay().fadeOut("");
                    $("#UploadFileMOU").modal("hide");
                    GetDataFormAppointmentMoreDocByAppointmentDate(
                        $("#appointment_datetime").val()
                    );

                } else {
                    $("#AddNewMore").attr("disabled", false);
                    $("#loading").delay().fadeOut("");


                    if (dataresponse.status == "filefail") {

                        Swal.fire({
                            imageUrl: important_data.warningIconUrl,
                            title: `<span class="lang_invalid_file_type">ชนิดไฟล์เอกสารไม่ถูกต้อง</span>`,
                            html: `<span class="lang_uploaded_invalid">เนื่องจากชนิดของไฟล์เอกสาร์ที่ท่านอัปโหลดไม่ถูกต้อง กรุณาแก้ไขเอกสารในหน้าการแนบเอกสาร และดำเนินการส่งคำขอใหม่อีกครั้ง</span>`,
                            confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                            didRender: function () {
                                changeDatalang();
                            },
                        }).then((result) => {

                        });
                    } else {
                        Swal.fire({
                            imageUrl: important_data.errorIconUrl,
                            title: '<span class="lang_Notification">แจ้งเตือน</span>',
                            html:
                                '<div style="font-size:16px;"><span>' +
                                dataresponse.msg +
                                "</span></div>",
                            confirmButtonText: '<span class= "lang_lbl_comfirm">ยืนยัน</span>',
                            reverseButtons: true,
                        }).then((result) => {

                        });
                        changeDatalang();
                    }


                }
            },
            error: function (xhr) {
                $("#AddNewMore").attr("disabled", false);
                $("#loading").delay().fadeOut("");

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
    } else {
        $("#AddNewMore").attr("disabled", false);
    }
    $("#loading").delay().fadeOut("");
    changeDatalang();
}
function SendListDocMoreRequestFormToCheck() {
    var formData = new FormData();
    var form_id = [];

    $("#AppointmentMoreDocForSend")
        .DataTable()
        .rows()
        .every(function () {
            let row = this.node();
            $(row)
                .find('input[name="form_sendDoc"]:checked')
                .each(function () {
                    form_id.push(this.id);
                });
        });

    if (core_datas.default_profile.user_type_id == "5") {
        formData.append("user_id", core_datas.user_id);
    } else {
        formData.append("user_id", core_datas.current_profile.user_id);
    }

    formData.append("group_id", getParameterByName("group_id"));
    formData.append("form_id", form_id);

    $.ajax({
        method: "post",
        processData: false,
        contentType: false,
        cache: false,
        data: formData,
        url: url_SendListDocMoreRequestFormToCheck,
        success: function (response) {
            var dataresponse = JSON.parse(response);
            if (dataresponse.status == "success") {
                $("#loading").delay().fadeOut("");
                $("#sendDocument").modal("hide");
                GetDataFormAppointmentMoreDocByAppointmentDate(
                    $("#appointment_datetime").val()
                );
                GetDataFormAppointmentMoreDocByGroupID();

            } else {
                $("#loading").delay().fadeOut("");
                $("#success").removeClass("active");
                $("#payment").addClass("active");
                $("#step5").css("display", "none");
                $("#step4").css("display", "block");
                $("#step4").css("opacity", "1");
                Swal.fire({
                    imageUrl: important_data.errorIconUrl,
                    title: `<span class="lang_failed_request">ยื่นคำขออนุญาตไม่สำเร็จ</span>`,
                    showCancelButton: false,
                    confirmButtonText: "<span class='lang_lbl_ok'>ตกลง</span>",
                    reverseButtons: true,
                    didRender: () => {
                        changeDatalang();
                    },
                }).then((result) => {
                    $("#success").removeClass("active");
                    $("#payment").addClass("active");
                    $("#step5").css("display", "none");
                    $("#step4").css("display", "block");
                    $("#step4").css("opacity", "1");
                });
            }
            changeDatalang();
        },
    });
    changeDatalang();
}
function checkfileUpdateDocMore() {
    var chkData = true;
    jQuery("[name='file_name_emp']").each(function (index) {

        var fileUpload = jQuery("#" + this.id);
        var ckrequired = jQuery("#requiredvalue_" + this.id).val();
        var ckhave_file_current = jQuery("#have_file_current_" + this.id).val();
        var hidden_check_file_name = $("#hidden_check_" + this.id).val();
        var doc_new_lastest = $("#doc_new_lastest_" + this.id).val();
        var hidden_select = $("#hidden_select_" + this.id).val();
        var fileUp = fileUpload[0].files[0];
        var chIDFile = fileUpload[0].id;
        var aTag = document.createElement("a");
        aTag.id = "file_link_" + chIDFile;
        aTag.innerHTML = "<img src='/assets/images/document.svg' class='iconbook'>";
        $("#file_link_" + chIDFile).empty();


        let chkis_upload = $("#chkis_upload").val();
        let health_condition = $("#health_condition").val();

        if (
            this.id === "file_name_41" &&
            hidden_check_file_name === "UC" &&
            (chkis_upload === "N" || chkis_upload === "H")
        ) {

            if (chkis_upload === "H" && fileUp != undefined) {
                return true;
            }

            if (!health_condition || health_condition === "PASS") {
                $("#err_file_name_41").hide();
                return true;
            }

            const messages = {
                NOT_PASS: `เนื่องจากผลการตรวจสุขภาพหรือข้อมูลประกันสุขภาพของท่านไม่เป็นไปตามเกณฑ์ที่กำหนด จึงไม่สามารถยื่นคำขอได้`,
                NO_DATA: `เนื่องจากไม่พบข้อมูลและเอกสารของผลการตรวจสุขภาพหรือประกันสุขภาพในระบบ`
            };

            if (messages[health_condition]) {
                Swal.fire({
                    imageUrl: important_data.warningIconUrl,
                    title: `<span>ไม่สามารถดำเนินการต่อได้</span>`,
                    html: `<span>${messages[health_condition]}</span>`,
                    confirmButtonText: '<span class="lang_lbl_comfirm">ยืนยัน</span>',
                    reverseButtons: true,
                    allowOutsideClick: false,
                    allowEscapeKey: false
                });
                chkData = false;
                return false;
            }
        }


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
                            } else {
                                chkData = false;
                                $("#err_" + chIDFile + "").css("display", "inline");
                            }
                        }
                    } else {
                    }
                }
            }
        }
    });

    var arrayFileRequest = [];
    var chkDataGroupEdit_namelist = true;
    var group_false_namelist = [];
    var group_true_N_namelist = [];
    var chkDataGroupEdit_Y_namelist = true;
    var group_true_namelist = [];
    jQuery("[name='file_name_emp_group']").each(function (index) {
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
        var fileUp = fileUpload[0].files[0];
        var arrayFileNew = [];
        var chIDFile = fileUpload[0].id;

        $("#file_link_" + chIDFile).empty();

        if (!filevalidate.includes(group_required)) {
            filevalidate.push(group_required);
        }


        if (ckrequired == "N") {

            if (ckhave_file_current != "N") {
                if (hidden_check_file_name == "UC") {
                    if (fileUp != undefined) {
                        chkDataGroupEdit_namelist = true;
                        KeepFile(group_required);
                    } else {
                        if (doc_new_lastest == "true") {
                            KeepFile(group_required);
                        } else {
                            jQuery("#file_link_" + chIDFile).text("-");
                            $(".link-table a").css("text-decoration", "none");
                            if (hidden_select == "0") {
                                KeepFile(group_required);
                            } else {
                                if (hidden_check_file_name == "UC") {
                                    chkDataGroupEdit_namelist = false;
                                    group_false_namelist.push(group_required);
                                    Removefiledupilcate(group_required);
                                } else if (chkDataGroupEdit_namelist) {
                                    chkDataGroupEdit_namelist = true;
                                } else {
                                    chkDataGroupEdit_namelist = false;
                                    group_false_namelist.push(group_required);
                                    Removefiledupilcate(group_required);
                                }
                            }
                        }
                    }
                } else {
                    KeepFile(group_required);
                }
            } else {
                if (hidden_check_file_name == "UC") {
                    if (fileUp != undefined) {
                        KeepFile(group_required);

                        if (!group_true_N_namelist.includes(group_required)) {
                            group_true_N_namelist.push(group_required);
                        }
                        chkDataGroupEdit_namelist = true;
                    } else {
                        if (doc_new_lastest == "true") {
                        } else {
                            jQuery("#file_link_" + chIDFile).text("-");
                            $(".link-table a").css("text-decoration", "none");

                            if (hidden_select == "0") {
                                KeepFile(group_required);
                            } else {
                                KeepFile(group_required);
                                chkDataGroupEdit_namelist = true;
                                group_false_namelist = group_false_namelist.filter(function (
                                    value
                                ) {
                                    return !group_true_N_namelist.includes(value);
                                });
                            }
                        }
                    }
                } else {
                    if (fileUp != undefined) {
                        KeepFile(group_required);
                    } else {
                        if (doc_new_lastest == "true") {
                        } else {
                            if (
                                ckrequired == "N" &&
                                hidden_check_file_name != "UC" &&
                                hidden_check_file_name !== undefined
                            ) {
                                if (!arrayfileUploadGroupEdit.includes(group_required)) {
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
                                chkDataGroupEdit_namelist = true;
                                group_false_namelist = group_false_namelist.filter(function (
                                    value
                                ) {
                                    return !group_true_N_namelist.includes(value);
                                });
                                KeepFile(group_required);
                            } else {
                                jQuery("#file_link_" + chIDFile).text("-");
                                $(".link-table a").css("text-decoration", "none");
                                Removefiledupilcate(group_required);
                            }
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
                    if (!group_true_namelist.includes(group_required)) {
                        group_true_namelist.push(group_required);
                    }
                    chkDataGroupEdit_Y_namelist = true;
                    KeepFile(group_required);
                } else {
                    if (hidden_select != "0" && hidden_select != "Y") {
                        chkDataGroupEdit_Y_namelist = true;
                        KeepFile(group_required);
                    } else {
                        if (group_true_namelist.length != 0) {
                            chkDataGroupEdit_Y_namelist = true;
                            group_false_namelist = group_false_namelist.filter(function (
                                value
                            ) {
                                return !group_true_namelist.includes(value);
                            });
                        } else {
                            chkDataGroupEdit_Y_namelist = false;
                            group_false_namelist.push(group_required);
                        }
                    }
                    jQuery("#file_link_" + chIDFile).text("-");
                    $(".link-table a").css("text-decoration", "none");
                }
            } else {
                if (hidden_check_file_name == "UC") {
                    if (fileUp != undefined) {
                        KeepFile(group_required);
                    } else {
                        if (doc_new_lastest == "true") {
                            KeepFile(group_required);
                        } else {
                            jQuery("#file_link_" + chIDFile).text("-");
                            $(".link-table a").css("text-decoration", "none");
                            if (hidden_select == "0") {
                                KeepFile(group_required);
                            } else {
                                chkDataGroupEdit_Y_namelist = false;
                                group_false_namelist.push(group_required);
                                Removefiledupilcate(group_required);
                            }
                        }
                    }
                } else {
                    KeepFile(group_required);
                    arrayFileRequest.push(ckrequired);
                    if (fileUp != undefined) {
                        arrayFileRequest.push(objectUrl);
                    } else {
                        if (group_false_namelist.length != 0) {
                            chkDataGroupEdit_Y_namelist = false;
                            let indexToRemove = group_false_namelist.indexOf(group_required);
                            if (indexToRemove !== -1) {
                                group_false_namelist.splice(indexToRemove, 1);
                            }
                        } else {
                            chkDataGroupEdit_Y_namelist = true;
                            if (!group_true_namelist.includes(group_required)) {
                                group_true_namelist.push(group_required);
                            }
                        }

                        KeepFile(group_required);
                    }
                }
            }
        }
    });
    var allFieldsValid = true;
    if (
        chkDataGroupEdit_namelist == false ||
        chkDataGroupEdit_Y_namelist == false
    ) {
        filevalidate.forEach(function (fileName) {
            var errorElement = $("#err_file_name_" + fileName);
            if (group_false_namelist.includes(fileName)) {
                allFieldsValid = false;
                errorElement.css("display", "inline");
            } else {
                errorElement.css("display", "none");
            }
        });
    } else {
        filevalidate.forEach(function (fileName) {
            var errorElement = $("#err_file_name_" + fileName);
            if (arrayfileUploadGroupEdit.includes(fileName)) {
                errorElement.css("display", "none");
            } else {

                allFieldsValid = false;
                errorElement.css("display", "inline");
            }
        });
    }

    if (chkData && allFieldsValid) {
        UpdateDocMoreRequestForm46_59();
    }
}
function UpdateDocMoreRequestForm46_59() {
    $("#loading").css("display", "block");

    var formData = new FormData();
    var dataarray = [];
    var arraytype = [];
    var tmpfiletitle = [];

    jQuery("[name='file_name_emp']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file_up = fileUpload[0].files[0];
        if (file_up != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file_up.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#hidden_form_type_" + this.id).val();
            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file_up);
        }
    });
    jQuery("[name='file_name_emp_group']").each(function (index) {
        dataarray.push(this.id);

        var fileUpload = jQuery("#" + this.id);

        var file_group = fileUpload[0].files[0];
        if (file_group != undefined) {
            var columns = this.id;

            var datatype = columns.split("_");
            datatype = datatype[2];
            var filenamere = file_group.name.replace("'", "_").replace(",", "");
            var file_name_form_type_doc_id = $("#hidden_form_type_" + this.id).val();
            tmpfiletitle.push(filenamere);
            arraytype.push(file_name_form_type_doc_id);
            formData.append("docfile", file_group);
        }
    });

    formData.append("document_title", tmpfiletitle.toString());
    formData.append("form_type_doc_id", arraytype.toString());
    if (core_datas.default_profile.user_type_id == "5") {
        formData.append("user_id", core_datas.user_id);
    } else {
        formData.append("user_id", core_datas.current_profile.user_id);
    }

    formData.append("group_id", getParameterByName("group_id"));
    formData.append("form_id", $("#form_id").val());
    formData.append("relate_account_id", core_datas.current_profile.profile_id);
    formData.append("update_doc_health", $("#doc_new_lastest").val() || "-");

    $.ajax({
        method: "post",
        processData: false,
        contentType: false,
        cache: false,
        data: formData,
        url: url_UpdateDocMoreRequestForm41_4_59,
        success: function (response) {
            var dataresponse = JSON.parse(response);
            if (dataresponse.status == "success") {
                $("#loading").delay().fadeOut("");
                $("#UploadFileMOU").modal("hide");
                GetDataFormAppointmentMoreDocByAppointmentDate(
                    $("#appointment_datetime").val()
                );

            } else {
                $("#loading").delay().fadeOut("");
                $("#success").removeClass("active");
                $("#payment").addClass("active");
                $("#step5").css("display", "none");
                $("#step4").css("display", "block");
                $("#step4").css("opacity", "1");
                Swal.fire({
                    imageUrl: important_data.errorIconUrl,
                    title: `<span class="lang_failed_request">ยื่นคำขออนุญาตไม่สำเร็จ</span>`,
                    showCancelButton: false,
                    confirmButtonText: "<span class='lang_lbl_ok'>ตกลง</span>",
                    reverseButtons: true,
                    didRender: () => {
                        changeDatalang();
                    },
                }).then((result) => {
                    $("#success").removeClass("active");
                    $("#payment").addClass("active");
                    $("#step5").css("display", "none");
                    $("#step4").css("display", "block");
                    $("#step4").css("opacity", "1");
                });
            }
        },
        error: function (xhr) {
            $("#loading").delay().fadeOut("");
            $("#success").removeClass("active");
            $("#payment").addClass("active");
            $("#step5").css("display", "none");
            $("#step4").css("display", "block");
            $("#step4").css("opacity", "1");
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
    changeDatalang();
}
function openPageEditMOU() {
    var status = $("#statusGroup").val();
    var group_id = getParameterByName("group_id");
    var form_type = getParameterByName("form_type");
    if (status == "WA") {
        $("#div_tab1").removeClass("active");
        $("#tab_default_1").removeClass("active");
        $("#div_tab2").removeClass("active");
        $("#tab_default_2").removeClass("active");
        $("#div_tab2_1").removeClass("active");
        $("#tab_default_2_1").removeClass("active");
        $("#div_tab4").removeClass("active");
        $("#tab_default_4").removeClass("active");
        $("#div_tab5").removeClass("active");
        $("#tab_default_5").removeClass("active");
        $("#div_tab6").removeClass("active");
        $("#tab_default_6").removeClass("active");
        $("#div_tab3").addClass("active");
        $("#tab_default_3").addClass("active");
    } else if (status == "WP" || status == "WP3") {
        $("#div_tab1").removeClass("active");
        $("#tab_default_1").removeClass("active");
        $("#div_tab2").removeClass("active");
        $("#tab_default_2").removeClass("active");
        $("#div_tab2_1").removeClass("active");
        $("#tab_default_2_1").removeClass("active");
        $("#div_tab3").removeClass("active");
        $("#tab_default_3").removeClass("active");
        $("#div_tab5").removeClass("active");
        $("#tab_default_5").removeClass("active");
        $("#div_tab6").removeClass("active");
        $("#tab_default_6").removeClass("active");
        $("#div_tab4").addClass("active");
        $("#tab_default_4").addClass("active");
    } else {
        if (important_data.is_mobile == "Y") {
            var status = getParameterByName("status");
            var group_id = getParameterByName("group_id");
            var user_id = getParameterByName("user_id");
            var form_type = getParameterByName("form_type");
            location.href =
                "/RequestForm41/EditFormRequest41?App=Y&status=" +
                status +
                "&group_id=" +
                group_id +
                "&user_id=" +
                user_id;
        } else {
            location.href =
                "/RequestForm41/EditFormRequest41?status=" +
                status +
                "&group_id=" +
                group_id +
                "&form_type=" +
                form_type;
        }
    }
}
function GetDocFilePathByDocumentIDMOU(
    key_id_data,
    document_id,
    document_type_id
) {
    var app = important_data.is_mobile;
    let objRequest = {
        key_id: key_id_data,
        document_id: document_id,
        form_type_doc_id: document_type_id,
        version_id: "",
    };
    var request = jQuery.post(
        url_GetDocFilePathByDocumentID,
        objRequest,
        function (data) {
            if (data.status == "success") {
                var objData = data.data;
                if (app == "Y") {
                    mobileScriptChannel.postMessage("openFile:" + objData.document_path);
                } else {
                    window.open(objData.document_path, "_blank");
                }
            }
        },
        "json"
    );
}
let loadedTabs = {};

$('a[data-toggle="tab"]').on('click', function (e) {
    let target = $(e.target).attr("href");
    var group_id = getParameterByName('group_id')


    if (loadedTabs[target]) return;


    loadedTabs[target] = true;


    switch (target) {
        case '#tab_default_1':

            break;
        case '#tab_default_2':

            break;
        case '#tab_default_3':

            break;
        case '#tab_default_4':
            if (
                (objData_setPayment.request_by_user_id == core_datas.user_id &&
                    objData_setPayment.request_by_user_id_proxy == "") ||
                objData_setPayment.request_by_user_id_proxy == core_datas.user_id
            ) {

                GetDataPaymentInquiryTransactionMOU(objData_setPayment.group_id, false);
            } else {
                GetDataPaymentInquiryTransactionMOU(objData_setPayment.group_id, true);

            }
            break;
        case '#tab_default_5':
            GetDataFormAppointmentByGroupID()
            break;
        case '#tab_default_6':
            GetDetailDocumentListMOU(group_id);
            break;
        case '#tab_default_7':
            AppealList(group_id);
            break;
    }
});
