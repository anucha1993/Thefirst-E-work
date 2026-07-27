async function GetDataRequestFormEtrackingForAlien() {
    table = $("#datatableE_Tracking").DataTable({
        destroy: true,
        paging: true,
        ordering: false,
        searching: false,
        processing: true,
        serverSide: true,
        info: true,
        lengthChange: true,
        lengthMenu: [10, 25, 50, 100],
        drawCallback: function (settings) {
            console.log(settings);
            console.log(settings?.json?.statusAmont);
       
            function checkNumber(element, value) {
                element.text(value);
                //if (value && value != 0) {
                //    element.text(value);
                //} 
            }
            checkNumber($("#numberform"), settings?.json?.recordsTotal);
            checkNumber($("#number_WP"), settings?.json?.statusAmont?.wp || 0);
            checkNumber($("#number_WCOSNA"), settings?.json?.statusAmont?.wcosna || 0);
            checkNumber($("#number_NCOS"), settings?.json?.statusAmont?.ncos || 0);
            checkNumber($("#number_APSS"), settings?.json?.statusAmont?.apss || 0);
            checkNumber($("#number_CCOS"), settings?.json?.statusAmont?.ccos || 0);
            checkNumber($("#number_AP"), settings?.json?.statusAmont?.ap || 0);
            checkNumber($("#number_WA"), settings?.json?.statusAmont?.wa || 0);
            checkNumber($("#number_SS"), settings?.json?.statusAmont?.ss || 0);
            changeDatalang();
        },
        columns: [
            { data: "group_id" },
            { data: "group_id" },
            { data: "form_type_name_th", className: "text-left" },
            { data: "alien_first_name", className: "text-left" },
            { data: "createdtimestamp" },
            { data: "group_status_id" },
            { data: "group_id" },
        ],
        columnDefs: [
            
            {
                targets: 0,
                render: function (data, type, row, meta) {
                    return `${meta.row + meta.settings._iDisplayStart + 1}`;
                },
            },
            {
                targets: 1,
                render: function (data, type, row, meta) {

                    let html = "";
                    html += `<span>${row.group_id}</span><br><span style="color:#6F6F85;font-size:14px;"><span data_th='${row.request_by_full_name_th}' data_en='${row.request_by_full_name_en}'></span></span><br>`;

                    if (row.id) {
                        html += `<a href='javascript:void(0)' onclick='openDetail(${row.request_by_user_id},&quot;${row.group_id}&quot;,&quot;${row.group_status_id}&quot;,&quot;${row.form_type_id}&quot;,&quot;${row.institution_id}&quot;,&quot;${row.id}&quot;)'><span class="btndetail">ดูรายละเอียด <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M8.3332 5L7.1582 6.175L10.9749 10L7.1582 13.825L8.3332 15L13.3332 10L8.3332 5Z" fill="#fff"></path></svg></span></a >`;

                    } else {
                        html += `<a href='javascript:void(0)' onclick='openDetail(${row.request_by_user_id},&quot;${row.group_id}&quot;,&quot;${row.group_status_id}&quot;,&quot;${row.form_type_id}&quot;,&quot;${row.institution_id}&quot;)'> <span class="btndetail">ดูรายละเอียด <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M8.3332 5L7.1582 6.175L10.9749 10L7.1582 13.825L8.3332 15L13.3332 10L8.3332 5Z" fill="#fff"></path></svg></span></a >`;

                    }

                    return html;
                },
            },

            {
                targets: 2,
                render: function (data, type, row, meta) {

                    let html = "";

                    if (row.form_type_id == "MT_63_2" || row.form_type_id == "MT_63_2_AGN_RENEWAL") {
                        html = `
                                <span data_th="${row.form_type_name_th?.toSafeJsonHtml()}" data_en="${row.form_type_name_en}"></span>
                                <br>
                                <span style="color:#6F6F85;font-size:14px;" data_th="${row.form_desc_th?.toSafeJsonHtml()}" data_en="${row.form_desc_en}"></span>
                                <br>
                                <span style="color:#6F6F85;font-size:14px;" 
                                    data_th="${row.alien_full_name_th?.toSafeJsonHtml() || row.alien_full_name_en}" 
                                    data_en="${row.alien_full_name_en || row.alien_full_name_th}"></span>
                            `;
                    }
                    else if (row.form_type_id == "NAMELIST_SURVEY_19") {
                        html = `
                            <span data_th="${row.form_type_name_th?.toSafeJsonHtml()}" data_en="${row.form_type_name_en}"></span>
                            <br>
                            <span style="color:#6F6F85;font-size:14px;">
                                <span class="lang_submission_period_from">ระยะเวลายื่นตั้งเเต่</span> 
                                <span data_th="${convert_to_thai_date(row.submit_start_dt)}" data_en="${convert_to_eng_date(row.submit_start_dt)}"></span>
                                <span class="lang_lbl_to"></span> 
                                <span data_th="${convert_to_thai_date(row.submit_end_dt)}" data_en="${convert_to_eng_date(row.submit_end_dt)}"></span>
                            </span>
                            <br>
                            <span class="badgestatus" style="color:#DCB745;background-color:#FEFAED">
                                <span class="lang_alien_documents_submit">ยื่นเอกสารต่างด้าวเเล้ว</span> 
                                ${row.num_namelist_pay_success}/${row.num_namelist_all}
                            </span>
    `;
                    } else if (row.form_type_id == "MT_63_2_AGN_CHANGE") {
                        html = `
                                <span data_th="${row.form_type_name_th?.toSafeJsonHtml()}" data_en="${row.form_type_name_en}"></span>
                                <br>
                                <span style="color:#6F6F85;font-size:14px;" 
                                    data_th="${row.alien_full_name_th?.toSafeJsonHtml() || row.alien_full_name_en}" 
                                    data_en="${row.alien_full_name_en || row.alien_full_name_th}"></span>
                            `;                      
                    }
                    else {
                        let desc = `<span style="color:#6F6F85;font-size:14px;" data_th="${row.form_desc_th?.toSafeJsonHtml()}" data_en="${row.form_desc_en}"></span>`;
                        let inform = row.inform_list_name ? `<br><span style="color:#6F6F85;font-size:14px;">${row.inform_list_name}</span>` : "";

                        html = `
                            <span data_th="${row.form_type_name_th?.toSafeJsonHtml()}" data_en="${row.form_type_name_en}"></span>
                            <br>
                            ${desc}
                            ${inform}
    `;
                    }

                    // ✅ เพิ่มท้ายถ้า is_proxy = Y
                    if (row.is_proxy === "Y") {
                        if (row.form_type_id == "MT_13_EXIT" || row.form_type_id == "MT_13_1_INFORM") {
                            html += `<br><span class="span-box-justify-content"><span class="lang_do_emp">กระทำการแทนนายจ้าง</span></span>`;
                        } else {
                            html += `<br><span class="span-box-justify-content"><span class="lang_do_alien">กระทำการแทนคนต่างด้าว</span></span>`;
                        }
                    }

                    return html;

                },
            },
            {
                targets: 3,
                render: function (data, type, row, meta) {
                    return (
                        '<span data_th="' +
                        convert_to_thai_datetime(row.created_timestamp) +
                        '" data_en="' +
                        convert_to_eng_datetime(row.created_timestamp) +
                        '"></span>'
                    );
                },
            },
            {
                targets: 4,
                render: function (data, type, row, meta) {
                    return (
                        '<span data_th="' +
                        convert_to_thai_datetime(row.updated_timestamp) +
                        '" data_en="' +
                        convert_to_eng_datetime(row.updated_timestamp) +
                        '"></span>'
                    );
                },
            },
            {
                targets: 5,
                render: function (data, type, row, meta) {
                    return (
                        '<span class="badgestatus" style="color:' + row.status_color_font + ";background-color:" + row.status_color + '" data_th="' + row.transection_status_th?.toSafeJsonHtml() + '" data_en="' +
                        row.transection_status_en +
                        '"></span><br><span class="text_tran" data_th="' +
                        row.transection_activity_th?.toSafeJsonHtml() +
                        '" data_en="' +
                        row.transection_activity_en +
                        '"></span></span>'
                    );
                },
            },

            {
                targets: 6,
                render: function (data, type, row, meta) {
                    if (row.group_id == "67125000000131") {
                        console.log(row);
                    }
                    if (row.id) {
                        return `<a href='javascript:void(0)' onclick='openDetail(${row.request_by_user_id},&quot;${row.group_id}&quot;,&quot;${row.group_status_id}&quot;,&quot;${row.form_type_id}&quot;,&quot;${row.institution_id}&quot;,&quot;${row.id}&quot;)'><svg xmlns='http://www.w3.org/2000/svg' width='20' height='20' viewBox='0 0 20 20' fill='none'><path d='M8.3332 5L7.1582 6.175L10.9749 10L7.1582 13.825L8.3332 15L13.3332 10L8.3332 5Z' fill='#E24C86'/></svg></a >`;

                    } else {
                        return `<a href='javascript:void(0)' onclick='openDetail(${row.request_by_user_id},&quot;${row.group_id}&quot;,&quot;${row.group_status_id}&quot;,&quot;${row.form_type_id}&quot;,&quot;${row.institution_id}&quot;)'><svg xmlns='http://www.w3.org/2000/svg' width='20' height='20' viewBox='0 0 20 20' fill='none'><path d='M8.3332 5L7.1582 6.175L10.9749 10L7.1582 13.825L8.3332 15L13.3332 10L8.3332 5Z' fill='#E24C86'/></svg></a >`;

                    }
                },
            },
        ],
        ajax: {
            url: url_GetDataRequestFormEtrackingForAlien,
            type: "POST",
            dataType: "json",
            data: function (d) {
                console.log(d)
                var search_start_date = "";
                var search_end_date = "";
                if ($("#search_start_date").val() != "") {
                    search_start_date = ConvertDate($("#search_start_date").val())
                }
                if ($("#search_end_date").val() != "") {
                    search_end_date = ConvertDate($("#search_end_date").val())
                }
                d.user_id = core_datas.user_id;
                d.relate_account_id = core_datas.current_profile.profile_id;
                d.search_group_id = $("#search_group_id").val();
                d.search_start_date = search_start_date;
                d.search_end_date = search_end_date;
                d.filter_action_by = $("#Filter_request_submit :selected").val();
                d.search_mt = $("#Filter_request_list :selected").val();
                d.employer_id = "";
                d.status_id = jQuery("[name = 'checkbox_formstatus_alien']:checked")
                    .map(function () {
                        return $(this).val();
                    })
                    .get()
                    .toString();
                d.page = null;
                d.limit = null;
            },
        },
    });
}