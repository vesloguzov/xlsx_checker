/* Javascript for XlsxCheckerXBlock. */
function XlsxCheckerXBlock(runtime, element) {

   var upload_student_file = runtime.handlerUrl(element, 'upload_student_file');

   var download_student_file = runtime.handlerUrl(element, 'download_student_file');
   $('.download_student_file', element).attr('href', download_student_file);
   
   var student_filename = runtime.handlerUrl(element, 'student_filename');
    
   // var download_instruction = runtime.handlerUrl(element, 'download_instruction');
   // $('.download_instruction', element).attr('href', download_instruction);

   var student_submit = runtime.handlerUrl(element,'student_submit');

    function successLoadStudentFile(result) {
        $.ajax({
            url: student_filename,
            type: 'GET',
            success: function(result){
                $('.download_student_file', element).html(result["student_filename"]);
            }

        });
    }

    $(':button.upload-student-file').on('click', function() {
        $.ajax({
            url: upload_student_file,
            type: 'POST',
            data: new FormData($('form.student')[0]),
            cache: false,
            contentType: false,
            processData: false,
            xhr: function() {
                var myXhr = $.ajaxSettings.xhr();
                if (myXhr.upload) {
                    myXhr.upload.addEventListener('progress', function(evt) {
                        if (evt.lengthComputable) {
                            //Сделать лоадер
                        }
                    } , false);
                }
                return myXhr;
            },
            success: successLoadStudentFile

        });
    });

    function successCheck(result) {
        $('.analyze-all', element).empty()
        $('.analyze-errors', element).empty()
        $('.global-errors', element).empty()
        analyze = {}
        analyze = result["analyze"];
        console.log(analyze);
        if(analyze["errors"].length > 0){
                var errors = document.createElement("div");
                analyze["errors"].forEach(function(item, i, arr) {
                    var criterion_element_error = document.createElement("div");
                    criterion_element_error.innerHTML = item;
                    errors.appendChild(criterion_element_error);

                });
            $('.global-errors', element).append(errors)
        }
        else{
            Object.keys(analyze).map(function(item, i, arr) {
                var one_obj = analyze[item]
                var criterion_error = document.createElement("div");
                criterion_error.className = 'error' + item;

                var criterion_all = document.createElement("div");
                criterion_all.className = item;

                if (item == "conditional_formatting"){
                        var criterion_element_all = document.createElement("div");
                        criterion_element_all.innerHTML = one_obj["message"];
                        criterion_all.appendChild(criterion_element_all);
                        if (one_obj["status"] == false){
                            var criterion_element_error = document.createElement("div");
                            criterion_element_error.innerHTML = one_obj["message"];

                            criterion_error.appendChild(criterion_element_error);
                        }
                }
                if (item == "formats"){
                    Object.keys(one_obj).map(function(item, i, arr) {
                        var criterion_element_all = document.createElement("div");
                        criterion_element_all.innerHTML = one_obj[item]["message"];
                        criterion_all.appendChild(criterion_element_all);
                        if (one_obj[item]["status"] == false){
                            var criterion_element_error = document.createElement("div");
                            criterion_element_error.innerHTML = one_obj[item]["message"];

                            criterion_error.appendChild(criterion_element_error);
                        }
                    });
                }
                if (item == "functions"){
                    Object.keys(one_obj).map(function(item, i, arr) {
                        var criterion_element_all = document.createElement("div");
                        criterion_element_all.innerHTML = one_obj[item]["message"];
                        criterion_all.appendChild(criterion_element_all);
                        if (one_obj[item]["status"] == false){
                            var criterion_element_error = document.createElement("div");
                            criterion_element_error.innerHTML = one_obj[item]["message"];

                            criterion_error.appendChild(criterion_element_error);
                        }
                    });
                }
                if (item == "formulas"){
                    Object.keys(one_obj).map(function(item, i, arr) {
                        var criterion_element_all = document.createElement("div");
                        criterion_element_all.innerHTML = one_obj[item]["message"];
                        criterion_all.appendChild(criterion_element_all);
                        if (one_obj[item]["status"] == false){
                            var criterion_element_error = document.createElement("div");
                            criterion_element_error.innerHTML = one_obj[item]["message"];

                            criterion_error.appendChild(criterion_element_error);
                        }
                    });
                }
        
                $('.analyze-errors', element).append(criterion_error);
                $('.analyze-all', element).append(criterion_all);
            });
        }

    }

    $(element).find('.Check').bind('click', function() {
        console.log("CHECK");
        $.ajax({
            type: "POST",
            url: student_submit,
            data: JSON.stringify({"picture": "resultImage" }),
            success: successCheck
        });

    });

    $(function ($) {
        /* Here's where you'd do things on page load. */
    });
}
