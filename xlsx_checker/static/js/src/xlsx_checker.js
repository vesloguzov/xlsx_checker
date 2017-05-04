/* Javascript for XlsxCheckerXBlock. */
function XlsxCheckerXBlock(runtime, element, data) {

   var xlsx_analyze = data["xlsx_analyze"];  
   var lab_scenario = data["lab_scenario"];
   var student_xlsx_name = data["student_xlsx_name"];

   if(!jQuery.isEmptyObject(xlsx_analyze)){
        // console.log("KEK@", jQuery.isEmptyObject(xlsx_analyze));
		showBlockAnalyze(xlsx_analyze);
   }
   else{
    console.log("YTEK@", xlsx_analyze);
       $('.block-analyze', element).hide();
   }
   
   if(student_xlsx_name){
    console.log("Y");
        $('.current-student-file', element).show();
   }
   else{
    console.log("N");
        $('.current-student-file', element).hide();
   }

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
                $('.current-student-file', element).show();
                $('input[name="studentFile"]').val("");
            }

        });
    }

    $(':button.upload-student-file', element).on('click', function() {
        var file = $('input[name="studentFile"]').val().trim();
        if(file){
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
        }
        else{
            alert("Необходимо  выбрать документ!");
        }
    });

    function showLab1FullAnalyze(analyze_object){
        analyze = analyze_object;
        $('.analyze-all', element).empty();
        // delete analyze["errors"];

        Object.keys(analyze).map(function(item, i, arr) {
                var one_obj = analyze[item]
                var criterion_all = document.createElement("div");
                criterion_all.className = item + " criterion-block";    

                var criterion_header = document.createElement("div");
                if (item == "conditional_formatting"){
                        criterion_header.innerHTML = "Условное форматирование";
                        criterion_all.appendChild(criterion_header);

                        var criterion_element_all = document.createElement("p");
                        criterion_element_all.innerHTML = one_obj["message"];
                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj["status"];
                        criterion_all.appendChild(criterion_element_all);
                }
                if (item == "formats"){
                    criterion_header.innerHTML = "Форматирование ячеек";
                    criterion_all.appendChild(criterion_header);

                    Object.keys(one_obj).map(function(item, i, arr) {
                        var criterion_element_all = document.createElement("p");
                        criterion_element_all.innerHTML = one_obj[item]["message"];
                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item]["status"];
                        criterion_all.appendChild(criterion_element_all);
                    });
                }
                if (item == "functions"){
                    criterion_header.innerHTML = "Применение функций";
                    criterion_all.appendChild(criterion_header);

                    Object.keys(one_obj).map(function(item, i, arr) {
                        var criterion_element_all = document.createElement("p");
                        criterion_element_all.innerHTML = one_obj[item]["message"];
                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item]["status"];
                        criterion_all.appendChild(criterion_element_all);
                    });
                }
                if (item == "formulas"){
                    criterion_header.innerHTML = "Использование формул";
                    criterion_all.appendChild(criterion_header);

                    Object.keys(one_obj).map(function(item, i, arr) {
                        var criterion_element_all = document.createElement("p");
                        criterion_element_all.innerHTML = one_obj[item]["message"];
                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item]["status"];
                        criterion_all.appendChild(criterion_element_all);
                    });
                }
                if (item == "charts"){
                    criterion_header.innerHTML = "Графики";
                    criterion_all.appendChild(criterion_header);

                    var bar_obj = one_obj["bar_chart"];

                    if(bar_obj["errors"].length != 0){
                        bar_obj["errors"].forEach(function(item, i, arr) {
                            var error = document.createElement("p");
                            error.innerHTML = "Гистограмма: " + item;
                            error.className = 'one-criterion criterion-complete-false';
                            criterion_all.append(error);
                        });
                    }
                    else{
                        Object.keys(bar_obj).map(function(item, i, arr) {
                            if(item!="errors"){
                                var criterion_element_all = document.createElement("p");
                                criterion_element_all.innerHTML = "Гистограмма: " + bar_obj[item]["message"];
                                criterion_element_all.className = 'one-criterion criterion-complete-'+bar_obj[item]["status"];
                                criterion_all.appendChild(criterion_element_all);
                            }
                        });
                    }

                }
                
                $('.analyze-all', element).append(criterion_all);
            });
    }

    function showLab2FullAnalyze(analyze_object){
        analyze = analyze_object;
        $('.analyze-all', element).empty();
        // delete analyze["errors"];

        Object.keys(analyze).map(function(item, i, arr) {
                var one_obj = analyze[item]
                var criterion_all = document.createElement("div");
                criterion_all.className = item + " criterion-block";    

                var criterion_header = document.createElement("div");

               if (item == "ws1"){
                    criterion_header.innerHTML = "График 1";
                    criterion_all.appendChild(criterion_header);
                    Object.keys(one_obj).map(function(item, i, arr) {
                        if (item == "data"){
                            var criterion_element_all = document.createElement("p");
                            criterion_element_all.innerHTML = one_obj[item]["message"];
                            criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item]["status"];
                            criterion_all.appendChild(criterion_element_all);
                        }
                        if (item == "graphic"){
                            if(one_obj[item]["errors"].length != 0){
                                one_obj[item]["errors"].forEach(function(subitem, i, arr) {
                                    var error = document.createElement("p");
                                    error.innerHTML = subitem;
                                    error.className = 'one-criterion criterion-complete-false';
                                    criterion_all.append(error);
                                });
                            }

                            else{
                                Object.keys(one_obj[item]).map(function(subitem, i, arr) {
                                    if(subitem!="errors"){
                                        var criterion_element_all = document.createElement("p");
                                        criterion_element_all.innerHTML = one_obj[item][subitem]["message"];
                                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item][subitem]["status"];
                                        criterion_all.appendChild(criterion_element_all);
                                    }
                                });
                            }

                        }
                    });
                }

                if (item == "ws2"){
                    criterion_header.innerHTML = "График 2";
                    criterion_all.appendChild(criterion_header);
                    Object.keys(one_obj).map(function(item, i, arr) {
                        if (item == "data"){
                            var criterion_element_all = document.createElement("p");
                            criterion_element_all.innerHTML = one_obj[item]["message"];
                            criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item]["status"];
                            criterion_all.appendChild(criterion_element_all);
                        }
                        if (item == "graphic"){
                            if(one_obj[item]["errors"].length != 0){
                                one_obj[item]["errors"].forEach(function(subitem, i, arr) {
                                    var error = document.createElement("p");
                                    error.innerHTML = subitem;
                                    error.className = 'one-criterion criterion-complete-false';
                                    criterion_all.append(error);
                                });
                            }

                            else{
                                Object.keys(one_obj[item]).map(function(subitem, i, arr) {
                                    if(subitem!="errors"){
                                        var criterion_element_all = document.createElement("p");
                                        criterion_element_all.innerHTML = one_obj[item][subitem]["message"];
                                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item][subitem]["status"];
                                        criterion_all.appendChild(criterion_element_all);
                                    }
                                });
                            }

                        }
                    });
                }
                
                $('.analyze-all', element).append(criterion_all);
            });
    }

    function showLab3FullAnalyze(analyze_object){
        analyze = analyze_object;
        $('.analyze-all', element).empty();
        // delete analyze["errors"];

        Object.keys(analyze).map(function(item, i, arr) {
                var one_obj = analyze[item]
                var criterion_all = document.createElement("div");
                criterion_all.className = item + " criterion-block";    

                var criterion_header = document.createElement("div");
                if (item == "sort"){
                        criterion_header.innerHTML = "Сортировка товара";
                        criterion_all.appendChild(criterion_header);

                        var criterion_element_all = document.createElement("p");
                        criterion_element_all.innerHTML = one_obj["message"];
                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj["status"];
                        criterion_all.appendChild(criterion_element_all);
                }
                
                if (item == "results"){
                        criterion_header.innerHTML = "Лист итогов";
                        criterion_all.appendChild(criterion_header);

                        var criterion_element_all = document.createElement("p");
                        criterion_element_all.innerHTML = one_obj["message"];
                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj["status"];
                        criterion_all.appendChild(criterion_element_all);
                }

                if (item == "formats"){
                    criterion_header.innerHTML = "Форматирование ячеек";
                    criterion_all.appendChild(criterion_header);
                    Object.keys(one_obj).map(function(item, i, arr) {
                        var criterion_element_all = document.createElement("p");
                        criterion_element_all.innerHTML = one_obj[item]["message"];
                        criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item]["status"];
                        criterion_all.appendChild(criterion_element_all);
                    });
                }

               if (item == "filters"){
                    criterion_header.innerHTML = "Фильтрация";
                    criterion_all.appendChild(criterion_header);
                    Object.keys(one_obj).map(function(item, i, arr) {
                        if (item == "custom"){
                            var criterion_element_all = document.createElement("p");
                            criterion_element_all.innerHTML = one_obj[item]["message"];
                            criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item]["status"];
                            criterion_all.appendChild(criterion_element_all);
                        }
                        if (item == "year"){
                            var criterion_element_all = document.createElement("p");
                            criterion_element_all.innerHTML = one_obj[item]["message"];
                            criterion_element_all.className = 'one-criterion criterion-complete-'+one_obj[item]["status"];
                            criterion_all.appendChild(criterion_element_all);
                        }
                    });
                }
                
                $('.analyze-all', element).append(criterion_all);
            });
    }

    function showLabErrors(errors){
    	$('.analyze-errors', element).empty();
    	errors.forEach(function(item, i, arr) {
            var error = document.createElement("p");
            error.innerHTML = item;
            error.className = 'one-criterion criterion-complete-false';
            $('.analyze-errors', element).append(error);
		});
    }

    function showBlockAnalyze(analyze_object){
    	$('.full-analyze', element).hide();
        $('.errors-analyze', element).hide();
    	

        console.log(analyze_object);


    	if(lab_scenario == 1){
        	if(analyze_object["errors"].length == 0){
            	showLab1FullAnalyze(analyze_object);
            	$('.full-analyze', element).show();
        	}
        	else{
        		showLabErrors(analyze_object["errors"]);
        		$('.errors-analyze', element).show();
        	}
        }
        else if(lab_scenario == 2){
        	if(analyze_object["errors"].length == 0){
            	showLab2FullAnalyze(analyze_object);
            	$('.full-analyze', element).show();
           	}
        	else{
        		showLabErrors(analyze_object["errors"]);
        		$('.errors-analyze', element).show();
        	}
        }
        else if(lab_scenario == 3){
        	if(analyze_object["errors"].length == 0){
            	showLab3FullAnalyze(analyze_object);
            	$('.full-analyze', element).show();
        	}
            else{
        		showLabErrors(analyze_object["errors"]);
        		$('.errors-analyze', element).show();
        	}
		}
		$('.block-analyze', element).show(300);
    }

    function successCheck(result) {
        updatePointsAttempts(result);
        analyze_object = result["xlsx_analyze"];
        showBlockAnalyze(analyze_object);
    }

    $(element).find('.Check').bind('click', function() {
        $('.block-analyze', element).hide();
        $('.full-analyze', element).hide();
        $('.errors-analyze', element).hide();

        $('input[name="studentFile"]').val("");

        $.ajax({
            type: "POST",
            url: student_submit,
            data: JSON.stringify({"picture": "resultImage" }),
            success: successCheck
        });

    });

    function updatePointsAttempts(result) {
        $('.attempts', element).text(result.attempts);
        $(element).find('.weight').html('Набрано баллов: <me-span class="points"></span>');
        $('.points', element).text(result.points + ' из ' + result.weight);

        if (result.max_attempts && result.max_attempts <= result.attempts) {
            $('.Check', element).remove();
            $('.Save', element).remove();
        };
    };

    $(function ($) {
        /* Here's where you'd do things on page load. */
    });
}
