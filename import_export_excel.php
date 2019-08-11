<?php
/**
* Plugin Name: Import Export Excel Plugin
* Plugin URI: https://github.com/putusaputra/
* Description: Plugin for importing excel data to post and exporting post data to excel
* Version: 1.0
* Author: I Putu Saputra
* Author URI: https://github.com/putusaputra/
**/

require_once( plugin_dir_path(__FILE__) . "libraries/phpoffice_phpspreadsheet_1.8.2.0_require/vendor/autoload.php" );

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

// register assets
add_action('admin_enqueue_scripts', 'register_assets');
function register_assets() {
    wp_register_style('custom_style', plugins_url('/assets/css/style.css', __FILE__));
    wp_enqueue_style('custom_style');

    wp_register_script('custom_script', plugins_url('/assets/js/custom.js', __FILE__));
    wp_enqueue_script('custom_script');
}

// ajax declaration
add_action('wp_ajax_nopriv_import_excel', 'importExcel');
add_action('wp_ajax_import_excel', 'importExcel');
add_action('wp_ajax_nopriv_export_excel', 'exportExcel');
add_action('wp_ajax_export_excel', 'exportExcel');

function importExcel() {
    $file_mimes = array("application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    if(!isset($_FILES['import_from_excel']['name']))
    {
        $errors[] = 'Please upload correct excel file';
    }

    if(!in_array($_FILES['import_from_excel']['type'], $file_mimes))
    {
        $errors[] = 'Please upload correct excel file';
    }

    if(!sizeof($errors))
    {
        $arr_file = explode('.', $_FILES['import_from_excel']['name']);
        $extension = end($arr_file);
     
        if('xls' == $extension) {
            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
        } else {
            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        }
     
        $spreadsheet = $reader->load($_FILES['import_from_excel']['tmp_name']);
         
        $sheetData = $spreadsheet->getActiveSheet()->toArray();
        
        for ($i = 0;$i < count($sheetData); $i++) {
            
            if ($i == 0) {
                continue;
            }
            
            $post_title = $sheetData[$i][0];
            $post_content = $sheetData[$i][1];

            $newPost = [
                'post_title' => $post_title,
                'post_content' => $post_content,
                'post_type' => 'post',
                'post_status' => 'publish',
            ];

            wp_insert_post($newPost);
        }

        wp_send_json([
          'message' => 'Import excel to post success',
          'success' => true
        ]);
    }

    wp_send_json_error($errors);

    wp_die();
}

function exportExcel() {
    $post_type = sanitize_text_field($_POST['post_type']);
    $initposts = get_posts(array(
            'numberposts'   => -1,
            'post_type'     => $post_type
        ));
    $posts = (array) $initposts;
    
    $newPosts = array();
    foreach ($posts as $key => $post) {
        $newPosts[$key][0] = $post->post_title;
        $newPosts[$key][1] = $post->post_content;
    }

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    
    //prepare header
    $sheet->setCellValue('A1', 'Post Title');
    $sheet->setCellValue('B1', 'Post Content');

    foreach ($newPosts as $key => $newPost) {
        $sheet->setCellValue('A' . ($key + 2), $newPost[0]);
        $sheet->setCellValue('B' . ($key + 2), $newPost[1]);
    }

    $filename = "export-" . $post_type . "-" . date('Ymd') . ".xlsx";
    $location = plugin_dir_path(__FILE__) . 'download/' . date('Ymd');
    $location_client = plugins_url() . '/import_export_excel/download/' . date('Ymd');
    $file_location = $location . '/' . $filename;
    $file_location_client = $location_client . '/' . $filename;

    if (!file_exists($location)) {
        mkdir($location, 0777, true);
    }

    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($file_location);

    wp_send_json([
      'message' => 'Export post to excel success',
      'file'    => $file_location_client,
      'success' => true
    ]);
}

//front end
/*add import excel button to wordpress backend menu*/
add_action('restrict_manage_posts', 'add_import_excel_button');
function add_import_excel_button() {
    $screen = get_current_screen();
    if (isset($screen->parent_file) && ($screen->post_type === "post")) {
        ?>
        <script type="text/javascript">
            jQuery(function($) {
                var formElement = '<form enctype="multipart/form-data" id="excel-form-upload" name="excel-form-upload">'+
                                    '<label class="custom-file-upload" id="import_from_excel_wrapper" style="margin-right:5px;">'+
                                    '<input type="file" name="import_from_excel" id="import_from_excel" accept=".xls,.xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" />'+
                                        '<i class="fa fa-file-excel"></i> Browse Excel File'+
                                        '</label>'+
                                    '<button type="button" name="submit_import_excel" id="submit_import_excel"><span id="btnText">Import Excel File</span></button>'+
                                  '</form>';
                jQuery(formElement).insertBefore('.wp-header-end');
                
                jQuery("#submit_import_excel").on("click", function(e){
                    var file_data = jQuery('#import_from_excel').prop('files')[0];
                    var form_data = new FormData();
                    form_data.append('import_from_excel', file_data)
                    form_data.append('action', 'import_excel');

                    var pluginDir = '<?php echo plugins_url("assets/images/spinner.gif", __FILE__); ?>';
                    var spinner = '<img id="img-spinner" src="'+ pluginDir +'" width="30px" height="30px" style="vertical-align:bottom;">';
                    jQuery(this).prop('disabled', true);
                    jQuery(this).css('cursor', 'not-allowed');
                    jQuery(spinner).insertAfter('#submit_import_excel');

                    jQuery.ajax({
                        url: ajaxurl, 
                        type: "POST",
                        data: form_data,
                        contentType: false,
                        cache: false,
                        processData: false,
                        success: function(data){
                            var messages = data.message;
                            var successMessage = "";
                            var errorMessage = "";
                            if (data.success)
                            {   
                                var successMessage = "<span style='color:green;margin-left:5px;' class='notif-message'>"+messages+"</span>";
                                jQuery('#excel-form-upload').append(successMessage);
                                setTimeout(
                                    function() {
                                        jQuery('#excel-form-upload .notif-message').remove();
                                    }, 5000);
                                    
                                location.reload(true);
                            } else if (data.success === false) {
                                var error = data.data;

                                var i = 0;
                                for (i = 0; i < error.length; i++) {
                                    errorMessage += "<span style='color:red;margin-left:5px;' class='notif-message'>"+error[i]+"</span>";
                                }

                                jQuery('#excel-form-upload').append(errorMessage);
                            }

                            jQuery('#submit_import_excel').prop('disabled', false);
                            jQuery('#submit_import_excel').removeAttr('style');
                            jQuery('#img-spinner').remove();
                        },
                    })

                });
            });
                
        </script>
        <?php
    }
}

/*add export excel button to wordpress backend menu*/
add_action('restrict_manage_posts', 'add_export_excel_button');
function add_export_excel_button() {
    $screen = get_current_screen();
    if (isset($screen->parent_file) && ($screen->post_type === "post")) {
        ?>
        <script type="text/javascript">
            jQuery(function($) {
                var buttonElement = '<button type="button" name="submit_export_excel" id="submit_export_excel"><span>Export Post to Excel</span></button>';
                                  
                jQuery("#excel-form-upload").append(buttonElement);
                
                jQuery("#submit_export_excel").on("click", function(){
                    var file_data = '<?php echo $screen->post_type; ?>';
                    var form_data = new FormData();
                    form_data.append('post_type', file_data)
                    form_data.append('action', 'export_excel');

                    var pluginDir = '<?php echo plugins_url("assets/images/spinner.gif", __FILE__); ?>';
                    var spinner = '<img id="img-spinner" src="'+ pluginDir +'" width="30px" height="30px" style="vertical-align:bottom;">';
                    jQuery(this).prop('disabled', true);
                    jQuery(this).css('cursor', 'not-allowed');
                    jQuery('#excel-form-upload').append(spinner);

                    jQuery.ajax({
                        url: ajaxurl, 
                        type: "POST",
                        data: form_data,
                        contentType: false,
                        cache: false,
                        processData: false,
                        success: function(data){
                            var messages = data.message;
                            var successMessage = "";
                            var errorMessage = "";

                            jQuery('.notif-message').remove();

                            if (data.success)
                            {   
                                var successMessage = "<span style='color:green;margin-left:5px;' class='notif-message'>"+messages+"</span>";
                                jQuery('#excel-form-upload').append(successMessage);
                                window.location.href = data.file;
                            } else if (data.success === false) {
                                var error = data.data;

                                var i = 0;
                                for (i = 0; i < error.length; i++) {
                                    errorMessage += "<span style='color:red;margin-left:5px;' class='notif-message'>"+error[i]+"</span>";
                                }

                                jQuery('#excel-form-upload').append(errorMessage);
                            }

                            jQuery('#submit_export_excel').prop('disabled', false);
                            jQuery('#submit_export_excel').removeAttr('style');
                            jQuery('#img-spinner').remove();
                        },
                    })

                });
            });
                
        </script>
        <?php
    }
}