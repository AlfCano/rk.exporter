// this code was generated using the rkwarddev package.
// perhaps don't make changes here, but in the rkwarddev script instead!



function preprocess(is_preview){
	// add requirements etc. here
	echo("require(purrr)\n");	echo("require(flextable)\n");	echo("require(officer)\n");
}

function calculate(is_preview){
	// read in variables from dialog


	// the R code to be evaluated

    function getColName(fullPath) {
        if (!fullPath) return '';
        if (fullPath.indexOf('$') > -1) {
            return fullPath.split('$')[1];
        } else if (fullPath.indexOf('[[') > -1) {
            var inner = fullPath.split('[[')[1].replace(']]', '');
            return inner.split('"').join('').split(String.fromCharCode(39)).join('');
        }
        return fullPath;
    }
  
    var obj = getValue('tbl_list');
    var mode = getValue('tbl_mode');
    var out_dir = getValue('tbl_dir').replace(/\\/g, '/');
    var out_file = getValue('tbl_file').replace(/\\/g, '/');
    var auto_ext = getValue('tbl_auto_ext') == 'TRUE';
    var orient = getValue('tbl_orient');
    var fmt = (mode == 'ind') ? getValue('tbl_ind_fmt') : getValue('tbl_comb_fmt');

    // NEW JS VARIABLES
    var prefix = getValue('tbl_prefix');
    var use_dict = getValue('tbl_use_dict');
    var dict_df = getValue('tbl_dict_df');
    var dict_key = getColName(getValue('tbl_dict_key'));
    var dict_val = getColName(getValue('tbl_dict_val'));

    echo('target_obj <- ' + obj + '\n');

    // FEATURE 1: Automatic naming based on Workspace
    echo('if (!is.list(target_obj) || inherits(target_obj, "flextable")) {\n');
    echo('  target_obj <- list(target_obj)\n');
    echo('  names(target_obj) <- "' + obj + '"\n'); // Extracts the literal string from GUI
    echo('}\n');

    // FEATURE 3: Custom prefix
    echo('if (is.null(names(target_obj))) { names(target_obj) <- paste0("' + prefix + '", seq_along(target_obj)) }\n');
    echo('names(target_obj) <- ifelse(trimws(names(target_obj)) == "", paste0("' + prefix + '", seq_along(target_obj)), names(target_obj))\n');

    // FEATURE 2: Dictionary Mapping
    if (use_dict == 'true') {
        echo('\n# Dictionary Mapping\n');
        echo('dict_data <- as.data.frame(' + dict_df + ')\n');
        echo('key_col <- trimws(as.character(dict_data[["' + dict_key + '"]]))\n');
        echo('val_col <- trimws(as.character(dict_data[["' + dict_val + '"]]))\n');
        echo('matched_idx <- match(trimws(names(target_obj)), key_col)\n');
        echo('valid_matches <- !is.na(matched_idx)\n');
        echo('names(target_obj)[valid_matches] <- val_col[matched_idx[valid_matches]]\n');
    }

    echo('names(target_obj) <- gsub("[^A-Za-z0-9_.-]", "_", names(target_obj))\n\n');

    if (mode == 'ind') {
        echo('if ("' + out_dir + '" == "") stop("Error: Output Directory is required.")\n');
        echo('require(purrr)\nrequire(flextable)\nrequire(officer)\n');
        echo('dir.create("' + out_dir + '", showWarnings = FALSE, recursive = TRUE)\n');

        echo('purrr::iwalk(target_obj, function(.x, .y) {\n');
        echo('  ruta <- file.path("' + out_dir + '", paste0(.y, ".' + fmt + '"))\n');
        if (fmt == 'docx') {
            echo('  sect_prop <- officer::prop_section(page_size = officer::page_size(orient = "' + orient + '"))\n');
            echo('  flextable::save_as_docx(.x, path = ruta, pr_section = sect_prop)\n');
        } else if (fmt == 'pptx') {
            echo('  flextable::save_as_pptx(.x, path = ruta)\n');
        } else if (fmt == 'html') {
            echo('  flextable::save_as_html(.x, path = ruta)\n');
        }
        echo('})\n');
        echo('res_msg <- paste(length(target_obj), "tables successfully exported to:", "' + out_dir + '")\n');
    } else {
        echo('if ("' + out_file + '" == "") stop("Error: Output File is required.")\n');
        echo('require(flextable)\nrequire(officer)\n');

        echo('out_file <- "' + out_file + '"\n');
        if (auto_ext) {
            echo('ext_pattern <- paste0("\\\\.", "' + fmt + '", "$")\n');
            echo('if (!grepl(ext_pattern, out_file, ignore.case = TRUE)) out_file <- paste0(out_file, ".", "' + fmt + '")\n');
        }

        // IMPROVED LOGIC: Iterative loop to add page breaks
        if (fmt == 'docx') {
            echo('doc <- officer::read_docx()\n');
            echo('for (i in seq_along(target_obj)) {\n');
            echo('  doc <- flextable::body_add_flextable(doc, value = target_obj[[i]])\n');
            echo('  if (i < length(target_obj)) doc <- officer::body_add_break(doc)\n');
            echo('}\n');
            echo('sect_prop <- officer::prop_section(page_size = officer::page_size(orient = "' + orient + '"))\n');
            echo('doc <- officer::body_set_default_section(doc, sect_prop)\n');
            echo('print(doc, target = out_file)\n');
            echo('res_msg <- paste("Combined Word exported to:", out_file)\n');

        } else if (fmt == 'pptx') {
            // PowerPoint automatically creates one slide per table with save_as_pptx
            echo('do.call(flextable::save_as_pptx, c(target_obj, list(path = out_file)))\n');
            echo('res_msg <- paste("Combined PowerPoint exported to:", out_file)\n');
        }
    }
  
}

function printout(is_preview){
	// printout the results
	new Header(i18n("Batch Table Exporter results")).print();

    echo('rk.header("Batch Table Export Results", level=2)\n');
    echo('rk.print(res_msg)\n');
  

}

