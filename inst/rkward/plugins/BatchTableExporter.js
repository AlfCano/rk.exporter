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

    echo('target_obj <- ' + obj + '\n');
    echo('if (!is.list(target_obj) || inherits(target_obj, "flextable")) { target_obj <- list(table_1 = target_obj) }\n');
    echo('if (is.null(names(target_obj))) { names(target_obj) <- paste0("table_", seq_along(target_obj)) }\n');
    echo('names(target_obj) <- ifelse(names(target_obj) == "", paste0("table_", seq_along(target_obj)), names(target_obj))\n');
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

        // LÓGICA MEJORADA: Un bucle iterativo para colocar saltos de página
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
            // PowerPoint automáticamente crea una diapositiva por tabla con save_as_pptx
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

