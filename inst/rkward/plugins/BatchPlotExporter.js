// this code was generated using the rkwarddev package.
// perhaps don't make changes here, but in the rkwarddev script instead!



function preprocess(is_preview){
	// add requirements etc. here
	echo("require(purrr)\n");	echo("require(ggplot2)\n");	echo("require(officer)\n");	echo("require(svglite)\n");	echo("require(rvg)\n");	echo("require(flextable)\n");
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
  
    var obj = getValue('plt_list'); var mode = getValue('plt_mode'); var out_dir = getValue('plt_dir').replace(/\\/g, '/'); var out_file = getValue('plt_file').replace(/\\/g, '/');
    var auto_ext = getValue('plt_auto_ext') == 'TRUE'; var ind_fmt = getValue('plt_ind_fmt'); var comb_fmt = getValue('plt_comb_fmt');
    var orient = getValue('plt_orient'); var w = getValue('plt_w'); var h = getValue('plt_h'); var dpi = getValue('plt_dpi');
    var strat = getValue('plt_name_strategy'); var prefix = getValue('plt_prefix'); var dict_df = getValue('plt_dict_df');
    var dict_key = getColName(getValue('plt_dict_key')); var dict_val = getColName(getValue('plt_dict_val'));

    echo('target_obj <- ' + obj + '\n');

    // ESTRATEGIA 1: Aceptar tanto ggplot como recordedplot
    echo('if (!is.list(target_obj) || inherits(target_obj, "ggplot") || inherits(target_obj, "recordedplot")) {\n');
    echo('  target_obj <- list(target_obj)\n');
    echo('  names(target_obj) <- "' + obj + '"\n');
    echo('}\n');

    if (strat == 'seq') {
        echo('names(target_obj) <- paste0("' + prefix + '", seq_along(target_obj))\n');
    } else {
        echo('if (is.null(names(target_obj))) names(target_obj) <- paste0("' + obj + '_", seq_along(target_obj))\n');
        echo('empty_idx <- trimws(names(target_obj)) == ""\n');
        echo('if (any(empty_idx)) names(target_obj)[empty_idx] <- paste0("' + obj + '_", which(empty_idx))\n');

        if (strat == 'dict') {
            echo('\n# Dictionary Mapping\n');
            echo('dict_data <- as.data.frame(' + dict_df + ')\n');
            echo('key_col <- trimws(as.character(dict_data[["' + dict_key + '"]]))\n');
            echo('val_col <- trimws(as.character(dict_data[["' + dict_val + '"]]))\n');
            echo('matched_idx <- match(trimws(names(target_obj)), key_col)\n');
            echo('valid_matches <- !is.na(matched_idx)\n');
            echo('names(target_obj)[valid_matches] <- val_col[matched_idx[valid_matches]]\n');
        }
    }
    echo('names(target_obj) <- gsub("[^A-Za-z0-9_.-]", "_", names(target_obj))\n\n');

    if (mode == 'ind') {
        echo('if ("' + out_dir + '" == "") stop("Error: Output Directory is required.")\n');
        echo('require(purrr)\nrequire(svglite)\n');
        echo('dir.create("' + out_dir + '", showWarnings = FALSE, recursive = TRUE)\n');
        echo('purrr::iwalk(target_obj, function(.x, .y) {\n');
        echo('  ruta <- file.path("' + out_dir + '", paste0(.y, ".' + ind_fmt + '"))\n');

        // ESTRATEGIA 2: Renderizado Universal con dispositivos Base R
        if (ind_fmt == 'rds') {
            echo('  saveRDS(.x, file = ruta)\n');
        } else if (ind_fmt == 'png') {
            echo('  png(filename = ruta, width = ' + w + ', height = ' + h + ', units = "in", res = ' + dpi + ')\n');
            echo('  print(.x)\n  dev.off()\n');
        } else if (ind_fmt == 'svg') {
            echo('  svglite::svglite(filename = ruta, width = ' + w + ', height = ' + h + ')\n');
            echo('  print(.x)\n  dev.off()\n');
        } else if (ind_fmt == 'pdf') {
            echo('  pdf(file = ruta, width = ' + w + ', height = ' + h + ')\n');
            echo('  print(.x)\n  dev.off()\n');
        }

        echo('})\n');
        echo('res_msg <- paste(length(target_obj), "plots successfully exported to:", "' + out_dir + '")\n');

    } else {
        echo('if ("' + out_file + '" == "") stop("Error: Output File is required.")\n');
        echo('out_file <- "' + out_file + '"\n');
        if (auto_ext) {
            echo('ext_pattern <- paste0("\\\\.", "' + comb_fmt + '", "$")\n');
            echo('if (!grepl(ext_pattern, out_file, ignore.case = TRUE)) out_file <- paste0(out_file, ".", "' + comb_fmt + '")\n');
        }

        if (comb_fmt == 'rds') {
            echo('saveRDS(target_obj, file = out_file)\n');
            echo('res_msg <- paste("Combined list saved as RDS to:", out_file)\n');

        } else if (comb_fmt == 'pdf') {
            echo('if ("' + orient + '" == "landscape") { pdf(out_file, width = ' + w + ', height = ' + h + ') } else { pdf(out_file, width = ' + h + ', height = ' + w + ') }\n');
            echo('purrr::walk(target_obj, print)\n');
            echo('dev.off()\n');
            echo('res_msg <- paste("Combined PDF exported to:", out_file)\n');

        } else if (comb_fmt == 'docx') {
            echo('require(officer)\n');
            echo('doc <- officer::read_docx()\n');

            // ESTRATEGIA 3: Clausura segura para body_add_plot en Word
            echo('safe_plot_fn <- function(p) { force(p); function() print(p) }\n');

            echo('for (i in seq_along(target_obj)) {\n');
            echo('  doc <- officer::body_add_plot(doc, value = safe_plot_fn(target_obj[[i]]), width = ' + w + ', height = ' + h + ')\n');
            echo('  if (i < length(target_obj)) doc <- officer::body_add_break(doc)\n');
            echo('}\n');
            echo('sect_prop <- officer::prop_section(page_size = officer::page_size(orient = "' + orient + '"))\n');
            echo('doc <- officer::body_set_default_section(doc, sect_prop)\n');
            echo('print(doc, target = out_file)\n');
            echo('res_msg <- paste("Combined Word exported to:", out_file)\n');

        } else if (comb_fmt == 'pptx') {
            echo('require(officer)\nrequire(rvg)\n');
            echo('doc <- officer::read_pptx()\n');
            echo('for (i in seq_along(target_obj)) {\n');
            echo('  doc <- officer::add_slide(doc, layout = "Title and Content", master = "Office Theme")\n');
            echo('  doc <- officer::ph_with(doc, value = names(target_obj)[i], location = officer::ph_location_type(type = "title"))\n');

            // ESTRATEGIA 4: Uso de rvg::dml para PPTX mixto
            echo('  if (inherits(target_obj[[i]], "ggplot")) {\n');
            echo('    doc <- officer::ph_with(doc, value = target_obj[[i]], location = officer::ph_location_type(type = "body"))\n');
            echo('  } else {\n');
            echo('    doc <- officer::ph_with(doc, value = rvg::dml(code = print(target_obj[[i]])), location = officer::ph_location_type(type = "body"))\n');
            echo('  }\n');

            echo('}\n');
            echo('print(doc, target = out_file)\n');
            echo('res_msg <- paste("Combined PowerPoint exported to:", out_file)\n');
        }
    }
  
}

function printout(is_preview){
	// printout the results
	new Header(i18n("Batch Plot Exporter results")).print();

    echo('rk.header("Batch Plot Export Results", level=2)\n');
    echo('rk.print(res_msg)\n');
  

}

