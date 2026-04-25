local({
  # =========================================================================================
  # 1. Metadatos y Configuración
  # =========================================================================================
  require(rkwarddev)
  rkwarddev.required("0.10-3")

  package_about <- rk.XML.about(
    name = "rk.exporter",
    author = person(
      given = "Alfonso",
      family = "Cano Robles",
      email = "alfonso.cano@correo.buap.mx",
      role = c("aut", "cre")
    ),
    about = list(
      desc = "RKWard Plugin Suite for Batch Exporting lists of ggplot2 and flextable objects to individual files or combined Office documents (PDF, Word, PPTX).",
      version = "0.0.1",
      url = "https://github.com/AlfCano/rk.exporter",
      license = "GPL (>= 3)"
    )
  )

  common_hierarchy <- list("file", "Batch Exporters")

  js_parse_helper <- "
    function getColName(fullPath) {
        if (!fullPath) return '';
        if (fullPath.indexOf('$') > -1) {
            return fullPath.split('$')[1];
        } else if (fullPath.indexOf('[[') > -1) {
            var inner = fullPath.split('[[')[1].replace(']]', '');
            return inner.split('\"').join('').split(String.fromCharCode(39)).join('');
        }
        return fullPath;
    }
  "

  # =========================================================================================
  # COMPONENTE 1: Batch Plot Exporter (ggplot2)
  # =========================================================================================
  help_plt <- rk.rkh.doc(title = rk.rkh.title("Batch Plot Exporter"), summary = rk.rkh.summary("Exporta una lista de gráficos (ggplot2) a archivos individuales o a un documento combinado (PDF, Word, PPTX)."))

  var_sel_plt <- rk.XML.varselector(id.name = "var_sel_plt")
  plt_list <- rk.XML.varslot("Target object (List of ggplot2 objects or single ggplot)", source = var_sel_plt, required = TRUE, id.name = "plt_list")

  plt_mode <- rk.XML.radio("Export Mode", options = list(
    "Individual Files (e.g. plot_1.svg, plot_2.svg)" = list(val = "ind", chk = TRUE),
    "Combined Document (1 plot per page/slide)" = list(val = "comb")
  ), id.name = "plt_mode")

  plt_dir <- rk.XML.browser("Output Directory (For Individual Files)", type = "dir", required = FALSE, id.name = "plt_dir")
  plt_file <- rk.XML.browser("Output File (For Combined Document, e.g. plots.pptx)", type = "savefile", required = FALSE, id.name = "plt_file")
  plt_auto_ext <- rk.XML.cbox("Automatically append extension to file name if missing", value = "TRUE", chk = TRUE, id.name = "plt_auto_ext")

  plt_ind_fmt <- rk.XML.dropdown("Format for Individual Files", options = list("SVG" = list(val = "svg", chk=TRUE), "PNG" = list(val = "png"), "PDF" = list(val = "pdf")), id.name = "plt_ind_fmt")
  plt_comb_fmt <- rk.XML.dropdown("Format for Combined Document", options = list("PowerPoint (.pptx)" = list(val = "pptx", chk=TRUE), "Word (.docx)" = list(val = "docx"), "PDF (.pdf)" = list(val = "pdf")), id.name = "plt_comb_fmt")

  plt_w <- rk.XML.spinbox("Width (inches)", min = 1, max = 30, initial = 10, id.name = "plt_w")
  plt_h <- rk.XML.spinbox("Height (inches)", min = 1, max = 30, initial = 6, id.name = "plt_h")
  plt_dpi <- rk.XML.spinbox("Resolution (DPI) (For PNGs)", min = 50, max = 600, initial = 300, id.name = "plt_dpi")
  plt_orient <- rk.XML.dropdown("Page Orientation (For Word/PDF)", options = list("Landscape" = list(val = "landscape", chk=TRUE), "Portrait" = list(val = "portrait")), id.name = "plt_orient")

  tab_plt_in <- rk.XML.row(var_sel_plt, rk.XML.col(plt_list, plt_mode, rk.XML.frame(plt_dir, plt_file, plt_auto_ext, label="Destination Paths")))
  tab_plt_opt <- rk.XML.col(rk.XML.frame(plt_ind_fmt, plt_comb_fmt, plt_orient, label="Format Settings"), rk.XML.frame(rk.XML.row(plt_w, plt_h), plt_dpi, label="Dimensions"))

  dialog_plt <- rk.XML.dialog(label = "Batch Plot Exporter", child = rk.XML.tabbook(tabs = list("Input & Mode" = tab_plt_in, "Export Settings" = tab_plt_opt)))

  is_ind_plt <- rk.XML.convert(sources = "plt_mode.string", mode = c(equals = "ind"), id.name = "is_ind_plt")
  is_comb_plt <- rk.XML.convert(sources = "plt_mode.string", mode = c(equals = "comb"), id.name = "is_comb_plt")

  logic_plt <- rk.XML.logic(
      is_ind_plt,
      is_comb_plt,
      rk.XML.connect(governor = is_ind_plt, client = "plt_dir.enabled"),
      rk.XML.connect(governor = is_ind_plt, client = "plt_dir.required"),
      rk.XML.connect(governor = is_ind_plt, client = "plt_ind_fmt.enabled"),

      rk.XML.connect(governor = is_comb_plt, client = "plt_file.enabled"),
      rk.XML.connect(governor = is_comb_plt, client = "plt_file.required"),
      rk.XML.connect(governor = is_comb_plt, client = "plt_auto_ext.enabled"),
      rk.XML.connect(governor = is_comb_plt, client = "plt_comb_fmt.enabled"),
      rk.XML.connect(governor = is_comb_plt, client = "plt_orient.enabled")
  )

  js_calc_plt <- paste0(js_parse_helper, "
    var obj = getValue('plt_list');
    var mode = getValue('plt_mode');
    var out_dir = getValue('plt_dir').replace(/\\\\/g, '/');
    var out_file = getValue('plt_file').replace(/\\\\/g, '/');
    var auto_ext = getValue('plt_auto_ext') == 'TRUE';
    var ind_fmt = getValue('plt_ind_fmt');
    var comb_fmt = getValue('plt_comb_fmt');
    var orient = getValue('plt_orient');
    var w = getValue('plt_w');
    var h = getValue('plt_h');
    var dpi = getValue('plt_dpi');

    echo('target_obj <- ' + obj + '\\n');
    echo('if (!is.list(target_obj) || inherits(target_obj, \"ggplot\")) { target_obj <- list(plot_1 = target_obj) }\\n');
    echo('if (is.null(names(target_obj))) { names(target_obj) <- paste0(\"plot_\", seq_along(target_obj)) }\\n');
    echo('names(target_obj) <- ifelse(names(target_obj) == \"\", paste0(\"plot_\", seq_along(target_obj)), names(target_obj))\\n');
    echo('names(target_obj) <- gsub(\"[^A-Za-z0-9_.-]\", \"_\", names(target_obj))\\n\\n');

    if (mode == 'ind') {
        echo('if (\"' + out_dir + '\" == \"\") stop(\"Error: Output Directory is required.\")\\n');
        echo('require(purrr)\\nrequire(ggplot2)\\nrequire(svglite)\\n');
        echo('dir.create(\"' + out_dir + '\", showWarnings = FALSE, recursive = TRUE)\\n');
        echo('purrr::iwalk(target_obj, function(.x, .y) {\\n');
        echo('  ruta <- file.path(\"' + out_dir + '\", paste0(.y, \".' + ind_fmt + '\"))\\n');
        echo('  ggplot2::ggsave(filename = ruta, plot = .x, device = \"' + ind_fmt + '\", width = ' + w + ', height = ' + h + ', dpi = ' + dpi + ')\\n');
        echo('})\\n');
        echo('res_msg <- paste(length(target_obj), \"plots successfully exported to:\", \"' + out_dir + '\")\\n');
    } else {
        echo('if (\"' + out_file + '\" == \"\") stop(\"Error: Output File is required.\")\\n');

        echo('out_file <- \"' + out_file + '\"\\n');
        if (auto_ext) {
            echo('ext_pattern <- paste0(\"\\\\\\\\.\", \"' + comb_fmt + '\", \"$\")\\n');
            echo('if (!grepl(ext_pattern, out_file, ignore.case = TRUE)) out_file <- paste0(out_file, \".\", \"' + comb_fmt + '\")\\n');
        }

        if (comb_fmt == 'pdf') {
            echo('if (\"' + orient + '\" == \"landscape\") { pdf(out_file, width = ' + w + ', height = ' + h + ') } else { pdf(out_file, width = ' + h + ', height = ' + w + ') }\\n');
            echo('purrr::walk(target_obj, print)\\n');
            echo('dev.off()\\n');
            echo('res_msg <- paste(\"Combined PDF exported to:\", out_file)\\n');
        } else if (comb_fmt == 'docx') {
            echo('require(officer)\\n');
            echo('doc <- officer::read_docx()\\n');
            echo('for (i in seq_along(target_obj)) {\\n');
            echo('  doc <- officer::body_add_gg(doc, value = target_obj[[i]], width = ' + w + ', height = ' + h + ')\\n');
            echo('  if (i < length(target_obj)) doc <- officer::body_add_break(doc)\\n');
            echo('}\\n');
            echo('sect_prop <- officer::prop_section(page_size = officer::page_size(orient = \"' + orient + '\"))\\n');
            echo('doc <- officer::body_set_default_section(doc, sect_prop)\\n');
            echo('print(doc, target = out_file)\\n');
            echo('res_msg <- paste(\"Combined Word exported to:\", out_file)\\n');
        } else if (comb_fmt == 'pptx') {
            echo('require(officer)\\n');
            echo('doc <- officer::read_pptx()\\n');
            echo('for (i in seq_along(target_obj)) {\\n');
            echo('  doc <- officer::add_slide(doc, layout = \"Title and Content\", master = \"Office Theme\")\\n');
            echo('  doc <- officer::ph_with(doc, value = names(target_obj)[i], location = officer::ph_location_type(type = \"title\"))\\n');
            echo('  doc <- officer::ph_with(doc, value = target_obj[[i]], location = officer::ph_location_type(type = \"body\"))\\n');
            echo('}\\n');
            echo('print(doc, target = out_file)\\n');
            echo('res_msg <- paste(\"Combined PowerPoint exported to:\", out_file)\\n');
        }
    }
  ")

  js_print_plt <- "
    echo('rk.header(\"Batch Plot Export Results\", level=2)\\n');
    echo('rk.print(res_msg)\\n');
  "

  comp_plt <- rk.plugin.component("Batch Plot Exporter", xml = list(dialog = dialog_plt, logic = logic_plt), js = list(require = c("purrr", "ggplot2", "officer", "svglite"), calculate = js_calc_plt, printout = js_print_plt), hierarchy = common_hierarchy, rkh = list(help = help_plt))


  # =========================================================================================
  # COMPONENTE 2: Batch Table Exporter (flextable)
  # =========================================================================================
  help_tbl <- rk.rkh.doc(title = rk.rkh.title("Batch Table Exporter"), summary = rk.rkh.summary("Exporta una lista de tablas (flextable) a archivos individuales o a un documento combinado (Word, PPTX)."))

  var_sel_tbl <- rk.XML.varselector(id.name = "var_sel_tbl")
  tbl_list <- rk.XML.varslot("Target object (List of flextables or single flextable)", source = var_sel_tbl, required = TRUE, id.name = "tbl_list")

  tbl_mode <- rk.XML.radio("Export Mode", options = list(
    "Individual Files (e.g. table_1.docx, table_2.docx)" = list(val = "ind", chk = TRUE),
    "Combined Document (1 table per page/slide)" = list(val = "comb")
  ), id.name = "tbl_mode")

  tbl_dir <- rk.XML.browser("Output Directory (For Individual Files)", type = "dir", required = FALSE, id.name = "tbl_dir")
  tbl_file <- rk.XML.browser("Output File (For Combined Document, e.g. tables.docx)", type = "savefile", required = FALSE, id.name = "tbl_file")
  tbl_auto_ext <- rk.XML.cbox("Automatically append extension to file name if missing", value = "TRUE", chk = TRUE, id.name = "tbl_auto_ext")

  tbl_ind_fmt <- rk.XML.dropdown("Format for Individual Files", options = list("Word (.docx)" = list(val = "docx", chk=TRUE), "PowerPoint (.pptx)" = list(val = "pptx"), "HTML (.html)" = list(val = "html")), id.name = "tbl_ind_fmt")
  tbl_comb_fmt <- rk.XML.dropdown("Format for Combined Document", options = list("Word (.docx)" = list(val = "docx", chk=TRUE), "PowerPoint (.pptx)" = list(val = "pptx")), id.name = "tbl_comb_fmt")
  tbl_orient <- rk.XML.dropdown("Word Page Orientation", options = list("Landscape" = list(val = "landscape", chk=TRUE), "Portrait" = list(val = "portrait")), id.name = "tbl_orient")

  tab_tbl_in <- rk.XML.row(var_sel_tbl, rk.XML.col(tbl_list, tbl_mode, rk.XML.frame(tbl_dir, tbl_file, tbl_auto_ext, label="Destination Paths")))
  tab_tbl_opt <- rk.XML.col(rk.XML.frame(tbl_ind_fmt, tbl_comb_fmt, tbl_orient, label="Format Settings"))

  dialog_tbl <- rk.XML.dialog(label = "Batch Table Exporter", child = rk.XML.tabbook(tabs = list("Input & Mode" = tab_tbl_in, "Export Settings" = tab_tbl_opt)))

  is_ind_tbl <- rk.XML.convert(sources = "tbl_mode.string", mode = c(equals = "ind"), id.name = "is_ind_tbl")
  is_comb_tbl <- rk.XML.convert(sources = "tbl_mode.string", mode = c(equals = "comb"), id.name = "is_comb_tbl")

  logic_tbl <- rk.XML.logic(
      is_ind_tbl,
      is_comb_tbl,
      rk.XML.connect(governor = is_ind_tbl, client = "tbl_dir.enabled"),
      rk.XML.connect(governor = is_ind_tbl, client = "tbl_dir.required"),
      rk.XML.connect(governor = is_ind_tbl, client = "tbl_ind_fmt.enabled"),

      rk.XML.connect(governor = is_comb_tbl, client = "tbl_file.enabled"),
      rk.XML.connect(governor = is_comb_tbl, client = "tbl_file.required"),
      rk.XML.connect(governor = is_comb_tbl, client = "tbl_auto_ext.enabled"),
      rk.XML.connect(governor = is_comb_tbl, client = "tbl_comb_fmt.enabled"),
      rk.XML.connect(governor = is_comb_tbl, client = "tbl_orient.enabled")
  )

  js_calc_tbl <- paste0(js_parse_helper, "
    var obj = getValue('tbl_list');
    var mode = getValue('tbl_mode');
    var out_dir = getValue('tbl_dir').replace(/\\\\/g, '/');
    var out_file = getValue('tbl_file').replace(/\\\\/g, '/');
    var auto_ext = getValue('tbl_auto_ext') == 'TRUE';
    var orient = getValue('tbl_orient');
    var fmt = (mode == 'ind') ? getValue('tbl_ind_fmt') : getValue('tbl_comb_fmt');

    echo('target_obj <- ' + obj + '\\n');
    echo('if (!is.list(target_obj) || inherits(target_obj, \"flextable\")) { target_obj <- list(table_1 = target_obj) }\\n');
    echo('if (is.null(names(target_obj))) { names(target_obj) <- paste0(\"table_\", seq_along(target_obj)) }\\n');
    echo('names(target_obj) <- ifelse(names(target_obj) == \"\", paste0(\"table_\", seq_along(target_obj)), names(target_obj))\\n');
    echo('names(target_obj) <- gsub(\"[^A-Za-z0-9_.-]\", \"_\", names(target_obj))\\n\\n');

    if (mode == 'ind') {
        echo('if (\"' + out_dir + '\" == \"\") stop(\"Error: Output Directory is required.\")\\n');
        echo('require(purrr)\\nrequire(flextable)\\nrequire(officer)\\n');
        echo('dir.create(\"' + out_dir + '\", showWarnings = FALSE, recursive = TRUE)\\n');

        echo('purrr::iwalk(target_obj, function(.x, .y) {\\n');
        echo('  ruta <- file.path(\"' + out_dir + '\", paste0(.y, \".' + fmt + '\"))\\n');
        if (fmt == 'docx') {
            echo('  sect_prop <- officer::prop_section(page_size = officer::page_size(orient = \"' + orient + '\"))\\n');
            echo('  flextable::save_as_docx(.x, path = ruta, pr_section = sect_prop)\\n');
        } else if (fmt == 'pptx') {
            echo('  flextable::save_as_pptx(.x, path = ruta)\\n');
        } else if (fmt == 'html') {
            echo('  flextable::save_as_html(.x, path = ruta)\\n');
        }
        echo('})\\n');
        echo('res_msg <- paste(length(target_obj), \"tables successfully exported to:\", \"' + out_dir + '\")\\n');
    } else {
        echo('if (\"' + out_file + '\" == \"\") stop(\"Error: Output File is required.\")\\n');
        echo('require(flextable)\\nrequire(officer)\\n');

        echo('out_file <- \"' + out_file + '\"\\n');
        if (auto_ext) {
            echo('ext_pattern <- paste0(\"\\\\\\\\.\", \"' + fmt + '\", \"$\")\\n');
            echo('if (!grepl(ext_pattern, out_file, ignore.case = TRUE)) out_file <- paste0(out_file, \".\", \"' + fmt + '\")\\n');
        }

        // LÓGICA MEJORADA: Un bucle iterativo para colocar saltos de página
        if (fmt == 'docx') {
            echo('doc <- officer::read_docx()\\n');
            echo('for (i in seq_along(target_obj)) {\\n');
            echo('  doc <- flextable::body_add_flextable(doc, value = target_obj[[i]])\\n');
            echo('  if (i < length(target_obj)) doc <- officer::body_add_break(doc)\\n');
            echo('}\\n');
            echo('sect_prop <- officer::prop_section(page_size = officer::page_size(orient = \"' + orient + '\"))\\n');
            echo('doc <- officer::body_set_default_section(doc, sect_prop)\\n');
            echo('print(doc, target = out_file)\\n');
            echo('res_msg <- paste(\"Combined Word exported to:\", out_file)\\n');

        } else if (fmt == 'pptx') {
            // PowerPoint automáticamente crea una diapositiva por tabla con save_as_pptx
            echo('do.call(flextable::save_as_pptx, c(target_obj, list(path = out_file)))\\n');
            echo('res_msg <- paste(\"Combined PowerPoint exported to:\", out_file)\\n');
        }
    }
  ")

  js_print_tbl <- "
    echo('rk.header(\"Batch Table Export Results\", level=2)\\n');
    echo('rk.print(res_msg)\\n');
  "

  comp_tbl <- rk.plugin.component("Batch Table Exporter", xml = list(dialog = dialog_tbl, logic = logic_tbl), js = list(require = c("purrr", "flextable", "officer"), calculate = js_calc_tbl, printout = js_print_tbl), hierarchy = common_hierarchy, rkh = list(help = help_tbl))

  # =========================================================================================
  # CONSTRUCCIÓN DEL ESQUELETO
  # =========================================================================================

  rk.plugin.skeleton(
    about = package_about,
    path = ".",
    xml = list(dialog = dialog_plt, logic = logic_plt),
    js = list(require = c("purrr", "ggplot2", "officer", "svglite"), calculate = js_calc_plt, printout = js_print_plt),
    rkh = list(help = help_plt),
    components = list(comp_tbl),
    pluginmap = list(name = "Batch Plot Exporter", hierarchy = common_hierarchy),
    create = c("pmap", "xml", "js", "desc", "rkh"),
    load = TRUE,
    overwrite = TRUE,
    show = FALSE
  )

  cat("\nPlugin 'rk.exporter' generado con éxito.\n")
})
